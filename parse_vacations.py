# -*- coding: utf-8 -*-
"""Parse official ZUP T-7 and per-store vacation files into data.js."""
from __future__ import annotations

import json
import re
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
import xlrd

ROOT = Path(__file__).resolve().parent
STORE_DIR = ROOT / "Графики официальные"
OFFICIAL = ROOT / "график отпусков ЗУП ИП Пафнутьева ЕП. (офиц).xlsx"
STAFF_FILE = ROOT / "данные о работниках на 14.09.2026.mxl"
HIRE_FILE = ROOT / "даты приема.mxl"
OUT_JS = ROOT / "data.js"

HEADER_SKIP = {
    "структурное подразделение",
    "должность (специальность, профессия) по штатному расписанию",
    "фамилия, имя, отчество",
    "фамилия, имя, отчество сотрудника",
    "количество календарных дней",
    "запланированная",
    "1",
    "2",
    "3",
}

RETAIL_STORES = {
    "Дружба",
    "Авторемонтная",
    "Республики",
    "МСК 120А",
    "Бабарынка",
    "Сантехника",
    "Ожогина",
    "Червишево",
    "Антипино",
    "Березняки",
    "Чайка",
    "Винзили",
    "Набережная",
    "МСК 6 км",
    "Луговое",
    "Решетникова",
    "Тюнево",
    "Перевалово",
    "Успенка",
    "Ворошилова",
    "Мальково",
    "Щорса",
    "Ембаево",
    "Заводоуковск",
}

STORE_SHORT = [
    ("сантехника", "Сантехника"),
    ("дружбы 66", "Дружба"),
    ("авторемонтная", "Авторемонтная"),
    ("республики", "Республики"),
    ("московскийтракт 120", "МСК 120А"),
    ("московский тракт 120", "МСК 120А"),
    ("бабарынка", "Бабарынка"),
    ("ожогин", "Ожогина"),
    ("червишево", "Червишево"),
    ("старый тобольский", "Антипино"),
    ("березняков", "Березняки"),
    ("чайка", "Чайка"),
    ("винзили", "Винзили"),
    ("набережная", "Набережная"),
    ("московский тракт 6", "МСК 6 км"),
    ("луговое", "Луговое"),
    ("решетников", "Решетникова"),
    ("тюнево", "Тюнево"),
    ("перевалово", "Перевалово"),
    ("успенка", "Успенка"),
    ("ворошилова", "Ворошилова"),
    ("мальково", "Мальково"),
    ("щорса", "Щорса"),
    ("ембаево", "Ембаево"),
    ("заводоуковск", "Заводоуковск"),
    ("черемухов", "Заводоуковск"),
    ("складской и транспортной", "РЦ / логистика"),
    ("отдел закупок", "Закупки"),
    ("корпоративных продаж", "B2B"),
    ("электронной коммерции", "E-com"),
    ("розничной торговли", "ОРТ"),
    ("информационных технологий", "IT"),
    ("заботы о клиентах", "СЗК"),
    ("бизнес-процессам", "СУБП"),
    ("ревизионный", "Ревизия"),
    ("развития розничной", "Развитие сети"),
    ("ауп", "АУП"),
]

FILE_STORE = [
    ("авто", "Авторемонтная"),
    ("мальково", "Мальково"),
    ("щорса", "Щорса"),
    ("антипино", "Антипино"),
    ("березняки", "Березняки"),
    ("решетникова", "Решетникова"),
    ("мск6", "МСК 6 км"),
    ("винзили", "Винзили"),
    ("набережная", "Набережная"),
    ("луговое", "Луговое"),
    ("респа", "Республики"),
    ("червишево", "Червишево"),
    ("бабарынка", "Бабарынка"),
    ("перевалово", "Перевалово"),
    ("успенка", "Успенка"),
    ("дружба", "Дружба"),
    ("ожогино", "Ожогина"),
    ("ожогина", "Ожогина"),
    ("тюнево", "Тюнево"),
    ("сантехника", "Сантехника"),
    ("чайка", "Чайка"),
    ("заводоуковск", "Заводоуковск"),
]


def norm(s: str) -> str:
    s = (s or "").replace("\xa0", " ").replace("ё", "е").replace("Ё", "Е")
    s = s.lower().strip()
    s = re.sub(r"\s+", " ", s)
    return s


def store_short(name: str) -> str:
    n = norm(name)
    for needle, short in STORE_SHORT:
        if needle in n:
            return short
    if n.startswith("магазин"):
        return re.sub(r"\s+", " ", name).strip()[:40]
    return re.sub(r"\s+", " ", name or "").strip() or "Не указан"


def file_store(filename: str) -> str:
    n = norm(filename)
    for needle, short in FILE_STORE:
        if needle in n:
            return short
    return filename


REPORT_ROLES = ("Кладовщик", "СПК", "ПК", "Кассир", "РТЗ", "Менеджер")

# Short / maiden / abbreviated names → current full FIO
FIO_CANONICAL = {
    "котков са": "Котков Сергей Андреевич",
    "котков с а": "Котков Сергей Андреевич",
    "белоногов юв": "Белоногов Юрий Владимирович",
    "белоногов ю в": "Белоногов Юрий Владимирович",
    "григорьева мария ивановна": "Узкоглазова Мария Ивановна",
    "мостовских василий": "Мостовских Василий Васильевич",
    "диденко любовь валерьевна": "Диденко Любовь Валериевна",
    "бастриков александр": "Бастриков Александр Витальевич",
    "серебреникова ксения вячеславовна": "Юрлова Ксения Вячеславовна",
    "серебренникова ксения вячеславовна": "Юрлова Ксения Вячеславовна",
}

# Не показывать в отчёте (ошибки / не включать)
EXCLUDE_FROM_REPORT = {
    "ситникова полина андреевна",
    "харитонова наталья вячеславовна",
    "курбатов николай андреевич",
    "бердышева лейла агиловна",
    "перевалов леонид геннадьевич",
}

# Переводы / ручные правки магазина
STORE_OVERRIDES = {
    "кондратьева ирина владимировна": {
        "storeShort": "Заводоуковск",
        "store": "Магазин №31/г.Заводоуковск, ул. Черемуховая,23 с.1",
    },
}

# Отпуска, которых нет в файле магазина (перевод и т.п.)
MANUAL_STORE_VACATIONS = [
    {
        "fio": "Катрич Константин Дмитриевич",
        "storeShort": "Бабарынка",
        "start": "2026-09-14",
        "end": "2026-09-27",
        "days": 14,
        "note": "перевод в Бабарынку, даты по факту",
        "sourceFile": "ручная правка",
    },
    {
        "fio": "Кондратьева Ирина Владимировна",
        "storeShort": "Заводоуковск",
        "start": "2026-10-15",
        "end": "2026-10-28",
        "days": 14,
        "note": "магазин №31 Заводоуковск",
        "sourceFile": "ручная правка",
    },
]


def canonicalize_fio(fio: str) -> str:
    key = norm(fio)
    if key in FIO_CANONICAL:
        return FIO_CANONICAL[key]
    # "Котков С.А." / "Котков С. А."
    compact = re.sub(r"[.\s]+", " ", key).strip()
    if compact in FIO_CANONICAL:
        return FIO_CANONICAL[compact]
    return re.sub(r"\s+", " ", (fio or "")).strip()


def is_excluded(fio: str) -> bool:
    return norm(fio) in EXCLUDE_FROM_REPORT


def apply_store_override(rec: dict) -> None:
    ov = STORE_OVERRIDES.get(norm(rec.get("fio") or ""))
    if not ov:
        return
    rec["storeShort"] = ov["storeShort"]
    rec["store"] = ov.get("store") or ov["storeShort"]
    rec["isRetail"] = rec["storeShort"] in RETAIL_STORES


def apply_manual_store_vacations(people_list: list) -> None:
    by_key = {(norm(p["fio"]), p["storeShort"]): p for p in people_list}
    by_fio = defaultdict(list)
    for p in people_list:
        by_fio[norm(p["fio"])].append(p)
    for item in MANUAL_STORE_VACATIONS:
        fio = canonicalize_fio(item["fio"])
        store = item["storeShort"]
        vac = {
            "start": item["start"],
            "end": item["end"],
            "days": item["days"],
            "note": item.get("note") or "",
            "sourceFile": item.get("sourceFile") or "ручная правка",
        }
        rec = by_key.get((norm(fio), store))
        if not rec:
            # same person maybe still on old store name
            cands = by_fio.get(norm(fio), [])
            rec = next((p for p in cands if p["storeShort"] == store), None)
            if not rec and len(cands) == 1:
                rec = cands[0]
                rec["storeShort"] = store
                if store == "Заводоуковск":
                    rec["store"] = STORE_OVERRIDES.get(norm(fio), {}).get("store") or store
        if not rec:
            continue
        # avoid duplicates
        exists = any(v.get("start") == vac["start"] and v.get("end") == vac["end"] for v in rec.get("storeVacations") or [])
        if not exists:
            rec.setdefault("storeVacations", []).append(vac)
            rec["storeVacations"].sort(key=lambda x: x["start"] or "")
        rec["status"] = status_for(rec.get("official") or [], rec.get("storeVacations") or [])
        rec["isRetail"] = rec["storeShort"] in RETAIL_STORES
        if rec.get("roleGroup") in REPORT_ROLES:
            rec["isReportStaff"] = rec["isRetail"]


def role_group(position: str) -> str:
    p = norm(position)
    if p == "спк" or p.startswith("спк ") or p.startswith("спк/"):
        return "СПК"
    if "кладовщик" in p:
        return "Кладовщик"
    if "кассир" in p:
        return "Кассир"
    if "работник торгового" in p or p in {"ртз", "ртз."}:
        return "РТЗ"
    if "менеджер по работе с клиентами" in p:
        return "Менеджер"
    if "старш" in p and "продавец" in p:
        return "СПК"
    if "продавец" in p and "консультант" in p:
        return "ПК"
    if "управляющ" in p:
        return "Управляющий"
    if "оператор склада" in p:
        return "Оператор склада"
    if not p:
        return "Без должности"
    return position.strip()


def is_spk(rec) -> bool:
    return rec.get("roleGroup") == "СПК" or role_group(rec.get("position") or "") == "СПК"


def is_klad(rec) -> bool:
    return rec.get("roleGroup") == "Кладовщик" or role_group(rec.get("position") or "") == "Кладовщик"


STAFF_SKIP = {
    "штатная расстановка",
    "организация",
    "у михалыча",
    "дата отчета",
    "подразделение",
    "запланировано",
    "свободно",
    "учтено",
    "позиция",
    "сотрудник, состояние",
    "язык по умолчанию",
}
EMP_CELL_RE = re.compile(
    r"^([А-ЯЁ][а-яёА-ЯЁ\-]+(?:\s+[А-ЯЁа-яё\-]+){1,4}),\s+(.+)$"
)
NUM_CELL_RE = re.compile(r"^-?\d+,\d+$")


def parse_staff():
    """Current employees from 1C staffing MXL as of 14.09.2026."""
    if not STAFF_FILE.exists():
        return []
    raw = STAFF_FILE.read_bytes()
    start = raw.find(b"{")
    text = raw[start:].decode("utf-8", errors="replace") if start >= 0 else ""
    cells = [
        c.strip().replace("\xa0", " ")
        for c in re.findall(r'\{\s*"#"\s*,\s*"((?:[^"\\]|\\.)*)"\s*\}', text)
    ]
    current_store_raw = ""
    current_store = ""
    current_pos = ""
    rows = []
    for c in cells:
        if not c or c in {",", ", "} or NUM_CELL_RE.fullmatch(c) or norm(c) in STAFF_SKIP:
            continue
        m = EMP_CELL_RE.match(c)
        if m:
            rows.append(
                {
                    "fio": re.sub(r"\s+", " ", m.group(1)).strip(),
                    "status": m.group(2).strip(),
                    "position": current_pos,
                    "store": current_store_raw,
                    "storeShort": current_store,
                }
            )
            continue
        n = norm(c)
        if (
            "магазин" in n
            or c.startswith("Отдел")
            or c.startswith("Служба")
            or c.startswith("АУП")
            or c.startswith("АХО")
            or "розничн" in n
        ):
            current_store_raw = c
            current_store = store_short(c)
            current_pos = ""
            continue
        if "сентябр" in n or "дата" in n:
            continue
        current_pos = re.sub(r"\s+", " ", c).strip()
    return rows


def add_months(d: date, months: int) -> date:
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    leap = y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)
    dim = [31, 29 if leap else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
    return date(y, m, min(d.day, dim))


def is_summer_attention(d: date | None) -> bool:
    """Плановый отпуск в высокий сезон: 15.04 — 15.08."""
    if not d:
        return False
    start = date(d.year, 4, 15)
    end = date(d.year, 8, 15)
    return start <= d <= end


def parse_hire_dates():
    """Hire dates from 1C MXL export «даты приема»."""
    path = HIRE_FILE if HIRE_FILE.exists() else next(
        (p for p in ROOT.iterdir() if p.suffix.lower() == ".mxl" and "прием" in p.name.lower()),
        None,
    )
    if not path or not path.exists():
        return []
    raw = path.read_bytes()
    start = raw.find(b"{")
    text = raw[start:].decode("utf-8", errors="replace") if start >= 0 else ""
    cells = [
        c.strip().replace("\xa0", " ")
        for c in re.findall(r'\{\s*"#"\s*,\s*"((?:[^"\\]|\\.)*)"\s*\}', text)
    ]
    date_re = re.compile(r"^\d{1,2}\.\d{1,2}\.\d{4}$")
    fio_re = re.compile(r"^[А-ЯЁ][а-яёА-ЯЁ\-]+(?:\s+[А-ЯЁа-яё\-]+){1,3}$")
    skip = {
        "индикатор ошибки",
        "есть файлы",
        "дата приема",
        "дата",
        "номер",
        "организация",
        "сотрудник",
        "комментарий",
        "у михалыча",
    }
    rows = []
    for i, c in enumerate(cells):
        if not fio_re.match(c):
            continue
        if norm(c) in skip or c.lower().startswith("вместо"):
            continue
        back = cells[max(0, i - 5) : i]
        dates = [parse_date(x) for x in back if date_re.match(x)]
        dates = [d for d in dates if d]
        if not dates:
            continue
        rows.append({"fio": re.sub(r"\s+", " ", canonicalize_fio(c)).strip(), "hireDate": dates[0]})
    return rows


def match_hire_date(fio: str, hire_rows: list) -> date | None:
    fio_keys = name_keys(fio)
    fio_parts = set(name_tokens(fio))
    exact = [h for h in hire_rows if norm(h["fio"]) == norm(fio)]
    if exact:
        return exact[0]["hireDate"]
    strong = []
    for h in hire_rows:
        inter = fio_keys & name_keys(h["fio"])
        parts_inter = fio_parts & set(name_tokens(h["fio"]))
        if len(parts_inter) >= 2 or any(len(k.split()) >= 2 for k in inter):
            strong.append(h)
    if len(strong) == 1:
        return strong[0]["hireDate"]
    return None


def attach_staff(rec, staff_dict):
    hit, _how = match_person(rec["fio"], rec["storeShort"], staff_dict)
    if not hit:
        return None
    hit_store = hit.get("storeShort") or ""
    rec_store = rec.get("storeShort") or ""
    if rec_store in RETAIL_STORES or hit_store in RETAIL_STORES:
        if rec_store != hit_store:
            return None
    return hit


def parse_date(value, default_year=2026):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if not s or s in {"-", "—", "None", "nan"}:
        return None
    s = s.replace(",", ".").replace("г.", "").replace("г", "")
    s = re.sub(r"^[сcС]\s+", "", s.strip())
    # broken year 20266
    s = re.sub(r"20266", "2026", s)
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s[:10]):
        try:
            return date.fromisoformat(s[:10])
        except ValueError:
            pass
    m = re.search(r"(\d{1,2})[.\-/](\d{1,2})(?:[.\-/](\d{2,4}))?", s)
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), m.group(3)
    if mo > 12 and d <= 12:
        d, mo = mo, d
    if mo > 12 or d > 31:
        return None
    year = default_year
    if y:
        year = int(y)
        if year < 100:
            year += 2000
        if year > 2100:
            year = default_year
    try:
        return date(year, mo, d)
    except ValueError:
        return None


def parse_end_from_range(value, start, default_year=2026):
    """If cell looks like '01.02 - 15.02', return end date."""
    if value is None or isinstance(value, (date, datetime)):
        return None
    s = str(value)
    parts = re.split(r"\s*[-–—]\s*", s)
    if len(parts) < 2:
        return None
    end = parse_date(parts[-1], default_year=start.year if start else default_year)
    return end


def to_days(value):
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        n = int(value)
        return n if n > 0 else None
    s = str(value).strip().replace(",", ".")
    m = re.search(r"\d+", s)
    if not m:
        return None
    n = int(m.group())
    return n if 0 < n < 90 else None


def add_days(start: date, days: int) -> date:
    return start + timedelta(days=max(days, 1) - 1)


def parse_move_dates(value):
    if not value:
        return []
    if isinstance(value, (datetime, date)):
        d = parse_date(value)
        return [d] if d else []
    out = []
    for chunk in re.split(r"[;\n]+", str(value)):
        d = parse_date(chunk)
        if d:
            out.append(d)
    return out


def name_tokens(fio: str):
    s = norm(fio)
    s = s.replace(".", " ")
    s = re.sub(r"[^а-яa-z\s-]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return [t for t in s.split() if t]


def name_keys(fio: str) -> set[str]:
    parts = name_tokens(fio)
    keys = set()
    if not parts:
        return keys
    last = parts[0]
    keys.add(last)
    if len(parts) == 1:
        return keys
    rest = parts[1:]
    initials = "".join(p[0] for p in rest)
    keys.add(f"{last} {initials}")
    keys.add(f"{last}{initials}")
    keys.add(f"{last} {rest[0]}")
    keys.add(f"{last} {rest[0][0]}")
    if len(rest[0]) > 1:
        keys.add(f"{last} {rest[0][:4]}")
    if len(rest) >= 2:
        keys.add(f"{last} {rest[0]} {rest[1]}")
        keys.add(f"{last} {rest[0][0]}{rest[1][0]}")
        keys.add(f"{last} {rest[0][0]} {rest[1][0]}")
        if len(rest[0]) > 1:
            keys.add(f"{last} {rest[0]} {rest[1][0]}")
    return {k for k in keys if k}


def skip_cell(v) -> bool:
    if v is None:
        return True
    s = norm(str(v))
    if not s:
        return True
    if s in HEADER_SKIP:
        return True
    if "фамилия" in s or "наименование магазина" in s:
        return True
    if "подпись" in s or "составител" in s:
        return True
    if s.startswith("первая половина") or s.startswith("вторая часть"):
        return True
    return False


def rows_xlsx(path: Path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(values_only=True):
        yield row


def rows_xls(path: Path):
    book = xlrd.open_workbook(path)
    sh = book.sheet_by_index(0)
    for i in range(sh.nrows):
        vals = []
        for j in range(sh.ncols):
            cell = sh.cell(i, j)
            v = cell.value
            if cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    v = xlrd.xldate_as_datetime(v, book.datemode)
                except Exception:
                    pass
            vals.append(v)
        yield tuple(vals)


def iter_rows(path: Path):
    if path.suffix.lower() == ".xls":
        yield from rows_xls(path)
    else:
        yield from rows_xlsx(path)


def parse_official():
    people = {}  # key (name_norm_full, store_short) -> dict
    for row in iter_rows(OFFICIAL):
        if not row or len(row) < 7:
            continue
        store_raw = str(row[0] or "").strip()
        pos = str(row[1] or "").strip()
        fio = str(row[2] or "").strip()
        if skip_cell(fio) or skip_cell(store_raw):
            continue
        if skip_cell(pos) and pos in {"2", "1"}:
            continue
        if not re.search(r"[А-Яа-яЁё]", fio):
            continue
        days = to_days(row[4])
        planned = parse_date(row[6])
        actual = parse_date(row[7]) if len(row) > 7 else None
        moved = parse_move_dates(row[10] if len(row) > 10 else None)
        note = str(row[12] or "").strip() if len(row) > 12 else ""
        tab = str(row[3] or "").replace(".0", "").strip() if row[3] is not None else ""
        ss = store_short(store_raw)
        key = (norm(fio), ss)
        rec = people.get(key)
        if not rec:
            rec = {
                "fio": re.sub(r"\s+", " ", fio),
                "position": re.sub(r"\s+", " ", pos),
                "store": re.sub(r"\s+", " ", store_raw),
                "storeShort": ss,
                "tabNumber": tab,
                "official": [],
                "storeVacations": [],
                "sourceFile": "ЗУП Т-7",
            }
            people[key] = rec
        if not planned and not moved:
            continue
        start = planned
        d = days or 14
        end = add_days(start, d) if start else None
        rec["official"].append(
            {
                "start": start.isoformat() if start else None,
                "end": end.isoformat() if end else None,
                "days": d,
                "movedTo": [x.isoformat() for x in moved],
                "actual": actual.isoformat() if actual else None,
                "note": note,
            }
        )
    return people


def cell_store_name(rows_preview):
    for row in rows_preview[:6]:
        for cell in row[:4]:
            if not cell:
                continue
            s = str(cell).strip()
            ns = norm(s)
            if ns in {"обязательно", "наименование магазина", "на"}:
                continue
            if "магазин" in ns or store_short(s) != re.sub(r"\s+", " ", s).strip()[:40]:
                sh = store_short(s)
                if sh and sh != "Не указан":
                    return sh
            for _, short in STORE_SHORT:
                if short.lower() in ns:
                    return short
    return None


def parse_store_files():
    vacations = []  # unmatched raw vacations with storeShort + fio
    for path in sorted(STORE_DIR.iterdir()):
        if path.suffix.lower() not in {".xlsx", ".xls"}:
            continue
        rows = list(iter_rows(path))
        from_file = file_store(path.name)
        from_cell = cell_store_name(rows)
        ss = from_cell or from_file
        for row in rows:
            if not row or len(row) < 4:
                continue
            fio = row[1]
            if skip_cell(fio):
                continue
            fio = re.sub(r"\s+", " ", str(fio)).strip()
            if len(fio) < 3 or not re.search(r"[А-Яа-яЁё]", fio):
                continue
            days = to_days(row[2] if len(row) > 2 else None)
            raw_date = row[3] if len(row) > 3 else None
            fact_date = row[4] if len(row) > 4 else None
            note = str(row[6] or "").strip() if len(row) > 6 else ""
            start = parse_date(raw_date) or parse_date(fact_date)
            if not start:
                continue
            end_range = parse_end_from_range(raw_date, start)
            d = days
            if end_range and not d:
                d = (end_range - start).days + 1
            if not d:
                d = 14
            end = end_range or add_days(start, d)
            vacations.append(
                {
                    "fio": fio,
                    "storeShort": ss,
                    "sourceFile": path.name,
                    "start": start.isoformat(),
                    "end": end.isoformat(),
                    "days": d,
                    "note": note,
                    "rawDate": str(raw_date) if raw_date is not None else "",
                }
            )
    return vacations


def match_person(fio: str, store_short_name: str, people: dict):
    keys = name_keys(fio)
    last = name_tokens(fio)[0] if name_tokens(fio) else ""
    candidates = []
    for (nfio, ss), rec in people.items():
        rec_keys = name_keys(rec["fio"])
        if keys & rec_keys - {last} or (last and last in rec_keys and last in keys):
            score = 0
            inter = keys & rec_keys
            if rec["fio"].lower().startswith(last):
                score += 2
            if any(len(k.split()) >= 2 for k in inter):
                score += 5
            if ss == store_short_name:
                score += 8
            if last and last == name_tokens(rec["fio"])[0]:
                score += 3
            else:
                score -= 10
            candidates.append((score, rec))
    candidates.sort(key=lambda x: -x[0])
    strong = [c for c in candidates if c[0] >= 8]
    if not strong:
        # last name unique in this store
        same_store = [
            rec
            for (_, ss), rec in people.items()
            if ss == store_short_name and name_tokens(rec["fio"]) and name_tokens(rec["fio"])[0] == last
        ]
        if len(same_store) == 1:
            return same_store[0], "store-unique-lastname"
        return None, None
    best = strong[0]
    # if several with similar score at same store, require initials
    return best[1], "keys"


def periods_overlap(a_start, a_end, b_start, b_end) -> bool:
    return a_start <= b_end and b_start <= a_end


def parse_iso(s):
    return date.fromisoformat(s) if s else None


def fmt_ru(d: date | None) -> str:
    if not d:
        return "—"
    return d.strftime("%d.%m.%Y")


def fmt_period(start: str | None, end: str | None) -> str:
    return f"{fmt_ru(parse_iso(start))} — {fmt_ru(parse_iso(end))}"


def status_for(official, store_vacs):
    if official and store_vacs:
        for o in official:
            os, oe = parse_iso(o["start"]), parse_iso(o["end"])
            if not os:
                continue
            for s in store_vacs:
                ss, se = parse_iso(s["start"]), parse_iso(s["end"])
                if abs((os - ss).days) <= 3 or periods_overlap(os, oe, ss, se):
                    return "match"
        return "differ"
    if official:
        return "only_official"
    if store_vacs:
        return "only_store"
    return "empty"


def effective_periods(rec, prefer_store=True):
    """Periods used for conflict checks: store file first, else official (moved dates if present)."""
    if prefer_store and rec.get("storeVacations"):
        return [
            (parse_iso(v["start"]), parse_iso(v["end"]), rec)
            for v in rec["storeVacations"]
            if v.get("start")
        ]
    out = []
    for o in rec.get("official") or []:
        days = o.get("days") or 14
        if o.get("movedTo"):
            for m in o["movedTo"]:
                st = parse_iso(m)
                out.append((st, add_days(st, days), rec))
        elif o.get("start"):
            out.append((parse_iso(o["start"]), parse_iso(o["end"]), rec))
    return out


def find_conflicts(people_list):
    """СПК+кладовщик and same-role overlaps within one store."""
    by_store = defaultdict(list)
    for rec in people_list:
        if rec["storeShort"] in RETAIL_STORES:
            by_store[rec["storeShort"]].append(rec)
    conflicts = []
    for store, recs in sorted(by_store.items()):
        # 1) СПК + кладовщик
        klad = [r for r in recs if is_klad(r)]
        sales = [r for r in recs if is_spk(r)]
        k_periods = [p for r in klad for p in effective_periods(r)]
        s_periods = [p for r in sales for p in effective_periods(r)]
        seen = set()
        for ks, ke, kr in k_periods:
            if not ks or not ke:
                continue
            for ss, se, sr in s_periods:
                if not ss or not se:
                    continue
                if not periods_overlap(ks, ke, ss, se):
                    continue
                os = max(ks, ss)
                oe = min(ke, se)
                key = ("spk_klad", store, kr["fio"], sr["fio"], os.isoformat(), oe.isoformat())
                if key in seen:
                    continue
                seen.add(key)
                sales_role = sr.get("roleGroup") or role_group(sr["position"])
                conflicts.append(
                    {
                        "type": "spk_klad",
                        "storeShort": store,
                        "overlapStart": os.isoformat(),
                        "overlapEnd": oe.isoformat(),
                        "days": (oe - os).days + 1,
                        "roleGroup": "СПК+Кладовщик",
                        "personA": kr["fio"],
                        "positionA": kr["position"],
                        "roleA": "Кладовщик",
                        "periodA": fmt_period(ks.isoformat(), ke.isoformat()),
                        "personB": sr["fio"],
                        "positionB": sr["position"],
                        "roleB": sales_role,
                        "periodB": fmt_period(ss.isoformat(), se.isoformat()),
                        # legacy fields for older UI
                        "kladovshchik": kr["fio"],
                        "kladPosition": kr["position"],
                        "spk": sr["fio"],
                        "spkPosition": sr["position"],
                        "spkRole": sales_role,
                        "kladPeriod": fmt_period(ks.isoformat(), ke.isoformat()),
                        "spkPeriod": fmt_period(ss.isoformat(), se.isoformat()),
                    }
                )

        # 2) same role: two cashiers / two SPK / two ПК / …
        by_role = defaultdict(list)
        for r in recs:
            rg = r.get("roleGroup") or role_group(r.get("position") or "")
            if rg in REPORT_ROLES:
                by_role[rg].append(r)
        for rg, role_recs in by_role.items():
            if len(role_recs) < 2:
                continue
            periods = [(p, r) for r in role_recs for p in effective_periods(r)]
            for i, ((as_, ae, ar), _) in enumerate(periods):
                if not as_ or not ae:
                    continue
                for (bs, be, br), _ in periods[i + 1 :]:
                    if not bs or not be:
                        continue
                    if ar["fio"] == br["fio"]:
                        continue
                    if not periods_overlap(as_, ae, bs, be):
                        continue
                    os = max(as_, bs)
                    oe = min(ae, be)
                    pair = tuple(sorted([ar["fio"], br["fio"]]))
                    key = ("same_role", store, rg, pair[0], pair[1], os.isoformat(), oe.isoformat())
                    if key in seen:
                        continue
                    seen.add(key)
                    conflicts.append(
                        {
                            "type": "same_role",
                            "storeShort": store,
                            "overlapStart": os.isoformat(),
                            "overlapEnd": oe.isoformat(),
                            "days": (oe - os).days + 1,
                            "roleGroup": rg,
                            "personA": ar["fio"],
                            "positionA": ar.get("position") or rg,
                            "roleA": rg,
                            "periodA": fmt_period(as_.isoformat(), ae.isoformat()),
                            "personB": br["fio"],
                            "positionB": br.get("position") or rg,
                            "roleB": rg,
                            "periodB": fmt_period(bs.isoformat(), be.isoformat()),
                        }
                    )
    conflicts.sort(key=lambda x: (x["type"] != "spk_klad", x["overlapStart"], x["storeShort"]))
    return conflicts


def merge_people_by_fio(people: dict):
    """Canonicalize FIO and merge duplicate keys after maiden/abbrev renames."""
    merged = {}
    for (_old_key_fio, ss), rec in list(people.items()):
        fio = canonicalize_fio(rec["fio"])
        rec["fio"] = fio
        key = (norm(fio), ss)
        existing = merged.get(key)
        if not existing:
            merged[key] = rec
            continue
        # merge vacations into the richer record
        for o in rec.get("official") or []:
            if o not in existing["official"]:
                existing["official"].append(o)
        for v in rec.get("storeVacations") or []:
            if v not in existing["storeVacations"]:
                existing["storeVacations"].append(v)
        if not existing.get("position") and rec.get("position"):
            existing["position"] = rec["position"]
        if rec.get("tabNumber") and not existing.get("tabNumber"):
            existing["tabNumber"] = rec["tabNumber"]
        if not existing.get("storeFile") and rec.get("storeFile"):
            existing["storeFile"] = rec["storeFile"]
    people.clear()
    people.update(merged)


def main():
    people = parse_official()
    # apply maiden/abbrev names early so matching works
    for rec in people.values():
        rec["fio"] = canonicalize_fio(rec["fio"])
    merge_people_by_fio(people)

    store_vacs = parse_store_files()
    unmatched = []
    matched_n = 0
    for v in store_vacs:
        v["fio"] = canonicalize_fio(v["fio"])
        rec, how = match_person(v["fio"], v["storeShort"], people)
        item = {
            "start": v["start"],
            "end": v["end"],
            "days": v["days"],
            "note": v["note"],
            "sourceFile": v["sourceFile"],
        }
        if rec:
            rec["storeVacations"].append(item)
            if not rec.get("storeFile"):
                rec["storeFile"] = v["sourceFile"]
            matched_n += 1
        else:
            unmatched.append(v)

    # add unmatched store people as extra rows
    extra_keys = {}
    for v in unmatched:
        key = (norm(v["fio"]), v["storeShort"])
        rec = extra_keys.get(key) or people.get(key)
        if not rec:
            rec = {
                "fio": v["fio"],
                "position": "",
                "store": v["storeShort"],
                "storeShort": v["storeShort"],
                "tabNumber": "",
                "official": [],
                "storeVacations": [],
                "sourceFile": v["sourceFile"],
                "unmatched": True,
            }
            extra_keys[key] = rec
            people[key] = rec
        rec["storeVacations"].append(
            {
                "start": v["start"],
                "end": v["end"],
                "days": v["days"],
                "note": v["note"],
                "sourceFile": v["sourceFile"],
            }
        )

    merge_people_by_fio(people)

    people_list = []
    for rec in people.values():
        apply_store_override(rec)
        rec["roleGroup"] = role_group(rec["position"])
        rec["isRetail"] = rec["storeShort"] in RETAIL_STORES
        rec["isReportStaff"] = rec["isRetail"] and rec["roleGroup"] in REPORT_ROLES
        rec["status"] = status_for(rec["official"], rec["storeVacations"])
        rec["official"].sort(key=lambda x: x["start"] or "")
        rec["storeVacations"].sort(key=lambda x: x["start"] or "")
        people_list.append(rec)
    people_list.sort(key=lambda r: (r["storeShort"], r["roleGroup"], r["fio"]))

    staff_rows = parse_staff()
    for s in staff_rows:
        s["fio"] = canonicalize_fio(s["fio"])
        apply_store_override(s)
    staff_dict = {}
    for s in staff_rows:
        key = (norm(s["fio"]), s["storeShort"])
        staff_dict[key] = {
            "fio": s["fio"],
            "storeShort": s["storeShort"],
            "position": s["position"],
            "status": s["status"],
            "store": s["store"],
        }

    dropped_inactive = []
    active_list = []
    for rec in people_list:
        if is_excluded(rec["fio"]):
            continue
        hit = attach_staff(rec, staff_dict)
        if not hit:
            dropped_inactive.append(rec)
            continue
        rec["active"] = True
        rec["staffStatus"] = hit.get("status") or ""
        staff_role = role_group(hit.get("position") or "")
        if staff_role in REPORT_ROLES:
            rec["roleGroup"] = staff_role
            rec["position"] = hit.get("position") or rec["position"]
        rec["isRetail"] = rec["storeShort"] in RETAIL_STORES
        rec["isReportStaff"] = rec["isRetail"] and rec["roleGroup"] in REPORT_ROLES
        active_list.append(rec)
    active_keys = {(norm(p["fio"]), p["storeShort"]) for p in active_list}
    for s in staff_rows:
        if is_excluded(s["fio"]):
            continue
        if s["storeShort"] not in RETAIL_STORES:
            continue
        staff_role = role_group(s.get("position") or "")
        if staff_role not in REPORT_ROLES:
            continue
        key = (norm(s["fio"]), s["storeShort"])
        existing = next(
            (p for p in active_list if norm(p["fio"]) == key[0] and p["storeShort"] == key[1]),
            None,
        )
        if existing:
            if not existing.get("official") and not existing.get("storeVacations"):
                existing["isNewHire"] = True
            continue
        if key in active_keys:
            continue
        active_list.append(
            {
                "fio": s["fio"],
                "position": s["position"],
                "store": s["store"],
                "storeShort": s["storeShort"],
                "tabNumber": "",
                "official": [],
                "storeVacations": [],
                "sourceFile": "штатка 14.09.2026",
                "roleGroup": staff_role,
                "isRetail": True,
                "isReportStaff": True,
                "status": "empty",
                "active": True,
                "staffStatus": s.get("status") or "",
                "isNewHire": True,
                "hireDate": None,
            }
        )
        active_keys.add(key)

    people_list = active_list
    apply_manual_store_vacations(people_list)
    for rec in people_list:
        apply_store_override(rec)
        rec["isRetail"] = rec["storeShort"] in RETAIL_STORES
        rec["isReportStaff"] = rec["isRetail"] and rec.get("roleGroup") in REPORT_ROLES
        rec["status"] = status_for(rec.get("official") or [], rec.get("storeVacations") or [])
    people_list.sort(key=lambda r: (r["storeShort"], r["roleGroup"], r["fio"]))

    report_people = [
        p
        for p in people_list
        if p["isReportStaff"]
        or (p["isRetail"] and not p.get("position") and p.get("storeVacations"))
    ]
    new_hires = []
    hire_rows = parse_hire_dates()
    hire_matched = 0
    summer_attention = 0
    for p in report_people:
        if not p.get("isNewHire"):
            continue
        hire = p.get("hireDate")
        if isinstance(hire, str) and hire:
            hire = parse_iso(hire)
        if not hire:
            hire = match_hire_date(p["fio"], hire_rows)
        planned = add_months(hire, 6) if hire else None
        summer = is_summer_attention(planned)
        if hire:
            hire_matched += 1
        if summer:
            summer_attention += 1
        new_hires.append(
            {
                "fio": p["fio"],
                "position": p.get("position") or "",
                "storeShort": p["storeShort"],
                "roleGroup": p.get("roleGroup") or "",
                "hireDate": hire.isoformat() if hire else None,
                "plannedVacation": planned.isoformat() if planned else None,
                "summerAttention": summer,
            }
        )
    new_hires.sort(
        key=lambda r: (
            0 if r.get("summerAttention") else 1,
            r["storeShort"],
            r["roleGroup"],
            r["fio"],
        )
    )

    dismissed = []  # блок «уволенные» больше не публикуем

    conflicts = find_conflicts([p for p in report_people if not p.get("isNewHire")])
    conflict_stores = sorted({c["storeShort"] for c in conflicts})
    spk_klad = [c for c in conflicts if c.get("type") == "spk_klad"]
    same_role = [c for c in conflicts if c.get("type") == "same_role"]

    positions = sorted({p["position"] for p in report_people if p["position"] and not p.get("isNewHire")})
    stores = sorted({p["storeShort"] for p in report_people if p["storeShort"] and not p.get("isNewHire")})
    role_groups = [g for g in REPORT_ROLES if any(p["roleGroup"] == g and not p.get("isNewHire") for p in report_people)]

    active_scheduled = [p for p in report_people if not p.get("isNewHire")]
    payload = {
        "generated": date.today().isoformat(),
        "year": 2026,
        "stats": {
            "people": len(active_scheduled),
            "staffAsOf": "2026-09-14",
            "staffTotal": len(staff_rows),
            "droppedInactive": len(dropped_inactive),
            "officialRows": sum(len(p["official"]) for p in active_scheduled),
            "storeRows": sum(len(p["storeVacations"]) for p in active_scheduled),
            "matchedStoreRows": matched_n,
            "unmatchedStorePeople": sum(1 for p in active_scheduled if p.get("unmatched")),
            "conflicts": len(conflicts),
            "spkKladConflicts": len(spk_klad),
            "sameRoleConflicts": len(same_role),
            "conflictStores": len(conflict_stores),
            "differ": sum(1 for p in active_scheduled if p["status"] == "differ"),
            "match": sum(1 for p in active_scheduled if p["status"] == "match"),
            "onlyOfficial": sum(1 for p in active_scheduled if p["status"] == "only_official"),
            "onlyStore": sum(1 for p in active_scheduled if p["status"] == "only_store"),
            "newHires": len(new_hires),
            "newHiresWithDate": hire_matched,
            "newHiresSummer": summer_attention,
            "hireDatesFile": len(hire_rows),
        },
        "positions": positions,
        "stores": stores,
        "roleGroups": role_groups,
        "reportRoles": list(REPORT_ROLES),
        "retailStores": sorted(s for s in RETAIL_STORES if any(p["storeShort"] == s for p in active_scheduled)),
        "conflictStores": conflict_stores,
        "people": [p for p in people_list if not p.get("isNewHire") and not is_excluded(p["fio"])],
        "newHires": new_hires,
        "dismissed": [],
        "conflicts": conflicts,
    }
    OUT_JS.write_text(
        "window.VACATION_DATA = " + json.dumps(payload, ensure_ascii=False, indent=None) + ";\n",
        encoding="utf-8",
    )
    summary = ROOT / "_parse_stats.txt"
    lines = [
        f"people={payload['stats']['people']}",
        f"staffTotal={payload['stats']['staffTotal']}",
        f"droppedInactive={payload['stats']['droppedInactive']}",
        f"officialPeriods={payload['stats']['officialRows']}",
        f"storePeriods={payload['stats']['storeRows']}",
        f"matchedStoreRows={matched_n}",
        f"unmatchedStorePeople={payload['stats']['unmatchedStorePeople']}",
        f"conflicts={len(conflicts)}",
        f"conflictStores={payload['stats']['conflictStores']}",
        f"status match={payload['stats']['match']} differ={payload['stats']['differ']} onlyOfficial={payload['stats']['onlyOfficial']} onlyStore={payload['stats']['onlyStore']}",
        "",
        "DROPPED INACTIVE:",
    ]
    for rec in dropped_inactive:
        lines.append(f"  {rec['storeShort']:16} | {rec['fio']:40} | {rec.get('position','')}")
    lines.append("")
    lines.append("UNMATCHED:")
    for v in unmatched:
        lines.append(f"  {v['storeShort']:16} | {v['fio']:40} | {v['start']} {v['days']}д | {v['sourceFile']}")
    lines.append("")
    lines.append("CONFLICTS:")
    for c in conflicts:
        if c.get("type") == "same_role":
            lines.append(
                f"  SAME {c['roleGroup']:12} {c['storeShort']:16} {c['overlapStart']}—{c['overlapEnd']} | "
                f"{c['personA']} || {c['personB']}"
            )
        else:
            lines.append(
                f"  SPK+KLAD {c['storeShort']:16} {c['overlapStart']}—{c['overlapEnd']} | "
                f"клад {c.get('kladovshchik') or c['personA']} || СПК {c.get('spk') or c['personB']}"
            )
    summary.write_text("\n".join(lines), encoding="utf-8")
    print(
        "wrote",
        OUT_JS,
        "reportPeople",
        len(report_people),
        "allPeople",
        len(people_list),
        "conflicts",
        len(conflicts),
        "unmatched",
        payload["stats"]["unmatchedStorePeople"],
    )


if __name__ == "__main__":
    main()
