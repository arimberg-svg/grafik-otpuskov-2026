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
    "Ожогино",
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
}

STORE_SHORT = [
    ("сантехника", "Сантехника"),
    ("дружбы 66", "Дружба"),
    ("авторемонтная", "Авторемонтная"),
    ("республики", "Республики"),
    ("московскийтракт 120", "МСК 120А"),
    ("московский тракт 120", "МСК 120А"),
    ("бабарынка", "Бабарынка"),
    ("ожогин", "Ожогино"),
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
    ("ожогино", "Ожогино"),
    ("тюнево", "Тюнево"),
    ("сантехника", "Сантехника"),
    ("чайка", "Чайка"),
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


def role_group(position: str) -> str:
    p = norm(position)
    if "кладовщик" in p:
        return "Кладовщик"
    if "продавец" in p and "консультант" in p:
        return "СПК"
    if "старший продавец" in p:
        return "СПК"
    if "кассир" in p:
        return "Кассир"
    if "управляющ" in p:
        return "Управляющий"
    if "работник торгового" in p:
        return "Работник зала"
    if "оператор склада" in p:
        return "Оператор склада"
    if not p:
        return "Без должности"
    return position.strip()


def is_spk(position: str) -> bool:
    return role_group(position) == "СПК"


def is_klad(position: str) -> bool:
    return role_group(position) == "Кладовщик"


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
    by_store = defaultdict(list)
    for rec in people_list:
        if rec["storeShort"] in RETAIL_STORES:
            by_store[rec["storeShort"]].append(rec)
    conflicts = []
    for store, recs in sorted(by_store.items()):
        klad = [r for r in recs if is_klad(r["position"])]
        spk = [r for r in recs if is_spk(r["position"])]
        k_periods = [p for r in klad for p in effective_periods(r)]
        s_periods = [p for r in spk for p in effective_periods(r)]
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
                key = (store, kr["fio"], sr["fio"], os.isoformat(), oe.isoformat())
                if key in seen:
                    continue
                seen.add(key)
                conflicts.append(
                    {
                        "storeShort": store,
                        "overlapStart": os.isoformat(),
                        "overlapEnd": oe.isoformat(),
                        "days": (oe - os).days + 1,
                        "kladovshchik": kr["fio"],
                        "kladPosition": kr["position"],
                        "spk": sr["fio"],
                        "spkPosition": sr["position"],
                        "kladPeriod": f"{ks.isoformat()} — {ke.isoformat()}",
                        "spkPeriod": f"{ss.isoformat()} — {se.isoformat()}",
                    }
                )
    conflicts.sort(key=lambda x: (x["overlapStart"], x["storeShort"]))
    return conflicts


def main():
    people = parse_official()
    store_vacs = parse_store_files()
    unmatched = []
    matched_n = 0
    for v in store_vacs:
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
        rec = extra_keys.get(key)
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

    people_list = []
    for rec in people.values():
        rec["roleGroup"] = role_group(rec["position"])
        rec["isRetail"] = rec["storeShort"] in RETAIL_STORES
        rec["status"] = status_for(rec["official"], rec["storeVacations"])
        rec["official"].sort(key=lambda x: x["start"] or "")
        rec["storeVacations"].sort(key=lambda x: x["start"] or "")
        people_list.append(rec)
    people_list.sort(key=lambda r: (r["storeShort"], r["roleGroup"], r["fio"]))

    conflicts = find_conflicts(people_list)

    positions = sorted({p["position"] for p in people_list if p["position"]})
    stores = sorted({p["storeShort"] for p in people_list if p["storeShort"]})
    role_groups = []
    for preferred in ["Кладовщик", "СПК", "Кассир", "Управляющий", "Работник зала", "Оператор склада"]:
        if any(p["roleGroup"] == preferred for p in people_list):
            role_groups.append(preferred)
    for g in sorted({p["roleGroup"] for p in people_list}):
        if g not in role_groups:
            role_groups.append(g)

    payload = {
        "generated": date.today().isoformat(),
        "year": 2026,
        "stats": {
            "people": len(people_list),
            "officialRows": sum(len(p["official"]) for p in people_list),
            "storeRows": sum(len(p["storeVacations"]) for p in people_list),
            "matchedStoreRows": matched_n,
            "unmatchedStorePeople": len(extra_keys),
            "conflicts": len(conflicts),
            "differ": sum(1 for p in people_list if p["status"] == "differ"),
            "match": sum(1 for p in people_list if p["status"] == "match"),
            "onlyOfficial": sum(1 for p in people_list if p["status"] == "only_official"),
            "onlyStore": sum(1 for p in people_list if p["status"] == "only_store"),
        },
        "positions": positions,
        "stores": stores,
        "roleGroups": role_groups,
        "retailStores": sorted(RETAIL_STORES),
        "people": people_list,
        "conflicts": conflicts,
    }
    OUT_JS.write_text(
        "window.VACATION_DATA = " + json.dumps(payload, ensure_ascii=False, indent=None) + ";\n",
        encoding="utf-8",
    )
    summary = ROOT / "_parse_stats.txt"
    lines = [
        f"people={payload['stats']['people']}",
        f"officialPeriods={payload['stats']['officialRows']}",
        f"storePeriods={payload['stats']['storeRows']}",
        f"matchedStoreRows={matched_n}",
        f"unmatchedStorePeople={len(extra_keys)}",
        f"conflicts={len(conflicts)}",
        f"status match={payload['stats']['match']} differ={payload['stats']['differ']} onlyOfficial={payload['stats']['onlyOfficial']} onlyStore={payload['stats']['onlyStore']}",
        "",
        "UNMATCHED:",
    ]
    for v in unmatched:
        lines.append(f"  {v['storeShort']:16} | {v['fio']:40} | {v['start']} {v['days']}д | {v['sourceFile']}")
    lines.append("")
    lines.append("CONFLICTS:")
    for c in conflicts:
        lines.append(
            f"  {c['storeShort']:16} {c['overlapStart']}—{c['overlapEnd']} | "
            f"клад {c['kladovshchik']} || СПК {c['spk']}"
        )
    summary.write_text("\n".join(lines), encoding="utf-8")
    print("wrote", OUT_JS, "people", len(people_list), "conflicts", len(conflicts), "unmatched", len(extra_keys))


if __name__ == "__main__":
    main()
