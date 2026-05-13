"""Extract clean tour data from the two Excel files for embedding in index.html."""
import json
import openpyxl
import sys
import re
import pathlib

sys.stdout.reconfigure(encoding="utf-8")

ROOT = pathlib.Path(__file__).parent
TG = pathlib.Path(r"C:\Users\elbics\Downloads\Telegram Desktop")
MARCH = TG / "Тур для Адемы v3.xlsx"
JULY = TG / "Адема Июль 2026 (draft ver_5) (3).xlsx"


def clean(v):
    if v is None:
        return None
    s = str(v).strip()
    if s in ("", "#NAME?", "#VALUE!"):
        return None
    return s


def parse_excel_time(v):
    """Excel stores time as fractional day. Convert 0.5 → '12:00'."""
    if v is None:
        return None
    try:
        f = float(v)
        if 0 < f < 1:
            total_seconds = round(f * 86400)
            h = total_seconds // 3600
            m = (total_seconds % 3600) // 60
            return f"{h:02d}:{m:02d}"
    except (TypeError, ValueError):
        pass
    return v


def num(v):
    if v is None:
        return None
    try:
        f = float(v)
        return round(f, 2)
    except (TypeError, ValueError):
        return None


def get_rows(path, sheet):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    return list(ws.iter_rows(values_only=True))


def parse_march():
    rows = get_rows(MARCH, "Программа")
    days = []
    cur_day = None
    current_date = None
    for r in rows[1:]:  # skip header
        r = [clean(c) for c in r]
        # columns: [_, Date, Schedule, Time, Map, TransportCost, TransportTime, TransportType, Hours, Ticket, SiteLink, _]
        if not r or len(r) < 11:
            continue
        date_val = r[1] if len(r) > 1 else None
        schedule = r[2] if len(r) > 2 else None
        time_val = parse_excel_time(r[3] if len(r) > 3 else None)
        map_link = r[4] if len(r) > 4 else None
        t_cost = num(r[5] if len(r) > 5 else None)
        t_time = r[6] if len(r) > 6 else None
        t_type = r[7] if len(r) > 7 else None
        hours = r[8] if len(r) > 8 else None
        ticket = num(r[9] if len(r) > 9 else None)
        site = r[10] if len(r) > 10 else None

        # detect day header (has date but other fields might still be there)
        if date_val and "Марта" in date_val or date_val and "Апреля" in date_val:
            current_date = date_val
            cur_day = {"date": current_date, "events": []}
            days.append(cur_day)
            # If this row also has a schedule event, add it
            if schedule:
                cur_day["events"].append({
                    "name": schedule,
                    "time": time_val,
                    "map": map_link if map_link and (map_link.startswith("http") or "Открыть" in map_link or "откуда" in map_link) else None,
                    "transport_cost": t_cost,
                    "transport_time": t_time,
                    "transport_type": t_type,
                    "hours": hours,
                    "ticket": ticket,
                    "site": site,
                })
        elif date_val and ("28.03 бронь" in date_val or "На месте" in date_val or "17:00-17:30" in date_val or "Аниме место" in date_val):
            # special annotation on row; treat as note for current event
            if cur_day and schedule:
                cur_day["events"].append({
                    "name": schedule,
                    "time": time_val,
                    "tag": date_val,
                    "map": map_link if map_link and (map_link.startswith("http") or "Открыть" in map_link or "откуда" in map_link) else None,
                    "transport_cost": t_cost,
                    "transport_time": t_time,
                    "transport_type": t_type,
                    "hours": hours,
                    "ticket": ticket,
                    "site": site,
                })
        elif schedule and cur_day is not None:
            cur_day["events"].append({
                "name": schedule,
                "time": time_val,
                "map": map_link if map_link and (map_link.startswith("http") or "Открыть" in map_link or "откуда" in map_link) else None,
                "transport_cost": t_cost,
                "transport_time": t_time,
                "transport_type": t_type,
                "hours": hours,
                "ticket": ticket,
                "site": site,
            })

    # hotels
    hotel_rows = get_rows(MARCH, "Отели")
    hotels = []
    cur_hotel_section = None
    for r in hotel_rows:
        r = [clean(c) for c in r]
        if not r:
            continue
        first = r[1] if len(r) > 1 else None
        second = r[2] if len(r) > 2 else None
        third = r[3] if len(r) > 3 else None
        fourth = r[4] if len(r) > 4 else None
        # Section header detection
        if first and ("с " in first and ("марта" in first or "апреля" in first)):
            cur_hotel_section = first
            continue
        if first == "Отель":
            continue  # header within section
        # Hotel entry: name, link, price, питание
        if first and second and third and "Стоимость" not in first:
            hotels.append({
                "city_range": cur_hotel_section,
                "name": first,
                "link_text": second,
                "price": num(third),
                "meal": fourth,
                "note": r[5] if len(r) > 5 else None,
            })

    return {
        "title": "Тур для Адемы — v3",
        "dates": "16 марта — 13 апреля 2025",
        "summary": {
            "esim": 26,
            "hotels": 1176,
            "transport_jpy": 91520,
            "tickets_jpy": 79290,
            "currency_note": "Транспорт и билеты в иенах; отели и e-sim в долларах",
        },
        "days": days,
        "hotels": hotels,
    }


def parse_july():
    rows = get_rows(JULY, "Маршрут")
    days = []
    cur_day = None
    last_date = None
    last_dow = None
    last_city = None
    date_re = re.compile(r"^\d{1,2}\s+(July|August|Июля|Августа)", re.IGNORECASE)
    for r in rows[2:]:  # skip 2 header rows
        r = [clean(c) for c in r]
        if not r or len(r) < 13:
            continue
        date_val = r[1] if len(r) > 1 else None
        dow = r[2] if len(r) > 2 else None
        city = r[3] if len(r) > 3 else None
        # skip header row if present
        if date_val == "Dates" or city == "City":
            continue
        time_val = r[4] if len(r) > 4 else None
        place = r[5] if len(r) > 5 else None
        map_link = r[6] if len(r) > 6 else None
        t_cost = num(r[7] if len(r) > 7 else None)
        t_time = r[8] if len(r) > 8 else None
        t_type = r[9] if len(r) > 9 else None
        hours = r[10] if len(r) > 10 else None
        ticket = num(r[11] if len(r) > 11 else None)
        site = r[12] if len(r) > 12 else None

        if date_val:
            last_date = date_val
            last_dow = dow if dow else last_dow
            last_city = city if city else last_city
            cur_day = {"date": last_date, "dow": last_dow, "city": last_city, "events": []}
            days.append(cur_day)
        if city and not date_val:
            last_city = city  # city can change mid-day

        if place and cur_day is not None:
            ev = {
                "name": place,
                "time": time_val,
                "city": city if city else last_city,
                "map": map_link if map_link and map_link.startswith("http") else None,
                "transport_cost": t_cost,
                "transport_time": t_time,
                "transport_type": t_type,
                "hours": hours,
                "ticket": ticket,
                "site": site,
            }
            cur_day["events"].append(ev)

    # attractions catalog
    att_rows = get_rows(JULY, "Attractions")
    attractions = []
    cur_section = None
    CITY_NAMES = ("Удзи", "Киото", "Такасима", "Хиконе", "Нагоя", "Сидзуока", "Фудзикавагутико", "Токио", "Чиба")
    for r in att_rows:
        r = [clean(c) for c in r]
        if not r:
            continue
        first = r[1] if len(r) > 1 else None
        price = num(r[2] if len(r) > 2 else None)
        comment = r[3] if len(r) > 3 else None
        status = r[4] if len(r) > 4 else None
        # Section header: city name (may have "Цена/Комментарий/Статус" sub-headers in same row)
        if first in CITY_NAMES:
            cur_section = first
            continue
        # Skip sub-column-header rows
        if first in ("Цена", "Комментарий", "Статус"):
            continue
        if first and first != "Total":
            attractions.append({
                "city": cur_section,
                "name": first,
                "price": price,
                "comment": comment,
                "status": status,
            })

    # hotels
    hotel_rows = get_rows(JULY, "Отели")
    hotels = []
    cur_hotel_section = None
    for r in hotel_rows:
        r = [clean(c) for c in r]
        if not r:
            continue
        first = r[1] if len(r) > 1 else None
        room = r[2] if len(r) > 2 else None
        price = num(r[3] if len(r) > 3 else None)
        meal = r[4] if len(r) > 4 else None
        link = r[5] if len(r) > 5 else None
        district = r[6] if len(r) > 6 else None
        note = r[7] if len(r) > 7 else None
        # section header like "11-12 июля Кансай"
        if first and ("июля" in first or "августа" in first) and not room and not price:
            cur_hotel_section = first
            continue
        if first == "Отель":
            continue
        if first and price is not None:
            hotels.append({
                "city_range": cur_hotel_section,
                "name": first,
                "room": room,
                "price": price,
                "meal": meal,
                "link": link if link and link.startswith("http") else None,
                "district": district,
                "note": note,
            })

    # flights
    fl_rows = get_rows(JULY, "Авиабилеты")
    flights = []
    if len(fl_rows) >= 3:
        r = [clean(c) for c in fl_rows[2]]
        # ALM-OSA
        flights.append({
            "route": "Алматы → Осака",
            "date": r[0],
            "airline": r[1],
            "duration": r[2],
            "stop": r[3],
            "price": num(r[4]),
            "baggage": r[5],
            "link": r[6] if r[6] and r[6].startswith("http") else None,
        })
        flights.append({
            "route": "Токио → Алматы",
            "date": r[9],
            "airline": r[10],
            "duration": r[11],
            "stop": r[12],
            "price": num(r[13]),
            "baggage": r[14],
            "link": r[15] if r[15] and r[15].startswith("http") else None,
        })

    # directions
    dir_rows = get_rows(JULY, "Направления")
    directions = []
    for r in dir_rows[3:]:
        r = [clean(c) for c in r]
        if not r:
            continue
        name = r[1] if len(r) > 1 else None
        days_count = r[2] if len(r) > 2 else None
        logistics = num(r[3] if len(r) > 3 else None)
        if name and name != "Активный сценарий":
            directions.append({"city": name, "days": days_count, "logistics": logistics})

    return {
        "title": "Адема в Японии — июль 2026",
        "dates": "11 июля — 2 августа 2026",
        "summary": {
            "flights": 910,
            "esim": 25,
            "hotels": 1211,
            "luggage": 40.48,
            "logistics": 363.00,
            "tour_op": 200,
            "tickets": 548.40,
            "total": 3297.87,
            "currency_note": "Все суммы в USD. JPY→USD по курсу 151.93 (Kansai AP)",
        },
        "days": days,
        "attractions": attractions,
        "hotels": hotels,
        "flights": flights,
        "directions": directions,
    }


data = {
    "march": parse_march(),
    "july": parse_july(),
}

out_json = json.dumps(data, ensure_ascii=False, indent=2)
out_min = json.dumps(data, ensure_ascii=False, separators=(",", ":"))

(ROOT / "tour_data.json").write_text(out_json, encoding="utf-8")
(ROOT / "tour_data.min.js").write_text("window.TOURS=" + out_min + ";\n", encoding="utf-8")

print(f"Wrote tour_data.json ({len(out_json):,} chars)")
print(f"Wrote tour_data.min.js ({len(out_min):,} chars)")
print(f"\nMarch tour: {len(data['march']['days'])} days, {sum(len(d['events']) for d in data['march']['days'])} events, {len(data['march']['hotels'])} hotels")
print(f"July tour: {len(data['july']['days'])} days, {sum(len(d['events']) for d in data['july']['days'])} events, {len(data['july']['hotels'])} hotels, {len(data['july']['attractions'])} attractions, {len(data['july']['flights'])} flights")
