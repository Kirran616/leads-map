"""
Leadmap - a simple lead generation tool powered by OpenStreetMap.

Data sources (no API key required):
  * Nominatim  -> turns a city name into a bounding box
  * Overpass   -> finds businesses inside that box

Data is (c) OpenStreetMap contributors, licensed under the ODbL.
https://www.openstreetmap.org/copyright
"""
import csv
import io
import os
import re
import threading
import time
from datetime import datetime

import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

load_dotenv()

app = Flask(__name__)

# ---------------------------------------------------------------- settings
CONTACT = os.getenv("OSM_CONTACT", "").strip()
USER_AGENT = f"Leadmap/1.0 ({CONTACT or 'set OSM_CONTACT in .env'})"
NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
OVERPASS_URLS = [
    os.getenv("OVERPASS_URL", "https://overpass-api.de/api/interpreter"),
    "https://overpass.kumi.systems/api/interpreter",
]
ATTRIBUTION = "\u00a9 OpenStreetMap contributors (ODbL) - openstreetmap.org/copyright"
MAX_SPAN_DEGREES = 1.5   # refuse huge areas such as whole countries
ALLOWED_LIMITS = (50, 100, 250, 500)

# Friendly words -> OpenStreetMap tags
CATEGORIES = {
    "restaurant": [("amenity", "restaurant")],
    "cafe": [("amenity", "cafe")],
    "coffee": [("amenity", "cafe")],
    "bar": [("amenity", "bar"), ("amenity", "pub")],
    "pub": [("amenity", "pub")],
    "fast food": [("amenity", "fast_food")],
    "bakery": [("shop", "bakery")],
    "hotel": [("tourism", "hotel"), ("tourism", "guest_house")],
    "pharmacy": [("amenity", "pharmacy")],
    "dentist": [("amenity", "dentist"), ("healthcare", "dentist")],
    "doctor": [("amenity", "doctors"), ("healthcare", "doctor")],
    "clinic": [("amenity", "clinic"), ("healthcare", "clinic")],
    "hospital": [("amenity", "hospital")],
    "school": [("amenity", "school")],
    "bank": [("amenity", "bank")],
    "gym": [("leisure", "fitness_centre")],
    "salon": [("shop", "hairdresser"), ("shop", "beauty")],
    "hairdresser": [("shop", "hairdresser")],
    "barber": [("shop", "hairdresser")],
    "supermarket": [("shop", "supermarket")],
    "grocery": [("shop", "supermarket"), ("shop", "convenience")],
    "clothes": [("shop", "clothes")],
    "electronics": [("shop", "electronics")],
    "mobile": [("shop", "mobile_phone")],
    "car repair": [("shop", "car_repair")],
    "mechanic": [("shop", "car_repair")],
    "car dealer": [("shop", "car")],
    "real estate": [("office", "estate_agent")],
    "lawyer": [("office", "lawyer")],
    "accountant": [("office", "accountant")],
    "insurance": [("office", "insurance")],
    "travel agency": [("shop", "travel_agency"), ("office", "travel_agent")],
    "veterinary": [("amenity", "veterinary")],
    "florist": [("shop", "florist")],
    "bookstore": [("shop", "books")],
    "furniture": [("shop", "furniture")],
}

# ------------------------------------------------------------- helpers
_nominatim_lock = threading.Lock()
_last_nominatim_call = 0.0
_geocode_cache = {}


def clean_text(value, max_len=80):
    """Keep only safe characters so user text can't break an Overpass query."""
    value = re.sub(r"[^\w\s&'\-]", " ", value or "", flags=re.UNICODE)
    return re.sub(r"\s+", " ", value).strip()[:max_len]


def geocode(place):
    """City name -> (south, west, north, east, display_name). Max 1 request/second."""
    global _last_nominatim_call
    key = place.lower()
    if key in _geocode_cache:
        return _geocode_cache[key]

    with _nominatim_lock:
        wait = 1.1 - (time.time() - _last_nominatim_call)
        if wait > 0:
            time.sleep(wait)
        resp = requests.get(
            NOMINATIM_URL,
            params={"q": place, "format": "jsonv2", "limit": 1},
            headers={"User-Agent": USER_AGENT},
            timeout=15,
        )
        _last_nominatim_call = time.time()

    if resp.status_code == 429:
        raise ApiError("OpenStreetMap is rate limiting requests. Wait a minute and try again.", 429)
    resp.raise_for_status()
    data = resp.json()
    if not data:
        raise ApiError(f'Could not find a place called "{place}". Try "City, Country".', 404)

    s, n, w, e = (float(x) for x in data[0]["boundingbox"])
    result = (s, w, n, e, data[0].get("display_name", place))
    _geocode_cache[key] = result
    return result


def build_query(keyword, bbox, limit):
    s, w, n, e = bbox
    box = f"({s},{w},{n},{e})"
    kw = clean_text(keyword).lower()
    parts = []

    tags = CATEGORIES.get(kw) or CATEGORIES.get(kw.rstrip("s"))
    if tags:
        for k, v in tags:
            parts.append(f'nwr["{k}"="{v}"]["name"]{box};')
    else:
        # Unknown keyword: match the business name or its type
        rx = re.escape(kw).replace("\\ ", " ")
        parts.append(f'nwr["name"~"{rx}",i]{box};')
        for k in ("shop", "amenity", "office", "craft", "healthcare", "tourism"):
            parts.append(f'nwr["{k}"~"^{rx.replace(" ", "_")}$",i]["name"]{box};')

    body = "\n  ".join(parts)
    return f"[out:json][timeout:30];\n(\n  {body}\n);\nout center tags {limit};"


def run_overpass(query):
    last_error = "The OpenStreetMap search server did not respond."
    for url in OVERPASS_URLS:
        try:
            resp = requests.post(
                url, data={"data": query}, headers={"User-Agent": USER_AGENT}, timeout=45
            )
        except requests.RequestException:
            continue
        if resp.status_code == 200:
            return resp.json().get("elements", [])
        if resp.status_code == 429:
            last_error = "The OpenStreetMap server is busy. Wait a minute and try again."
        elif resp.status_code in (504, 502, 503):
            last_error = "The search timed out. Try a smaller city or a narrower keyword."
        elif resp.status_code == 400:
            raise ApiError("That keyword could not be searched. Try a simpler word.", 400)
    raise ApiError(last_error, 503)


def address_from_tags(t):
    street = " ".join(x for x in (t.get("addr:housenumber"), t.get("addr:street")) if x)
    parts = [street, t.get("addr:suburb"), t.get("addr:city") or t.get("addr:town") or t.get("addr:village"),
             t.get("addr:postcode"), t.get("addr:state")]
    return ", ".join(p for p in parts if p) or t.get("addr:full", "")


def category_from_tags(t):
    for k in ("amenity", "shop", "office", "craft", "healthcare", "tourism", "leisure"):
        if t.get(k):
            return t[k].replace("_", " ")
    return ""


def to_lead(el):
    t = el.get("tags", {})
    center = el.get("center") or {}
    lat = el.get("lat", center.get("lat"))
    lon = el.get("lon", center.get("lon"))
    return {
        "name": t.get("name", ""),
        "category": category_from_tags(t),
        "address": address_from_tags(t),
        "phone": t.get("phone") or t.get("contact:phone") or "",
        "website": t.get("website") or t.get("contact:website") or "",
        "email": t.get("email") or t.get("contact:email") or "",
        "osm_id": f'{el["type"]}/{el["id"]}',
        "lat": lat,
        "lon": lon,
    }


def dedupe(leads):
    seen, out = set(), []
    for lead in leads:
        key = (lead["name"].lower(), lead["phone"] or lead["address"].lower() or lead["osm_id"])
        if lead["osm_id"] in seen or key in seen:
            continue
        seen.add(lead["osm_id"])
        seen.add(key)
        out.append(lead)
    return out


class ApiError(Exception):
    def __init__(self, message, status=400):
        super().__init__(message)
        self.message, self.status = message, status


@app.errorhandler(ApiError)
def handle_api_error(err):
    return jsonify(error=err.message), err.status


@app.errorhandler(requests.RequestException)
def handle_network_error(_err):
    return jsonify(error="Could not reach OpenStreetMap. Check your internet connection."), 502


# --------------------------------------------------------------- routes
@app.get("/")
def index():
    return render_template("index.html", attribution=ATTRIBUTION, has_contact=bool(CONTACT))


@app.post("/api/search")
def search():
    body = request.get_json(silent=True) or {}
    keyword = clean_text(body.get("keyword"))
    location = (body.get("location") or "").strip()[:120]
    limit = int(body.get("limit") or 100)
    limit = limit if limit in ALLOWED_LIMITS else 100

    if not keyword:
        raise ApiError("Enter a business type or keyword, for example: dentist.")
    if not location:
        raise ApiError("Enter a city or location, for example: Lahore, Pakistan.")

    s, w, n, e, display = geocode(location)
    if (n - s) > MAX_SPAN_DEGREES or (e - w) > MAX_SPAN_DEGREES:
        raise ApiError("That area is too large. Enter a city or town instead of a region or country.")

    elements = run_overpass(build_query(keyword, (s, w, n, e), limit))
    leads = [to_lead(el) for el in elements if el.get("tags", {}).get("name")]
    raw_count = len(leads)
    leads = dedupe(leads)

    return jsonify(
        leads=leads,
        total=len(leads),
        duplicates_removed=raw_count - len(leads),
        place=display,
        attribution=ATTRIBUTION,
    )


COLUMNS = [
    ("name", "Business Name"), ("category", "Category"), ("address", "Address"),
    ("phone", "Phone"), ("website", "Website"), ("email", "Email"),
    ("osm_id", "OpenStreetMap ID"), ("lat", "Latitude"), ("lon", "Longitude"),
]


def safe_cell(value):
    """Stop spreadsheet formula injection (=, +, -, @ at the start of a cell)."""
    text = "" if value is None else str(value)
    return "'" + text if text[:1] in ("=", "+", "-", "@") else text


def get_rows():
    rows = (request.get_json(silent=True) or {}).get("leads")
    if not isinstance(rows, list) or not rows:
        raise ApiError("There are no results to export.")
    return rows[:5000]


@app.post("/api/export/csv")
def export_csv():
    rows = get_rows()
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow([label for _, label in COLUMNS] + ["Source"])
    for r in rows:
        writer.writerow([safe_cell(r.get(k)) for k, _ in COLUMNS] + [ATTRIBUTION])
    data = io.BytesIO(buf.getvalue().encode("utf-8-sig"))  # utf-8-sig opens cleanly in Excel
    return send_file(data, mimetype="text/csv", as_attachment=True, download_name=filename("csv"))


@app.post("/api/export/xlsx")
def export_xlsx():
    rows = get_rows()
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    ws.append([label for _, label in COLUMNS])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="14343B")
        cell.alignment = Alignment(vertical="center")
    for r in rows:
        ws.append([safe_cell(r.get(k)) for k, _ in COLUMNS])
    for i, (key, label) in enumerate(COLUMNS, start=1):
        longest = max([len(label)] + [len(str(r.get(key) or "")) for r in rows])
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(longest + 3, 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.append([])
    ws.append([ATTRIBUTION])

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(
        out,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename("xlsx"),
    )


def filename(ext):
    return f"leads-{datetime.now():%Y%m%d-%H%M}.{ext}"


if __name__ == "__main__":
    app.run(debug=True, port=5000)
