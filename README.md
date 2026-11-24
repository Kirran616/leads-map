# Leadmap – OpenStreetMap Lead Generation Tool

Search businesses by keyword and city, view them in a dashboard, and export to CSV or Excel.
**No API key and no billing needed.** Data comes from OpenStreetMap (Nominatim + Overpass).

## Run it (Python 3.11.9)

```bash
python -m venv venv
venv\Scripts\activate          # Windows
# source venv/bin/activate     # macOS / Linux
pip install -r requirements.txt
copy .env.example .env         # Windows  (cp .env.example .env on macOS/Linux)
# edit .env and set OSM_CONTACT to your email
python app.py
```

Open http://127.0.0.1:5000

## How it works

1. **Nominatim** turns the city name into a bounding box (max 1 request per second, cached).
2. **Overpass** finds named businesses inside that box. Common words (restaurant, dentist, gym…) map to
   OpenStreetMap tags; any other keyword searches business names and types.
3. Results are de-duplicated, shown 25 per page, and can be exported.

## Fields

| Field | Source |
|---|---|
| Name, Category, Address, Phone, Website, Email | OSM tags (only when a mapper added them) |
| OpenStreetMap ID (e.g. `node/123`) | replaces Google's Place ID |
| Latitude / Longitude | OSM |

**Not available:** ratings and reviews. OpenStreetMap does not have them.

## OpenStreetMap rules to follow

- **Attribution:** data is © OpenStreetMap contributors under the ODbL. Exports include this credit; keep it.
  Share-alike applies if you publish a database built from this data. See https://www.openstreetmap.org/copyright
- **Identify your app:** set `OSM_CONTACT` in `.env` (used in the User-Agent header, required by the Nominatim and Overpass policies).
- **Fair use:** the public servers are free but shared. No bulk scraping or running many searches in a loop.
  For heavy use, host your own Overpass instance and set `OVERPASS_URL`.
- **Privacy / marketing laws:** phone numbers and emails of sole traders can be personal data (GDPR, CAN-SPAM, PECA etc.). Check the rules before contacting leads.

## Notes

- Overpass has no page tokens, so "pagination" is a result cap (50–500) plus client-side paging in the table.
- Areas larger than ~1.5° (regions, countries) are rejected to protect the public server.
- Spreadsheet cells starting with `= + - @` are escaped to prevent formula injection.
