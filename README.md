# Travel Tracker

An Excel data store &amp; python backend based tool for displaying past travels such as cities &amp; countries visited, flight routes taken, etc. in interactive maps and dashboards.

## Synopsis

An interactive travel-tracking dashboard built inside an Excel workbook. Data lives in native Excel tables; a VBA layer hands the data off to a Python backend (using `xlwings` and `Folium`) which renders a set of interactive HTML pages — a dashboard, a yearly overview, a cities & countries map, and an aviation (flight routes) map — opened in your system's default browser.

## Features

- **Dashboard** — stat cards, top countries/cities, top flight routes, busiest years (toggle between # Countries / # New Countries)
- **Yearly Overview** — year-by-year breakdown of countries visited and new countries added
- **Cities & Countries Map** — Folium map with city markers (population-scaled) and country-level choropleth shading based on visit counts
- **Aviation Map** — visualization of flight routes between visited locations
- Country flags and outlines rendered both inside Excel (via `=IMAGE()`) and inside map tooltips/popups
- One-click generation: a control sheet with four buttons builds and launches each page in your browser

## Prerequisites

- **Microsoft Excel** (Windows OS required for full VBA support)
- **Python 3.x** installed and available via your PATH environment variable
- Python packages:
  ```bash
  pip install xlwings folium geopandas pandas
  ```
- `xlwings` add-in enabled in Excel (Excel Options → Add-ins, or run `xlwings addin install`)
- A free [GeoNames](https://www.geonames.org/login) account/username (used for city population and lat/lon lookups via their API)
- Internet connection (for GeoNames API lookups and [flagsapi.com](https://www.flagsapi.com/) flag images)
- The Natural Earth countries shapefile/GeoJSON (used for country choropleth shading) — see [Project Structure](#project-structure)

branca
certifi
charset-normalizer
folium
geopandas
idna
Jinja2
MarkupSafe
numpy
packaging
pandas
pyogrio
pyproj
python-dateutil
pywin32
requests
shapely
six
tzdata
urllib3
xlwings
xyzservices

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/<your-username>/<your-repo>.git
   ```
2. Open the workbook `TravelTracker.xlsm` in Excel.
3. Enable macros / content when prompted.
4. Update the GeoNames username inside the VBA code (or config sheet, if you add one) with your own.
5. Make sure the Python scripts (e.g. `map_generator.py`) and any supporting data files are in the paths the VBA macros expect — adjust file paths in the VBA code if your folder layout differs.

## Usage

Navigate to the control sheet (e.g. `Controls` / `Dashboard Buttons`) in the workbook. It contains four buttons, each of which builds the corresponding page and opens it in your default browser:

| Button | Output |
|---|---|
| Dashboard | Overview stats, top countries/cities, top routes, busiest years |
| Yearly Overview | Year-by-year travel breakdown |
| Cities & Countries Map | Interactive map with markers and choropleth shading |
| Aviation Map | Flight route visualization |

Each button triggers a VBA macro that calls into the Python backend via `xlwings`, regenerates the relevant HTML file, and opens it with your system's default browser.

## Populating Your Own Data

1. Open the main data sheet (e.g. `Data`), which contains a table (e.g. `MapData`) with columns such as:
   - `Name` — city name
   - `Lat` / `Lon` — coordinates
   - `Population`
   - `Visits`
   - `CountryCode`
   - `Alternative Name` — optional, e.g. native-script name for display in tooltips
2. Add one row per city you've visited. Lat/Lon and population can be looked up automatically via the GeoNames API integration if left blank (assuming your API username is configured).
3. Fill in visit counts per city — country-level totals are aggregated automatically from city-level visits.
4. If tracking flights, populate the flight/routes table (e.g. `Routes`) with origin/destination city pairs and dates.
5. If using the yearly breakdown, make sure your data includes a year/date field so trips can be grouped correctly.
6. Re-run the relevant button(s) on the control sheet to regenerate the pages with your updated data.

## Project Structure

```
travel-tracker/
├── TravelTracker.xlsm       # Main Excel workbook (data + VBA + control sheet)
├── map_generator.py         # Python script generating Folium maps
├── build_dashboard.py       # Python script generating the dashboard page
├── data/
│   └── countries.geojson    # Natural Earth admin-0 countries (for choropleth)
└── output/
    ├── dashboard.html
    ├── yearly_overview.html
    ├── map.html
    └── aviation_map.html
```

> Adjust file/sheet/button names above to match your actual workbook — this is a starting template.

## Notes

- Map rendering was intentionally moved out of an embedded browser control inside Excel (due to stability issues) in favor of generating standalone HTML files opened in the system browser.
- Country shading is computed by aggregating per-city visit counts up to the country level, joined against the Natural Earth GeoJSON via ISO country code.

## License

Add your preferred license here (e.g. MIT).