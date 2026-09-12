# Travel Tracker

An Excel data store &amp; python backend based tool for displaying past travels such as cities &amp; countries visited, flight routes taken, etc. in interactive maps and dashboards.

## Synopsis

An interactive travel-tracking dashboard built inside an Excel workbook. Data lives in native Excel tables; a VBA layer hands the data off to a Python backend (using **xlwings** and **Folium**) which renders a set of interactive HTML pages — a dashboard, a yearly overview, a cities & countries map, and an aviation (flight routes) map — opened in your system's default browser.

## Features

- **Dashboard** — stat cards, top countries/cities, top flight routes, busiest years (toggle between # Countries / # New Countries)
- **Yearly Overview** — year-by-year breakdown of countries visited and new countries added
- **Cities & Countries Map** — Folium map with city markers (population-scaled) and country-level choropleth shading based on visit counts
- **Aviation Map** — visualization of flight routes between visited locations
- Country flags and outlines rendered both inside Excel and inside map tooltips/popups
- One-click generation: a control sheet with four buttons builds and launches each page in your browser

## Prerequisites

- **Microsoft Excel** (Windows OS required for full VBA support)
- **Python 3.x** runtime environment installed and available via your PATH environment variable
- Package **xlwings**:
  - `xlwings` add-in enabled in Excel (Excel Options → Add-ins, or run `xlwings addin install`)
- Various **pip packages** according to `requirements.txt`
- A free [GeoNames](https://www.geonames.org/login) account/username (used for population and lattitude / longitude lookups via GeoNames API)
- A free API key for [REST countries](http://restcountries.com) (used for names lookups via REST countries API)
- Internet connection (for API lookups)
- The Natural Earth countries shapefile/GeoJSON (used for country choropleth shading) — included, see [Project Structure](#project-structure)

## Installation

1. Install latest **python runtime** if not already installed, freely available here: [Download the latest version for Windows](https://www.python.org/downloads/)
2. Clone this repository:<br>
   Clone the **travel-tracker** git repository by opening a PowerShell or Git Bash console window, navigating to a directory of your chosing and running this command:
   ```bash
   git clone https://github.com/DanDanR/travel-tracker.git
   ```
3. Install dependencies:<br>
   Run the `setup.bat` script by double clicking or otherwise launching it, this might take a while as a virtual environment (or a **venv**) will be created and some pip packages are being installed.<br>
   After script execution has finished it should show the message 
   ```bash
   Installation of pip packages into venv completed successfully.
   ```
4. Enter your credentials:<br>
   After the setup script has finished there should be a file `credentials.txt`. Open it in a text editor and enter your username for **GeoNames** as well as your API key for **REST countries** and save the file.
5. Enable macro execution:<br>
   Open the workbook `TravelTracker.xlsm` in Excel and enable / trust content when prompted via popup or via banner underneath menu ribbon.

## Usage

Open `TravelTracker_blank.xlsm` (or rename it to your liking first) and navigate to the control sheet `Build` in the workbook. It contains four buttons, each of which builds the corresponding page and opens it in your default browser:

| Button | Output |
|---|---|
| Dashboard | Overview stats, top countries/cities, top routes, busiest years |
| Yearly Overview | Year-by-year travel breakdown |
| Cities & Countries Map | Interactive map with markers for visited cities and higlighting for visited countries. |
| Aviation Map | Flight route visualization of flights taken. |

Each button triggers a VBA macro that calls the Python backend via `xlwings`, regenerates the relevant HTML file, and opens it in your system's default browser.


### Populating Your Own Data

#### Important Notice

The various sheets contain custom formulas that are bound to VBA code, which makes API calls using your credentials (from `credentials.txt`). Some of these API calls are slow-performing. Re-calculating all formulas on the sheet and thusly re-running all API calls might take a long time. Also and the publically available free-to-use have daily usage limits and might block you for the rest of the day when you make too many calls (in the case that formulas no longer work, try again the next day!).<br><br>
Therefore it is advisable, that once you made some calls, to convert the formula-based cell contents to fixed-value cell contents. A custon shortcut for that purpose is available. Simply **press F11** and the formula of any cell will be replaced with the calculated value and thusly prevent further API calls for that cell.<br><br>
**This works on a single cell AND on multiple cells at the same time!**<br><br>


#### Table Cities

The sheet `Cities` contains a table with all the cities visited (each entry equals one unique city).

| Column | Meaning |
|---|---|
| ISO2 | ISO alpha-2 country code — will be auto-filled once you enter a country name in the cell to the right (must be correctly spelled common name according to country list on sheet `Countries`). |
| Country | Country's common name according to country list on sheet `Countries`. |
| Name | City's name (commonly internationally used English name preferred). |
| Alternative Name | Optional, e.g. local name or name in native-script. |
| Lat / Lon | Coordinates of the city (lattitude and longitude) — will be auto-filled once you've entered country and city name. |
| Population | City's population — will be auto-filled once you've entered country and city name. |
| Visits | Number of past visits — will be auto-filled based on the trip list on sheet `Trips`. |
| Home | Optional — tick the checkbox to mark a city as a city you currently or in the past live(d) or work(ed) in. That way entries for this city and its surrounding country will not be shown on the map as visits as visits to a home or work place do not make sense in the context of a travel tracker.  |
| First Visit / Last Visit | Month and Year of the first and last visit to this city — will be auto-filled based on the trip list on sheet `Trips`. |


#### Table Countries

The sheet `Countries` contains a pre-populated list of most sovereign countries and partially recognized territories (such as Western Sahara) as well as relevant overseas territories (such as Gibraltar).<br><br>
**All columns except for `ISO2` are auto-filled!**

The table comes fully populated and doesn't have to be touched at all. However, you can extend the table should there be some territories visited missing, find a complete list of all ISO alpha-2 codes here: [Officially assigned code elements](https://en.wikipedia.org/wiki/ISO_3166-1_alpha-2#Officially_assigned_code_elements)

| Column | Meaning |
|---|---|
| ISO2 | ISO alpha-2 country code |
| Common Name | The country's common short-form name, e.g. `China` instead of `People's Republic of China` — will be auto-filled. |
| Official Name | The country's long-form name, e.g `Russian Federation` instead of `Russia` — will be auto-filled. |
| Population | The country's population. |
| Area [km²] | The country's land area in square kilometers — will be auto-filled. |
| Capital | The country's capital city — will be auto-filled. |
| Visits | Unique visits to the country — will be auto-filled according to data coming from the sheet `Trips`. |
| First Visit / Last Visit | First visit and last visit to the country — will be auto-filled according to data coming from the sheet `Trips`. |
| Wiki Page | Link to the Wikipedia page of the country — will be auto-filled. |
| Flag | Auto-filled — will be pulled from the internet and embedded in the cell. |


#### Table Trips

The sheet `Trips` is the place where the visits and their details are specified. This table is the basis for calculations in various other sheets.<br> **Each row contains one city, NOT one trip!**

| Column | Meaning |
|---|---|
| Year | Year the trip took place. |
| Month | Month the trip took place (decimal). |
| Date | Will be auto-filled once Year and Month columns contain values, generates date strings and will for instance turn `2025` + `2` into `Feb 2025`. |
| Trip Name | Optional — for your convenience only, name trips as you like, trip names will not show up on any map or dahsboard. |
| Country | Common name of the country the trip included according to country list on sheet `Countries`. |
| City | Name of the city the trip included according to the city's definition on sheet `Cities`. |
| Comment | Optional — for your convenience only, will not show up on any map or dashboard. |


#### Table By Year

The sheet `By Year` is the basis for the pages `Dashboard` as well as `Visits By Years`.<br>
Here you only need to define what years to track, the entire rest will be auto-filled according to the data coming from other sheets.

| Column | Meaning |
|---|---|
| Year | Year to track — define yourself. |
| # Countries | Number of countries visited during that year — will be auto-filled. |
| # New Countries | Number of countries visited for the first time during that year — will be auto-filled. |
| Countries | List of countries visited that year — will be auto-filled. |
| New Countries | List of countries visited for the first time during that year — will be auto-filled. |

#### Table Airports

The sheet `Airports` holds all airports you've ever flown in or out of. One row equals one airport.<br>
**All columns except for `IATA Code` are auto-filled!**

| Column | Meaning |
|---|---|
| City | Name of the main city the airport is serving — will be auto-filled. |
| Full Name | The airport's full name including eponym / toponym, will be auto-filled. |
| ISO2 | ISO alpha-2 code of thecountry hosting the aiport — will be auto-filled. |
| IATA Code | The airport's IATA airport code, has to be manually entered. See here for reference: [List of airports](https://en.wikipedia.org/wiki/Lists_of_airports#By_IATA_code) |
| Lat / Lon | Coordinates of the airport (lattitude and longitude) — will be auto-filled. |
| Dpertures | Number of departing flights taken from this airport, based on data coming from the `Flights` sheet. |
| Arrivals | Number of arriving flights taken to this airport, based on data coming from the `Flights` sheet. |
| Wiki Page | Link to the airport's Wikipedia page — will be auto-filled. |


#### Table Flights


| Column | Meaning |
|---|---|
|  |  |

#### Table Airlines


| Column | Meaning |
|---|---|
|  |  |


2. Add one row per city you've visited. Lat/Lon and population can be looked up automatically via the GeoNames API integration if left blank (assuming your API username is configured).
3. Fill in visit counts per city — country-level totals are aggregated automatically from city-level visits.
4. If tracking flights, populate the flight/routes table (e.g. `Routes`) with origin/destination city pairs and dates.
5. If using the yearly breakdown, make sure your data includes a year/date field so trips can be grouped correctly.
6. Re-run the relevant button(s) on the control sheet to regenerate the pages with your updated data.

### Project Structure

```
travel-tracker/
├── TravelTracker.xlsm       # Main Excel workbook (data + VBA + control sheet)
├── map_generator.py         # Python script generating Folium maps
├── data/
│   └── countries.geojson    # Natural Earth admin-0 countries (for choropleth)
└── output/
    ├── dashboard.html
    ├── yearly_overview.html
    ├── map.html
    └── aviation_map.html
```

> Adjust file/sheet/button names above to match your actual workbook — this is a starting template.

### Notes

- Map rendering was intentionally moved out of an embedded browser control inside Excel (due to stability issues) in favor of generating standalone HTML files opened in the system browser.
- Country shading is computed by aggregating per-city visit counts up to the country level, joined against the Natural Earth GeoJSON via ISO country code.
