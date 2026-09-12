#region imports
import time
import os
import json
import math
import requests
import xlwings as xw
import pandas as pd
import geopandas as gpd
import folium
import folium.plugins
import requests
import functools
from shapely.geometry import MultiPolygon, Polygon
from shapely.geometry import shape
from pyproj import Geod
import branca.colormap as cm

#endregion

#region globals

FOURSQUARE_API_KEY = ""
GOOGLE_API_KEY = ""
# https://github.com/foursquare/foursquare-places-api-samples/tree/main/places-api
# https://docs.foursquare.com/data-products/docs/places-pro-and-premium

CITY_SHEET = "Cities"
COUNTRY_SHEET = "Countries"
AIRPORT_SHEET = "Airports"
FLIGHT_SHEET = "Flights"
AIRLINE_SHEET = "Airlines"
BY_YEAR_SHEET = "By Year"
POI_SHEET = "POIs"

CITY_TABLE = "CityTable"
COUNTRY_TABLE = "CountryTable"
AIPORT_TABLE = "AirportTable"
FLIGHT_TABLE = "FlightTable"
AIRLINE_TABLE = "AirlineTable"
BY_YEAR_TABLE = "ByYearTable"
POI_TABLE = "PoiTable"

MIN_RADIUS = 3
MAX_RADIUS = 12
MIN_CONTOUR_WEIGHT = 0.5
MAX_CONTOUR_WEIGHT = 2
MIN_LINE_WEIGHT = 2.5
MAX_LINE_WEIGHT = 7
POP_CONTRAST = 2.5
POP_FLOOR = 10000
DASHBOARD_TOP_SLICE = 5

ARRIVAL_EMOJI = "🛬"
DEPARTURE_EMOJI = "🛫"

CITY_COUNTRY_HTML = "travel-map.html"
AVIATION_HTML = "aviation-map.html"
DASHBOARD_HTML = "travel-dashboard.html"
BY_YEAR_HTML = "yearly-overview.html"
POI_HTML = "POIs.html"
COUNTRIES_GEOJSON = "world-countries-50m.json"
COUNTRIES_SHAPEFILE = "ne_50m_admin_0_countries.shp"

_geod = Geod(ellps="WGS84")

_base_path = ""
_output_dir = ""
_workbook_name = ""
_google_api_session_token = ""

#endregion

#region experimental

def _create_google_api_session(lang_code):
    global _google_api_session_token
    session_response = requests.post(
        f"https://tile.googleapis.com/v1/createSession?key={GOOGLE_API_KEY}",
        json={"mapType": "roadmap", "language": lang_code, "region": "DE"},
    )

    session_response.raise_for_status()
    _google_api_session_token = session_response.json()["session"]

def _read_poi_table(wb):
    sheet = wb.sheets[POI_SHEET]
    table = sheet.tables[POI_TABLE]
    return table.range.options(pd.DataFrame, header=1, index=False).value

def _pre_populate_pois(poi_df):
    pass

def _add_poi_markers(fmap, poi_df, show_visited=True, show_planned=True, pre_populate=True):
    pois = {}
    if pre_populate:
        _pre_populate_pois(poi_df, pois)

    for _, row in poi_df.iterrows():
        name = row.get("Name", "")
        alt_name = row.get("Alternative Name", "")
        if pd.notna(alt_name) and str(alt_name).strip() != "":
            full_name = f"{name} ({alt_name})"
        else:
            full_name = name
        visited = _is_visited(row.get("Been", ""))
        planned = not visited
        rank = pd.to_numeric(row.get("Rank"), errors="coerce")
        #radius = _population_to_radius(population, min_pop, max_pop)
        #weight = _population_to_weight(population, min_pop, max_pop)

        tooltip_html = f"""
                <br><b>{full_name}</b>
                <br>Status: {"visited" if visited else "planned"}
                <br>Rank: {rank:,.2f}
                <br><a href="https://en.wikipedia.org/wiki/{name.replace(' ', '_')}" target="_blank">Wikipedia</a>
            """

        if (show_visited) and visited or (show_planned and planned):
            folium.CircleMarker(
                location=[row["Lat"], row["Lon"]],
                radius=6,
                weight=2,
                color="#f00707e6" if planned else "#2b6cb0",
                fill=True,
                fill_color="#dd1212b9" if planned else "#4299e1",
                fill_opacity=0.85,
                tooltip=folium.Tooltip(tooltip_html),
                popup=folium.Popup(tooltip_html, max_width=250),
            ).add_to(fmap)

def build_poi_map(open_browser=True, lang_code="en"):
    api_key = ""
    id = ""
    google_places_base_url = f"https://places.googleapis.com/v1/places/{id}?fields=addressComponents&key={google_places_base_url}"

    wb = xw.Book.caller()
    _set_paths(wb)

    poi_df = _read_poi_table(wb)
    poi_df = poi_df.dropna(subset=["Lat", "Lon"])
    
    geojson_data = _get_geojson_data()
    
    fmap = _create_base_map(poi_df, lang_code)
    _add_country_choropleth(fmap, poi_df, geojson_data, _get_exclusion_countries())
    _add_country_tooltips(fmap, poi_df, geojson_data, _get_exclusion_countries())
    _add_poi_markers(fmap, poi_df)
    
    _save_map_to_html(fmap, POI_HTML)

#endregion

#region helper functions

def _set_paths(wb):
    global _workbook_name, _base_path,_output_dir

    _workbook_name = wb.name.rsplit(".", 1)[0]
    _base_path = wb.fullname.rsplit("\\", 1)[0]
    _output_dir = os.path.join(_base_path, "output")

def _save_dashboard_to_html(html, filename, launch_browser = True):
    if not os.path.exists(_output_dir):
        os.makedirs(_output_dir)

    output = os.path.join(_output_dir, f"{_workbook_name}_{filename}")
    
    with open(output, "w", encoding="utf-8") as f:
        f.write(html)
    
    if launch_browser:
        # Open in default browser (Windows)
        os.startfile(output)

def _save_map_to_html(fmap, filename, launch_browser = True):
    if not os.path.exists(_output_dir):
        os.makedirs(_output_dir)

    output = os.path.join(_base_path, f"aaa")
    os.makedirs(os.path.dirname(output), exist_ok=True)
    fmap.save(output)
    
    if launch_browser:
        # Open in default browser (Windows)
        os.startfile(output)

def _is_visited(visited_str):
    visited = visited_str.lower()
    if visited == "y" or visited == "yes" or visited == "j" or visited == "ja" or visited == "t" or visited == "true" or visited == "1":
        return True
    else:
        return False

@xw.func
def get_population(country_name):
    """Excel formula: =get_population("Germany")"""
    resp = requests.get(f"https://restcountries.com/v3.1/name/{country_name}?fields=population")
    resp.raise_for_status()
    data = resp.json()
    return data[0]["population"]

@functools.lru_cache(maxsize=None)
def _fetch_wikipedia_image(page_title, assemble_from_wiki_page_url = False):
    """
    Returns the URL of a Wikipedia page's lead/thumbnail image, or None
    if the page doesn't exist or has no image.
    """
    if assemble_from_wiki_page_url:
        page_title = page_title.removeprefix("https://en.wikipedia.org/wiki/")
        url = f"https://en.wikipedia.org/api/rest_v1/page/summary/{page_title.replace(' ', '_')}"
    else:
        url = f"https://en.wikipedia.org/api/rest_v1/page/summary/{page_title.replace(' ', '_')}"

    try:
        resp = requests.get(url, headers={"User-Agent": "TravelTracker/1.0"}, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        return data.get("thumbnail", {}).get("source")
    except (requests.RequestException, ValueError):
        return None

def _get_flag_url(iso2):
    # exception rule for Northern Cyprus!
    if iso2 == "NY" or iso2 == "NC":
        return "https://upload.wikimedia.org/wikipedia/commons/thumb/1/1e/Flag_of_the_Turkish_Republic_of_Northern_Cyprus.svg/1280px-Flag_of_the_Turkish_Republic_of_Northern_Cyprus.svg.png"
    else:
        return f"https://flagsapi.com/{iso2}/flat/64.png"

def _build_country_svg_lookup(iso2_codes, flag_in_shape):
    geojson_data = _get_geojson_data()
    wanted = set(iso2_codes)
    lookup = {}
    for feature in geojson_data["features"]:
        props = feature["properties"]
        iso2 = props.get("ISO_A2")
        if iso2 not in wanted:
            continue
        geometry = shape(feature["geometry"])
        geometry = _get_main_landmass_only(geometry, iso2)
        lookup[iso2] = _country_shape_svg(geometry, iso2, "", flag_in_shape)
    return lookup

def _country_shape_svg(geometry, iso2, common_name="", flag_in_shape=False, size=120, stroke="#333", fill="#ddd"):
    polygons = geometry.geoms if isinstance(geometry, MultiPolygon) else [geometry]
    minx, miny, maxx, maxy = geometry.bounds
    width_span = maxx - minx or 1
    height_span = maxy - miny or 1
    scale = size / max(width_span, height_span)
    svg_height = height_span * scale

    def to_svg(x, y):
        # flip y: SVG y grows downward, latitude grows upward
        return (x - minx) * scale, (maxy - y) * scale

    if flag_in_shape:
        clip_id=f"clip-{iso2}"
        clip_shapes = []
        for poly in polygons:
            pts = " ".join(f"{sx:.1f},{sy:.1f}" for sx, sy in (to_svg(x, y) for x, y in poly.exterior.coords))
            clip_shapes.append(f'<polygon points="{pts}"/>')

    outline_shapes = []
    for poly in polygons:
        pts = " ".join(f"{sx:.1f},{sy:.1f}" for sx, sy in (to_svg(x, y) for x, y in poly.exterior.coords))
        outline_shapes.append(f'<polygon points="{pts}" fill="{'none' if flag_in_shape else fill}" stroke="{stroke}" stroke-width="1"/>')

    if flag_in_shape:
        return (f'<svg viewBox="0 0 {size:.1f} {svg_height:.1f}" width="{size}" height="{svg_height:.0f}" '
                f'preserveAspectRatio="xMidYMid meet" '
                f'xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"> '
                f'<title>{common_name}</title>'
                f'<defs><clipPath id="{clip_id}">{"".join(clip_shapes)}</clipPath></defs> '
                f'<image href="{_get_flag_url(iso2)}" xlink:href="{_get_flag_url(iso2)}" x="0" y="0" width="{size:.1f}" height="{svg_height:.1f}" '
                f'preserveAspectRatio="xMidYMid slice" clip-path="url(#{clip_id})"/>{"".join(outline_shapes)}</svg>')
    else:
        return (f'<svg viewBox="0 0 {size:.1f} {svg_height:.1f}" width="{size}" height="{svg_height:.0f}" '
                f'preserveAspectRatio="xMidYMid meet" '
                f'xmlns="http://www.w3.org/2000/svg">{"".join(outline_shapes)}</svg>')

def _get_main_landmass_only(geometry, iso_a2):
    """Drop small, disconnected polygon parts (e.g. distant overseas territories),
    keeping only parts at least `area_ratio_threshold` the size of the largest part."""

    area_ratio_threshold = 0
    problem_countries = [["FR", 0.15], ["NL", 0.15], ["US", 0.15], ["ES", 0.15], ["RU", 0.15], ["PT", 0.1], 
                         ["NZ", 0.1], ["DK", 0.1], ["TW", 0.1], ["NO", 0.3], ["KR", 0.2], ["JP", 0.05]]

    if not isinstance(geometry, MultiPolygon):
        return geometry  # already a single Polygon, nothing to filter
    
    for i in range(len(problem_countries)):
        if problem_countries[i][0] == iso_a2:
            area_ratio_threshold = problem_countries[i][1]
            break

    if area_ratio_threshold == 0:
        return geometry

    polygons = list(geometry.geoms)
    max_area = max(p.area for p in polygons)
    kept = [p for p in polygons if p.area >= max_area * area_ratio_threshold]

    return kept[0] if len(kept) == 1 else MultiPolygon(kept)

def _population_to_radius(population, min_pop, max_pop):
    fraction = _population_fraction(population, min_pop, max_pop)
    return MIN_RADIUS + fraction * (MAX_RADIUS - MIN_RADIUS)

def _population_to_weight(population, min_pop, max_pop):
    fraction = _population_fraction(population, min_pop, max_pop)
    return MIN_CONTOUR_WEIGHT + fraction * (MAX_CONTOUR_WEIGHT - MIN_CONTOUR_WEIGHT)

def _population_fraction(population, min_pop, max_pop):
    if population is None or population <= 0:
        return 0

    log_pop = math.log10(population)
    log_min = math.log10(max(min_pop, 1))
    log_max = math.log10(max(max_pop, 1))

    if log_max == log_min:
        return 0.5

    fraction = (log_pop - log_min) / (log_max - log_min)
    fraction = min(max(fraction, 0), 1)

    return fraction ** POP_CONTRAST

def _get_exclusion_countries():
    raw = xw.sheets["HomeOrWork"].range("A1").expand("down").value
    if raw is None:
        exclusion_countries = set()
    elif isinstance(raw, str):
        exclusion_countries = {raw}
    else:
        exclusion_countries = set(raw)

    return exclusion_countries

def _get_geojson_data():
    with open(os.path.join(_base_path, COUNTRIES_GEOJSON), encoding="utf-8") as f:
        geojson_data = json.load(f)
    return geojson_data

def _aggregate_flight_count_directionless(flights_df):
    df = flights_df.dropna(subset=["Origin IATA", "Destination IATA"]).copy()

    # sort each row's pair alphabetically so A→B and B→A become the same group
    df["RouteKey"] = df.apply(
        lambda row: tuple(sorted([row["Origin IATA"], row["Destination IATA"]])), axis=1
    )

    routes_df = df.groupby("RouteKey").size().reset_index(name="Flights")
    routes_df["Origin IATA"] = routes_df["RouteKey"].apply(lambda pair: pair[0])
    routes_df["Destination IATA"] = routes_df["RouteKey"].apply(lambda pair: pair[1])
    routes_df = routes_df.drop(columns=["RouteKey"])

    return routes_df

def _aggregate_flight_count_directional(flights_df):
    routes_df = (
        flights_df.dropna(subset=["Origin IATA", "Destination IATA"])
        .groupby(["Origin IATA", "Destination IATA"])
        .size()
        .reset_index(name="Flights")
    )
    return routes_df

def _great_circle_path(lat1, lon1, lat2, lon2, n_points=100):
    intermediate = _geod.npts(lon1, lat1, lon2, lat2, n_points)
    full_lonlat = [(lon1, lat1)] + intermediate + [(lon2, lat2)]

    segments = []
    current_segment = [[full_lonlat[0][1], full_lonlat[0][0]]]  # [lat, lon]

    for i in range(1, len(full_lonlat)):
        prev_lon = full_lonlat[i - 1][0]
        curr_lon = full_lonlat[i][0]

        if abs(curr_lon - prev_lon) > 180:
            # antimeridian crossing — end this segment, start a new one
            segments.append(current_segment)
            current_segment = []

        current_segment.append([full_lonlat[i][1], full_lonlat[i][0]])

    segments.append(current_segment)
    return segments

def _flights_to_weight(flights, min_flights, max_flights):
    if flights is None or max_flights == min_flights:
        return MIN_LINE_WEIGHT

    fraction = (flights - min_flights) / (max_flights - min_flights)
    fraction = min(max(fraction, 0), 1)

    return MIN_LINE_WEIGHT + fraction * (MAX_LINE_WEIGHT - MIN_LINE_WEIGHT)

#endregion

#region map functions

def _create_base_map(city_df, lang_code):
    if city_df is None or city_df.empty:
        center = [30, 0]
        zoom = 2
    else:
        center = [city_df["Lat"].mean(), city_df["Lon"].mean()]
        zoom = 3

    # basic map w/o any API requirements and no possibility for dynamic translation:
    #map = folium.Map(location=center, zoom_start=zoom, tiles="cartodbpositron")

    # with dynamic translation, subject to API changes or blocks due to excessive usage:
    map = folium.Map(location=[30, 0], zoom_start=2, tiles=None)
    folium.TileLayer(
        tiles=f"https://mt{{s}}.google.com/vt/lyrs=m&x={{x}}&y={{y}}&z={{z}}&hl={lang_code}",
        attr="Google",
        subdomains=["0", "1", "2", "3"],
        name="Google Roads (zh)",
        max_zoom=20,
    ).add_to(map)

    return map

def _add_country_choropleth(fmap, country_df, geojson_data, exclusion_countries):
    # Colormap scale excludes home countries entirely — they never render on the spectrum
    color_scale_df = country_df[
        ~country_df["ISO2"].isin(exclusion_countries) & country_df["Visits"].notna()
    ]
    colormap = cm.linear.YlOrRd_09.scale(
        color_scale_df["Visits"].min(), color_scale_df["Visits"].max()
    )

    choropleth = folium.Choropleth(
        geo_data=geojson_data,
        data=country_df,
        columns=["ISO2", "Visits"],
        key_on="feature.properties.ISO_A2",
        fill_color="YlOrRd",
        fill_opacity=0.7,
        line_opacity=0.7,
        nan_fill_color="lightgrey",         # countries with no matching row
        legend_name="Visits to country",
        name="Country shading",
        highlight=True
    ).add_to(fmap)

    value_by_iso2 = dict(zip(country_df["ISO2"], country_df["Visits"]))

    def country_style_function(feature):
        iso2 = feature['properties'].get('ISO_A2')
        value = value_by_iso2.get(iso2)

        if iso2 in exclusion_countries:
            fill_color = "#202320"
        elif value is None or (isinstance(value, float) and math.isnan(value)):
            fill_color = "lightgrey"
        else:
            fill_color = colormap(value)

        return {
            'fillColor': fill_color,
            'color': 'black',
            'weight': 1,
            'fillOpacity': 0.7,
        }

    choropleth.geojson.style_function = country_style_function

def _add_country_tooltips(fmap, country_df, geojson_data, exclusion_countries):
    #value_by_code = dict(zip(country_df["ISO2"], [country_df["Visits"], country_df["First Visit"], country_df["Last Visit"]]))

    value_by_code = dict(zip(country_df["ISO2"], zip(country_df["Visits"], country_df["First Visit"], country_df["Last Visit"])))

    for feature in geojson_data["features"]:
        code = feature["properties"]["ISO_A2"]
        name = feature["properties"]["NAME_EN"]
        value = value_by_code.get(code, (0, None, None))[0]
        try:
            visits = int(value)
        except:
            visits = value
        try:
            population = int(feature["properties"]["POP_EST"])
        except:
            population = "-"

        feature["properties"]["Value"] = value  # keep for the hover tooltip

        geometry = shape(feature["geometry"])  # converts the raw GeoJSON dict into a Shapely object
        geometry = _get_main_landmass_only(geometry, code)
        svg = _country_shape_svg(geometry, code, common_name="", flag_in_shape=False)

        if code in exclusion_countries:
            visits_string = f"live(d) there!"
        elif visits < 1:
            visits_string = "0"
        else:
            visits_string = f"{visits:,.0f} (first: {value_by_code.get(code, (0, None, None))[1]} / last: {value_by_code.get(code, (0, None, None))[2]})"

        feature["properties"]["VisitsString"] = visits_string

        feature["properties"]["PopupHTML"] = f"""
            <div style="width:220px">
                <div style="display:flex; align-items:center; gap:12px; margin-top:4px;">{svg}
                    <img src="{_get_flag_url(code)}" style="max-height:60px; border-radius:2px; border:1px solid #ccc;">
                </div>
                <br><b>{name}</b>
                <br>Population: {population:,.0f}
                <br>Visits: {visits_string}
                <br><a href="https://en.wikipedia.org/wiki/{name.replace(' ', '_')}" target="_blank">Wikipedia</a>
            </div>
        """
    
    folium.GeoJson(
        geojson_data,
        style_function=lambda f: {"fillOpacity": 0, "weight": 0},  # invisible
        highlight_function = lambda x: {'fillOpacity': 0.25, 'fillColor': 'blue', 'color': 'yellow', 'weight': 3},
        tooltip=folium.GeoJsonTooltip(
            fields=["FORMAL_EN", "VisitsString"],   # "name" = this GeoJSON's own country-name property
            aliases=["Country:", "Visited:"],
        ),
        popup=folium.GeoJsonPopup(
            fields=["PopupHTML"],
            labels=False,     # <-- critical: without this, Folium prefixes "PopupHTML: " before your HTML
            localize=True,
        ),
    ).add_to(fmap)

def _add_city_markers(fmap, city_df):
    if "Population" in city_df.columns:
        pop_series = pd.to_numeric(city_df["Population"], errors="coerce").dropna()
        min_pop = pop_series.min() if not pop_series.empty else POP_FLOOR
        max_pop = pop_series.max() if not pop_series.empty else POP_FLOOR
    else:
        min_pop = max_pop = POP_FLOOR

    home_cities = []
    for _, row in city_df.iterrows():
        name = row.get("Name", "")
        wiki_image_url = _fetch_wikipedia_image(name)
        alternative_name = row.get("Alternative Name", "")
        is_home_or_work = row.get("Home")
        if pd.notna(alternative_name) and str(alternative_name).strip() != "":
            full_name = f"{name} ({alternative_name})"
        else:
            full_name = name
        population = row.get("Population", "")
        visits = row.get("Visits", "")

        if is_home_or_work:
            visits_string = f"live(d) or work(ed) there!"
        else:
            visits_string = f"{visits:,.0f} (first: {row.get("First Visit", "")} / last: {row.get("Last Visit", "")})"

        image_tag = f'<img src="{wiki_image_url}" width="200">' if wiki_image_url else ""
        tooltip_html = f"""
            <b>{image_tag}</b>
            <br><b>{full_name}</b>
            <br>Population: {population:,.0f}
            <br>Visits: {visits_string}
            <br><a href="https://en.wikipedia.org/wiki/{name.replace(' ', '_')}" target="_blank">Wikipedia</a>
        """

        population = pd.to_numeric(row.get("Population"), errors="coerce")
        radius = _population_to_radius(population, min_pop, max_pop)
        weight = _population_to_weight(population, min_pop, max_pop)

        if is_home_or_work:
            color = "#e4f007e6"
            fill_color = "#dd1212b9"
        else:
            color = "#2b6cb0"
            fill_color = "#4299e1"

        city_marker = folium.CircleMarker(
            location=[row["Lat"], row["Lon"]],
            radius=radius,
            weight=weight,
            color=color,
            fill=True,
            fill_color=fill_color,
            fill_opacity=0.85,
            tooltip=folium.Tooltip(tooltip_html),
            popup=folium.Popup(tooltip_html, max_width=250),
        )

        if is_home_or_work:
                # remember & add to map last for purpose of top z-layer position
                home_cities.append(city_marker)
        else:
            city_marker.add_to(fmap)

    for city in home_cities:
        city.add_to(fmap)

def _add_airport_markers(fmap, aiport_df):
    radius = 6
    weight = 2

    for _, row in aiport_df.iterrows():
        name = row.get("Full Name", "")
        wiki_page_url = row.get("Wiki Page", "")
        wiki_image_url = _fetch_wikipedia_image(wiki_page_url, True)
        iata = row.get("IATA Code", "")
        city = row.get("City", "")
        country = row.get("Country Code", "")
        departures = row.get("Departures", "")
        arrivals = row.get("Arrivals", "")

        image_tag = f'<img src="{wiki_image_url}" width="200">' if wiki_image_url else ""
        tooltip_html = f"""
            <b>{image_tag}</b><br>
            <b>{name} ({iata})</b>
            <br>City: {city}
            <br>Country: {country}
            <br>{DEPARTURE_EMOJI}: {departures:,.0f}
            <br>{ARRIVAL_EMOJI}: {arrivals:,.0f}
            <br><a href="https://en.wikipedia.org/wiki/{name.replace(' ', '_')}" target="_blank">Wikipedia</a>
        """

        folium.CircleMarker(
            location=[row["Lat"], row["Lon"]],
            radius=radius,
            weight=weight,
            color="#2b6cb0",
            fill=True,
            fill_color="#4299e1",
            fill_opacity=0.85,
            tooltip=folium.Tooltip(tooltip_html),
            popup=folium.Popup(tooltip_html, max_width=250),
        ).add_to(fmap)

def _add_flight_routes(fmap, flights_df, airport_df, airlines_df):
    routes_df = _aggregate_flight_count_directional(flights_df)

    # TBD: fix this!
    min_flights = 1 #flights_series.min() if not flights_series.empty else 1
    max_flights = 10 #flights_series.max() if not flights_series.empty else 1

    for _, row in routes_df.iterrows():
        origin_iata = row.get("Origin IATA")
        destination_iata = row.get("Destination IATA")

        matchOrigin = airport_df[airport_df["IATA Code"] == origin_iata]
        matchDestination = airport_df[airport_df["IATA Code"] == destination_iata]
        matchRoute = routes_df[(routes_df["Origin IATA"] == origin_iata) & (routes_df["Destination IATA"] == destination_iata)]

        if not matchOrigin.empty:
            origin_lat = matchOrigin.iloc[0]["Lat"]
            origin_lon = matchOrigin.iloc[0]["Lon"]
            origin_city = matchOrigin.iloc[0]["City"]

        if not matchDestination.empty:
            destination_lat = matchDestination.iloc[0]["Lat"]
            destination_lon = matchDestination.iloc[0]["Lon"]
            destination_city = matchDestination.iloc[0]["City"]

        number_of_flights = matchRoute.iloc[0]["Flights"] if not matchRoute.empty else 0
        weight = _flights_to_weight(number_of_flights, min_flights, max_flights)

        segments = _great_circle_path(origin_lat, origin_lon, destination_lat, destination_lon)

        if number_of_flights == 1:
            matchFlight = flights_df[(flights_df["Origin IATA"] == origin_iata) & (flights_df["Destination IATA"] == destination_iata)]
            trip_string = f"<br>Trip: {matchFlight.iloc[0]["Trip Name"]}"
            date_string = f"<br>Date: {matchFlight.iloc[0]["Month"]:.0f}/{matchFlight.iloc[0]["Year"]:.0f}"
            airline_string = f"<br>Airline: {matchFlight.iloc[0]["Airline"]}"

            matchAirline = airlines_df[airlines_df["ICAO"] == matchFlight.iloc[0]["ICAO"]]
            if not matchAirline.empty:
                logo_url = matchAirline.iloc[0]["Logo URL"]
                logo_tag = f"<img src='{logo_url}' style='max-width:120px; max-height:60px; height:auto; width:auto;'>"
            else:
                logo_tag = "🗺️⁀જ✈︎"
        else:
            # TBD: implemnt date first flight + date last flight + list of airlines + list of trips + list of overlapping logos!"
            trip_string = ""
            date_string = ""
            airline_string = ""
            logo_tag = "🛫✈️🛬"

        popup_html = f"""
        <b>{logo_tag}</b><br>
        <br>Route: {origin_city} ({origin_iata})  ->  {destination_city} ({destination_iata})
        <br>Flight #: {number_of_flights:,.0f} flight{"" if number_of_flights == 1 else "s"}
        {trip_string}
        {date_string}
        {airline_string}
        """

        for segment in segments:
            polyline = folium.PolyLine(
                locations=segment,
                color="#067df4",
                weight=weight,
                opacity=0.7,
                tooltip=f"{origin_city} ({origin_iata}) → {destination_city} ({destination_iata}) [{number_of_flights:.0f} flights]",
                popup=folium.Popup(popup_html, max_width=350),
            ).add_to(fmap)

#endregion

#region read functions

def _read_city_table(wb):
    sheet = wb.sheets[CITY_SHEET]
    table = sheet.tables[CITY_TABLE]
    return table.range.options(pd.DataFrame, header=1, index=False).value

def _read_country_table(wb):
    sheet = wb.sheets[COUNTRY_SHEET]
    table = sheet.tables[COUNTRY_TABLE]
    return table.range.options(pd.DataFrame, header=1, index=False).value

def _read_airport_table(wb):
    sheet = wb.sheets[AIRPORT_SHEET]
    table = sheet.tables[AIPORT_TABLE]
    return table.range.options(pd.DataFrame, header=1, index=False).value

def _read_flight_table(wb):
    sheet = wb.sheets[FLIGHT_SHEET]
    table = sheet.tables[FLIGHT_TABLE]
    return table.range.options(pd.DataFrame, header=1, index=False).value

def _read_airline_table(wb):
    sheet = wb.sheets[AIRLINE_SHEET]
    table = sheet.tables[AIRLINE_TABLE]
    return table.range.options(pd.DataFrame, header=1, index=False).value

def _read_year_table(wb):
    sheet = wb.sheets[BY_YEAR_SHEET]
    table = sheet.tables[BY_YEAR_TABLE]
    df = table.range.options(pd.DataFrame, index=False, header=1).value
    df["Year"] = df["Year"].astype(int)  # avoids "2024.0" showing up as a label
    return df

#endregion

#region page functions

def _render_dashboard_html(stats, top_countries_df, top_cities_df, top_routes_df,
                            top_years_by_countries_df, top_years_by_new_df,
                            all_countries_df=None):

    def _bar_rows(df, label_col, value_col, value_suffix):
        if df.empty:
            return '<p style="color:#888; font-size:13px;">No data yet.</p>'

        max_value = df[value_col].max()
        rows = []
        for _, row in df.iterrows():
            label = row[label_col]
            value = row[value_col]
            pct = (value / max_value * 100) if max_value else 0
            rows.append(f"""
            <div style="margin-bottom:10px;">
              <div style="display:flex; justify-content:space-between; font-size:13px; margin-bottom:4px;">
                <span>{label}</span>
                <span style="color:#888;">{value:.0f} {value_suffix}</span>
              </div>
              <div style="background:#eee; border-radius:4px; height:6px;">
                <div style="background:#2b6cb0; width:{pct:.0f}%; height:6px; border-radius:4px;"></div>
              </div>
            </div>
            """)
        return "".join(rows)

    countries_rows = _bar_rows(top_countries_df, "Common Name", "Visits", "visits")
    cities_rows = _bar_rows(top_cities_df, "Name", "Visits", "visits")
    years_by_countries_rows = _bar_rows(top_years_by_countries_df, "Year", "# Countries", "countries")
    years_by_new_rows = _bar_rows(top_years_by_new_df, "Year", "# New Countries", "new countries")

    if top_routes_df.empty:
        routes_rows = '<p style="color:#888; font-size:13px;">No data yet.</p>'
    else:
        route_lines = []
        for _, row in top_routes_df.iterrows():
            route_lines.append(f"""
            <tr style="border-top:1px solid #eee;">
              <td style="padding:8px 0; color:#333;">{row['Origin City']} &rarr; {row['Destination City']}</td>
              <td style="padding:8px 0; text-align:right; color:#888;">{row['Flights']:.0f} flights</td>
            </tr>
            """)
        routes_rows = f'<table style="width:100%; border-collapse:collapse; font-size:13px;">{"".join(route_lines)}</table>'

    geojson_data = _get_geojson_data()
    wanted_codes = set(all_countries_df["ISO2"].dropna())
    iso2_to_name = dict(zip(all_countries_df["ISO2"], all_countries_df["Common Name"]))
    country_svg_by_iso2 = {}

    for feature in geojson_data["features"]:
        props = feature["properties"]
        iso2 = props.get("ISO_A2")
        common_name = iso2_to_name.get(iso2, iso2)

        if iso2 not in wanted_codes:
            continue

        geometry = shape(feature["geometry"])
        geometry = _get_main_landmass_only(geometry, iso2)
        country_svg_by_iso2[iso2] = _country_shape_svg(geometry, iso2, common_name, True)

    if all_countries_df is not None and not all_countries_df.empty:
        chips = "".join(
            f'''<div class="country-item">
                <img class="flag-view" src="{_get_flag_url(code)}" width="64" title="{iso2_to_name.get(code, code)}" style="border-radius:2px; border:1px solid #ccc;">
                <div class="outline-view">{country_svg_by_iso2.get(code, "")}</div>
                </div>'''
            for code in all_countries_df["ISO2"].dropna()
        )
        flag_strip = f"""
            <style>
                .country-item {{ display: inline-flex; align-items: center; }}
                .outline-view {{ display: none; }}
                #country-chips.show-outlines .outline-view {{ display: block; }}
                #country-chips.show-outlines .flag-view {{ display: none; }}
            </style>
            <div style="background:#fff; border:1px solid #eee; border-radius:12px; padding:1rem 1.25rem; margin-top:1.5rem;">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                    <h3 style="margin:0; font-size:16px; font-weight:500;">Countries & Territories Visited</h3>
                    <button id="toggle-view-btn" style="font-size:12px; padding:4px 10px; border-radius:6px; border:1px solid #ddd; background:#f7f7f7; cursor:pointer;">Show outlines</button>
                </div>
                <div id="country-chips" style="display:flex; flex-wrap:wrap; gap:8px;">{chips}</div>
            </div>
            <script>
                document.getElementById('toggle-view-btn').addEventListener('click', function() {{
                    const row = document.getElementById('country-chips');
                    row.classList.toggle('show-outlines');
                    this.textContent = row.classList.contains('show-outlines') ? 'Show flags' : 'Show outlines';
                }});
            </script>
        """
    else:
        flag_strip = ""

    years_panel = f"""
        <div class="panel">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                <h3 style="margin:0;">Busiest Years</h3>
                <div style="display:flex; gap:10px; font-size:12px; color:#555;">
                    <label style="display:flex; align-items:center; gap:4px; cursor:pointer;">
                        <input type="radio" name="year-metric" value="countries" checked> Countries
                    </label>
                    <label style="display:flex; align-items:center; gap:4px; cursor:pointer;">
                        <input type="radio" name="year-metric" value="new"> New Countries
                    </label>
                </div>
            </div>
            <div id="years-by-countries">{years_by_countries_rows}</div>
            <div id="years-by-new" style="display:none;">{years_by_new_rows}</div>
        </div>
        <script>
            document.querySelectorAll('input[name="year-metric"]').forEach(function(radio) {{
                radio.addEventListener('change', function() {{
                    const showNew = this.value === 'new';
                    document.getElementById('years-by-countries').style.display = showNew ? 'none' : 'block';
                    document.getElementById('years-by-new').style.display = showNew ? 'block' : 'none';
                }});
            }});
        </script>
        """
    
    countries_this_year_card = ""
    if "countries_this_year" in stats:
        countries_this_year_card = f"""
        <div style="background:#e6f1fb; border-radius:8px; padding:1rem;">
          <p style="font-size:13px; color:#185fa5; margin:0 0 4px;">Countries this year</p>
          <p style="font-size:24px; font-weight:500; margin:0; color:#185fa5;">{stats['countries_this_year']:.0f}</p>
        </div>
        """

    html = f"""<!DOCTYPE html>
                <html lang="en">
                <head>
                <meta charset="UTF-8">
                <title>Travel Dashboard</title>
                <style>
                body {{
                    font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif;
                    background: #f7f7f5;
                    margin: 0;
                    padding: 2rem;
                    color: #222;
                }}
                .wrap {{ max-width: 1100px; margin: 0 auto; }}
                h1 {{ font-size: 24px; font-weight: 500; margin: 0; }}
                h3 {{ font-size: 16px; font-weight: 500; margin: 0; }}
                .stat-card {{ background: #fff; border-radius: 8px; padding: 1rem; }}
                .panel {{ background: #fff; border: 1px solid #eee; border-radius: 12px; padding: 1rem 1.25rem; }}
                </style>
                </head>
                <body>
                <div class="wrap">

                <div style="display:flex; justify-content:space-between; align-items:baseline; margin-bottom:1.5rem;">
                    <h1>Travel Dashboard</h1>
                    <span style="font-size:13px; color:#888;">Last updated {pd.Timestamp.now():%Y-%m-%d %H:%M}</span>
                </div>

                <div style="display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:12px; margin-bottom:1.5rem;">
                    <div class="stat-card">
                    <p style="font-size:13px; color:#888; margin:0 0 4px;">Countries & Territories Visited</p>
                    <p style="font-size:24px; font-weight:500; margin:0;">{stats['countries']:.0f}</p>
                    </div>
                    <div class="stat-card">
                    <p style="font-size:13px; color:#888; margin:0 0 4px;">Cities Visited</p>
                    <p style="font-size:24px; font-weight:500; margin:0;">{stats['cities']:.0f}</p>
                    </div>
                    <div class="stat-card">
                    <p style="font-size:13px; color:#888; margin:0 0 4px;">Airports Visited</p>
                    <p style="font-size:24px; font-weight:500; margin:0;">{stats['airports']:.0f}</p>
                    </div>
                    <div class="stat-card">
                    <p style="font-size:13px; color:#888; margin:0 0 4px;">Total Flights</p>
                    <p style="font-size:24px; font-weight:500; margin:0;">{stats['flights']:.0f}</p>
                    </div>
                    <div class="stat-card">
                    <p style="font-size:13px; color:#888; margin:0 0 4px;">Distinct Flight Routes</p>
                    <p style="font-size:24px; font-weight:500; margin:0;">{stats['routes']:.0f}</p>
                    </div>
                    {countries_this_year_card}
                </div>

                <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-bottom:1.5rem;">
                    <div class="panel">
                        <h3 style="margin-bottom:12px;">Top Countries</h3>
                        {countries_rows}
                    </div>
                    <div class="panel">
                        <h3 style="margin-bottom:12px;">Top Cities</h3>
                        {cities_rows}
                    </div>
                </div>

                <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-bottom:1.5rem;">
                    <div class="panel">
                        <h3 style="margin-bottom:12px;">Top Flight Routes</h3>
                        {routes_rows}
                    </div>
                    {years_panel}
                </div>

                {flag_strip}

                </div>
                </body>
                </html>
            """
    return html

def _render_years_html(year_df, country_df):
    name_to_iso2 = dict(zip(country_df["Common Name"], country_df["ISO2"]))

    def _split_names(s):
        if not isinstance(s, str) or not s.strip():
            return []
        return [n.strip() for n in s.split(";") if n.strip()]

    all_names = set()
    for col in ["Countries", "New Countries"]:
        for s in year_df[col]:
            all_names.update(_split_names(s))
    all_iso2 = {name_to_iso2[n] for n in all_names if n in name_to_iso2}
    svg_lookup = _build_country_svg_lookup(all_iso2, False)

    def _chip_list(names_str):
        names = _split_names(names_str)
        if not names:
            return '<span style="color:#aaa; font-size:12px;">None</span>'
        chips = []
        for name in names:
            iso2 = name_to_iso2.get(name)
            flag_html = f'<img class="chip-flag" src="{_get_flag_url(iso2)}" width="30" title="{name}" style="border-radius:2px; vertical-align:middle;">' if iso2 else ""
            outline_html = f'<span class="chip-outline" title="{name}" style="width:60px; height:42px; vertical-align:middle;">{svg_lookup.get(iso2, "")}</span>' if iso2 else ""
            chips.append(f'''
            <span class="country-chip" style="display:inline-flex; align-items:center; gap:5px; background:#f1f1ef; border-radius:6px; padding:3px 8px; margin:2px; font-size:12px; color:#444;">
                {flag_html}{outline_html}<span class="chip-name">{name}</span>
            </span>
            ''')
        return "".join(chips)

    year_cards = []
    for _, row in year_df.iterrows():
        year_cards.append(f"""
        <div class="panel year-card">
            <div style="display:flex; justify-content:space-between; align-items:baseline; margin-bottom:10px;">
                <h2 style="margin:0; font-size:20px; font-weight:500;">{row['Year']:.0f}</h2>
                <div style="display:flex; gap:20px; font-size:13px; color:#888;">
                    <span><strong style="color:#222;">{row['# Countries']:.0f}</strong> revisits</span>
                    <span><strong style="color:#222;">{row['# New Countries']:.0f}</strong> first-time visits</span>
                </div>
            </div>
            <div style="margin-bottom:8px;">
                <p style="font-size:12px; color:#888; margin:0 0 4px;">countries & territories visited</p>
                {_chip_list(row['Countries'])}
            </div>
            <div>
                <p style="font-size:12px; color:#888; margin:0 0 4px;">new countries & territories visited</p>
                {_chip_list(row['New Countries'])}
            </div>
        </div>
        """)

    html = f"""<!DOCTYPE html>
                <html lang="en">
                <head>
                <meta charset="UTF-8">
                <title>Visits By Year</title>
                <style>
                body {{
                    font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif;
                    background: #f7f7f5;
                    margin: 0;
                    padding: 2rem;
                    color: #222;
                }}
                .wrap {{ max-width: 1000px; margin: 0 auto; }}
                h1 {{ font-size: 24px; font-weight: 500; margin: 0; }}
                .panel {{ background: #fff; border: 1px solid #eee; border-radius: 12px; padding: 1rem 1.25rem; }}

                body.mode-flags .chip-outline, body.mode-flags .chip-name {{ display: none; }}
                body.mode-outlines .chip-flag, body.mode-outlines .chip-name {{ display: none; }}
                body.mode-names .chip-flag, body.mode-names .chip-outline {{ display: none; }}
                .chip-outline svg {{ width: 100%; height: 100%; display: block; }}
                </style>
                </head>
                <body class="mode-flags">
                
                <div class="wrap">
                    <div style="display:flex; justify-content:space-between; align-items:baseline; margin-bottom:1.5rem;">
                        <h1>Visits By Year</h1>
                        <span style="font-size:13px; color:#888;">Last updated {pd.Timestamp.now():%Y-%m-%d %H:%M}</span>
                    </div>
                    <div style="display:flex; justify-content:flex-end; align-items:center; gap:14px; margin-top:12px; margin-bottom:1.5rem;">
                        <button id="sort-toggle" type="button" title="Sort ascending" aria-label="Sort ascending"
                            style="border:1px solid #ddd; background:#f7f7f7; border-radius:6px;
                                padding:0 10px; height:28px; cursor:pointer; font-size:13px; line-height:1;
                                display:flex; align-items:center; gap:6px; white-space:nowrap;">
                            Sort chronologically   ▼
                        </button>
                        <div style="display:flex; gap:14px; font-size:12px; color:#555;">
                            <label style="display:flex; align-items:center; gap:4px; cursor:pointer;">
                                <input type="radio" name="chip-mode" value="flags" checked> Flags
                            </label>
                            <label style="display:flex; align-items:center; gap:4px; cursor:pointer;">
                                <input type="radio" name="chip-mode" value="outlines"> Outlines
                            </label>
                            <label style="display:flex; align-items:center; gap:4px; cursor:pointer;">
                                <input type="radio" name="chip-mode" value="names"> Names
                            </label>
                        </div>
                    
                </div>
                
                <div id="year-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:16px;">
                    {"".join(year_cards)}
                </div>
                </div>

                </div>
                <script>
                    document.querySelectorAll('input[name="chip-mode"]').forEach(function(radio) {{
                        radio.addEventListener('change', function() {{
                            document.body.className = 'mode-' + this.value;
                        }});
                    }});
                </script>
                <script>
                    let sortDescending = true; // matches the server-rendered default: newest year first
                    const sortBtn = document.getElementById('sort-toggle');
                    const yearGrid = document.getElementById('year-grid');
        
                    sortBtn.addEventListener('click', function() {{
                        const cards = Array.from(yearGrid.children);
                        cards.reverse().forEach(function(card) {{
                            yearGrid.appendChild(card); // re-appending an existing node moves it, doesn't duplicate it
                        }});
                        sortDescending = !sortDescending;
                        sortBtn.textContent = sortDescending ? 'Sort chronologically   ▼' : 'Sort chronologically   ▲';
                        sortBtn.title = sortDescending ? 'Sort ascending' : 'Sort descending';
                        sortBtn.setAttribute('aria-label', sortBtn.title);
                    }});
                </script>
                </body>
                </html>
            """
    return html

def build_cities_countries_map(open_browser=False, lang_code="en"):
    wb = xw.Book.caller()
    _set_paths(wb)

    city_df = _read_city_table(wb)
    city_df = city_df.dropna(subset=["Lat", "Lon"])

    country_df = _read_country_table(wb)
    country_df["Visits"] = pd.to_numeric(country_df["Visits"], errors="coerce").fillna(0)

    geojson_data = _get_geojson_data()

    fmap = _create_base_map(city_df, lang_code)
    _add_country_choropleth(fmap, country_df, geojson_data, _get_exclusion_countries())
    _add_country_tooltips(fmap, country_df, geojson_data, _get_exclusion_countries())
    _add_city_markers(fmap, city_df)

    _save_map_to_html(fmap, CITY_COUNTRY_HTML)

def build_aviation_map(open_browser=False, lang_code="en"):
    wb = xw.Book.caller()
    _set_paths(wb)

    with open(os.path.join(_base_path, COUNTRIES_GEOJSON), encoding="utf-8") as f:
        geojson_data = json.load(f)

    country_df = _read_country_table(wb)
    filtered_country_df = country_df.dropna(subset=["Visits"])
    airport_df = _read_airport_table(wb)
    filtered_airport_df = airport_df.dropna(subset=["Lat", "Lon"])
    flights_df = _read_flight_table(wb)
    airlines_df = _read_airline_table(wb)

    fmap = _create_base_map(None, lang_code)
    _add_country_choropleth(fmap, filtered_country_df, geojson_data,  _get_exclusion_countries())
    _add_airport_markers(fmap, filtered_airport_df)
    _add_flight_routes(fmap, flights_df, filtered_airport_df, airlines_df)

    _save_map_to_html(fmap, AVIATION_HTML)

def build_dashboard(open_browser=True):
    wb = xw.Book.caller()
    _set_paths(wb)

    country_df = _read_country_table(wb)
    city_df = _read_city_table(wb)

    airport_df = _read_airport_table(wb)
    airport_city_by_iata = dict(zip(airport_df["IATA Code"], airport_df["City"]))
    routes_df = _aggregate_flight_count_directional(_read_flight_table(wb))
    routes_df["Origin City"] = routes_df["Origin IATA"].map(airport_city_by_iata)
    routes_df["Destination City"] = routes_df["Destination IATA"].map(airport_city_by_iata)

    year_df = _read_year_table(wb)

    stats = {
        "countries": len(country_df[country_df["Visits"] > 0]),
        "cities": len(city_df[city_df["Visits"] > 0]),
        "flights": routes_df["Flights"].sum(),
        "airports": len(airport_df[(airport_df["Departures"] > 0) | (airport_df["Arrivals"] > 0)]),
        "routes": len(routes_df),
        "years": len(year_df),
    }

    all_countries_df=country_df[country_df["Visits"] > 0]
    top_countries = country_df.sort_values("Visits", ascending=False).head(DASHBOARD_TOP_SLICE)
    top_cities = city_df.sort_values("Visits", ascending=False).head(DASHBOARD_TOP_SLICE)
    top_routes = routes_df.sort_values("Flights", ascending=False).head(DASHBOARD_TOP_SLICE)

    top_years_by_countries = year_df.sort_values("# Countries", ascending=False).head(DASHBOARD_TOP_SLICE)
    top_years_by_new = year_df.sort_values("# New Countries", ascending=False).head(DASHBOARD_TOP_SLICE)

    dashboard_html = _render_dashboard_html(stats, top_countries, top_cities, top_routes, top_years_by_countries, top_years_by_new, all_countries_df)
    _save_dashboard_to_html(dashboard_html, DASHBOARD_HTML)

def build_years_page(open_browser=True):
    wb = xw.Book.caller()
    _set_paths(wb)

    year_df = _read_year_table(wb).sort_values("Year", ascending=False).reset_index(drop=True)
    country_df = _read_country_table(wb)

    years_html = _render_years_html(year_df, country_df)
    _save_dashboard_to_html(years_html, BY_YEAR_HTML)

#endregion

if __name__ == "__main__":
    xw.Book.caller = lambda: xw.books.active
    
    build_cities_countries_map()
    print(f"Cities and countries map written to {_base_path}\\{CITY_COUNTRY_HTML}")

    build_aviation_map()
    print(f"Aviation map written to {_base_path}\\{AVIATION_HTML}")

    build_dashboard()
    print(f"Dashboard written to {_base_path}\\{DASHBOARD_HTML}")

    build_years_page()
    print(f"Yearly overview written to {_base_path}\\{BY_YEAR_HTML}")

    input("Press any key to exit ...")
