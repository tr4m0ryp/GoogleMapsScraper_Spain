import googlemaps
import time
from openpyxl import Workbook
from rich.console import Console
from rich.table import Table
from shapely.geometry import shape, Point
import json
from datetime import datetime, timedelta

# Initialize rich console for better terminal output
console = Console()

# Set your Google Maps API key here
API_KEY = 'API KEY'  # Replace with your actual API key

# Load Spain's borders from a GeoJSON file (assumed to be in the same directory)
with open('spain_borders.geojson', 'r') as f:
    geojson_data = json.load(f)
SPAIN_BORDERS = shape(geojson_data['features'][0]['geometry'])

# Number of grids to split the bounding box into
NUM_GRIDS = 1000

# Define maximum number of API requests per day
MAX_REQUESTS_PER_DAY = 1000

def create_excel_file(data, filename):
    try:
        workbook = Workbook()
        sheet = workbook.active
        headers = ["Name", "Address", "Phone Number", "Website"]
        sheet.append(headers)
        for item in data:
            sheet.append(item)
        workbook.save(filename)
        console.log(f"Data has been written to [bold green]{filename}[/bold green]")
    except Exception as e:
        console.log(f"[bold red]Error writing to Excel file:[/bold red] {e}")

def get_place_details(gmaps, place_id):
    try:
        place_details = gmaps.place(place_id=place_id)
        result = place_details.get('result', {})
        phone_number = result.get('formatted_phone_number', "")
        website = result.get('website', "")
        return phone_number, website
    except Exception as e:
        console.log(f"[bold red]Error fetching place details:[/bold red] {e}")
        return "", ""

def get_places(api_key, query, location, radius):
    gmaps = googlemaps.Client(key=api_key)
    places = []
    unique_places = set()
    identical_count = 0
    next_page_token = None
    total_requests = 0

    while True:
        try:
            places_result = gmaps.places_nearby(location=location, radius=radius, type='lodging', page_token=next_page_token)
            total_requests += 1

            for place in places_result.get('results', []):
                name = place.get('name')
                address = place.get('vicinity')
                place_id = place.get('place_id')
                phone_number, website = get_place_details(gmaps, place_id)
                
                place_identifier = (name, address)
                
                if place_identifier not in unique_places:
                    unique_places.add(place_identifier)
                    places.append([name, address, phone_number, website])
            
            next_page_token = places_result.get('next_page_token')
            if not next_page_token:
                break
            time.sleep(2)  # Avoid hitting the rate limit

            if total_requests >= MAX_REQUESTS_PER_DAY:
                now = datetime.now()
                midnight = datetime.combine(now.date() + timedelta(days=1), datetime.min.time())
                sleep_seconds = (midnight - now).total_seconds()
                console.log(f"Sleeping for {sleep_seconds / 3600:.2f} hours.")
                time.sleep(sleep_seconds)
                total_requests = 0

        except googlemaps.exceptions.ApiError as e:
            if 'OVER_QUERY_LIMIT' in str(e):
                console.log("[bold red]Reached API rate limit. Waiting before retrying...[/bold red]")
                time.sleep(10)
            else:
                console.log(f"[bold red]API Error fetching places:[/bold red] {e}")
                break

        except Exception as e:
            console.log(f"[bold red]Error fetching places:[/bold red] {e}")
            break

    return places

def generate_grid_coordinates(borders, num_pieces):
    min_lng, min_lat, max_lng, max_lat = borders.bounds
    lat_step = (max_lat - min_lat) / (num_pieces ** 0.5)
    lng_step = (max_lng - min_lng) / (num_pieces ** 0.5)
    coordinates = []
    lat = min_lat
    while lat < max_lat:
        lng = min_lng
        while lng < max_lng:
            point = Point(lng + lng_step / 2, lat + lat_step / 2)
            if borders.contains(point):
                coordinates.append((point.y, point.x))
            lng += lng_step
        lat += lat_step
    return coordinates

def main():
    api_key = API_KEY
    if not api_key:
        console.log("[bold red]API key not found. Please set the API_KEY variable.[/bold red]")
        return

    query = "hotel"
    radius = 5000  # Set a radius that balances between coverage and API limits

    # Generate grid coordinates
    coordinates = generate_grid_coordinates(SPAIN_BORDERS, NUM_GRIDS)
    console.log(f"Generated {len(coordinates)} grid coordinates within Spain's borders.")

    all_data = []

    for i, coord in enumerate(coordinates):
        console.log(f"Fetching data for grid {i + 1}/{len(coordinates)}: Center at {coord}")
        data = get_places(api_key, query, coord, radius)
        all_data.extend(data)
        time.sleep(1)

    if all_data:
        filename = f"{query}_in_Spain.xlsx"
        create_excel_file(all_data, filename)
    else:
        console.log("[bold yellow]No data found.[/bold yellow]")

if __name__ == "__main__":
    main()
