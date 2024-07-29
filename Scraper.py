import googlemaps
import time
from openpyxl import Workbook
from rich.console import Console
from datetime import datetime, timedelta
from shapely.geometry import shape, Point
import json
import os

# Initialize rich console for better terminal output
console = Console()

# Set your Google Maps API key here
API_KEY = 'YOUR_Ow'  # Replace with your actual API key

# Load Spain's borders from a GeoJSON file
GEOJSON_FILE_PATH = 'C:\\Users\\vps\\Downloads\\esp_adm0.geojson'

console.log("[bold blue]Starting the process...[/bold blue]")

if not os.path.exists(GEOJSON_FILE_PATH):
    console.log(f"[bold red]GeoJSON file not found at path: {GEOJSON_FILE_PATH}[/bold red]")
else:
    try:
        console.log(f"[bold blue]Loading GeoJSON file from: {GEOJSON_FILE_PATH}[/bold blue]")
        with open(GEOJSON_FILE_PATH, 'r') as f:
            geojson_data = json.load(f)
        SPAIN_BORDERS = shape(geojson_data['features'][0]['geometry'])
        console.log("[bold green]GeoJSON file loaded successfully.[/bold green]")
    except Exception as e:
        console.log(f"[bold red]Error loading GeoJSON file: {e}[/bold red]")

# Number of grids to split the bounding box into
NUM_GRIDS = 10000  # Increase to ensure detailed coverage
console.log(f"[bold blue]Number of grids set to: {NUM_GRIDS}[/bold blue]")

# Define maximum number of API requests per day
MAX_REQUESTS_PER_DAY = 500000
MAX_REQUESTS_PER_MINUTE = 2000
console.log(f"[bold blue]Max API requests per day: {MAX_REQUESTS_PER_DAY}[/bold blue]")
console.log(f"[bold blue]Max API requests per minute: {MAX_REQUESTS_PER_MINUTE}[/bold blue]")

def create_excel_file(data, filename):
    try:
        console.log(f"[bold blue]Creating Excel file: {filename}[/bold blue]")
        workbook = Workbook()
        sheet = workbook.active
        headers = ["Name", "Address", "Phone Number", "Website"]
        sheet.append(headers)
        for item in data:
            sheet.append(item)
        workbook.save(filename)
        console.log(f"[bold green]Data has been written to {filename}[/bold green]")
    except Exception as e:
        console.log(f"[bold red]Error writing to Excel file: {e}[/bold red]")

def get_place_details(gmaps, place_id):
    try:
        console.log(f"[bold blue]Fetching details for place ID: {place_id}[/bold blue]")
        place_details = gmaps.place(place_id=place_id)
        result = place_details.get('result', {})
        phone_number = result.get('formatted_phone_number', "")
        website = result.get('website', "")
        console.log(f"[bold green]Retrieved details: Phone - {phone_number}, Website - {website}[/bold green]")
        return phone_number, website
    except Exception as e:
        console.log(f"[bold red]Error fetching place details: {e}[/bold red]")
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
            console.log(f"[bold blue]Requesting places for location: {location} with radius {radius}[/bold blue]")
            places_result = gmaps.places_nearby(location=location, radius=radius, type='lodging', page_token=next_page_token)
            total_requests += 1
            console.log(f"[bold blue]Total requests so far: {total_requests}[/bold blue]")

            for place in places_result.get('results', []):
                name = place.get('name')
                address = place.get('vicinity')
                place_id = place.get('place_id')
                phone_number, website = get_place_details(gmaps, place_id)
                
                # Unique identifier for the place
                place_identifier = (name, address)
                
                if place_identifier not in unique_places:
                    unique_places.add(place_identifier)
                    places.append([name, address, phone_number, website])
                    console.log(f"[bold green]Found place: {name}, {address}[/bold green]")
                else:
                    identical_count += 1
                    console.log(f"[bold yellow]Identical place skipped: {name}, {address}[/bold yellow]")
            
            next_page_token = places_result.get('next_page_token')
            if not next_page_token:
                console.log("[bold green]No more pages to fetch for this location.[/bold green]")
                break
            time.sleep(2)  # Avoid hitting the rate limit

            # Check if the request limit per minute is reached
            if total_requests % MAX_REQUESTS_PER_MINUTE == 0:
                console.log(f"[bold red]Rate limit reached. Pausing for a minute...[/bold red]")
                time.sleep(60)  # Pause to respect the rate limit

        except googlemaps.exceptions.ApiError as e:
            if 'OVER_QUERY_LIMIT' in str(e):
                console.log("[bold red]Reached API rate limit. Waiting before retrying...[/bold red]")
                time.sleep(10)  # Wait before retrying
            else:
                console.log(f"[bold red]API Error fetching places: {e}[/bold red]")
                break

        except Exception as e:
            console.log(f"[bold red]Error fetching places: {e}[/bold red]")
            break

    console.log(f"[bold blue]Total identical places skipped: {identical_count}[/bold blue]")
    return places

def generate_grid_coordinates(borders, num_grids):
    try:
        console.log(f"[bold blue]Generating grid coordinates with {num_grids} grids...[/bold blue]")
        min_lng, min_lat, max_lng, max_lat = borders.bounds
        grid_size = int(num_grids ** 0.5) + 1  # Ensure at least num_grids points
        lat_step = (max_lat - min_lat) / grid_size
        lng_step = (max_lng - min_lng) / grid_size
        coordinates = []
        lat = min_lat
        while lat <= max_lat:
            lng = min_lng
            while lng <= max_lng:
                point = Point(lng + lng_step / 2, lat + lat_step / 2)
                if borders.contains(point):
                    coordinates.append((point.y, point.x))
                    console.log(f"[bold green]Point added: {point.y}, {point.x}[/bold green]")
                lng += lng_step
            lat += lat_step
        console.log(f"[bold green]Generated {len(coordinates)} grid coordinates.[/bold green]")
        return coordinates
    except Exception as e:
        console.log(f"[bold red]Error generating grid coordinates: {e}[/bold red]")
        return []

def main():
    console.log("[bold blue]Starting main process...[/bold blue]")
    api_key = API_KEY
    if not api_key:
        console.log("[bold red]API key not found. Please set the API_KEY variable.[/bold red]")
        return

    query = "hotel"
    radius = 5000  # Set a radius that balances between coverage and API limits
    console.log(f"[bold blue]Query: {query}, Radius: {radius}[/bold blue]")

    # Generate grid coordinates
    coordinates = generate_grid_coordinates(SPAIN_BORDERS, NUM_GRIDS)
    if not coordinates:
        console.log("[bold yellow]No valid grid coordinates generated. Check the GeoJSON data and boundaries.[/bold yellow]")
        return

    all_data = []
    total_coordinates = len(coordinates)
    console.log(f"[bold blue]Total coordinates to process: {total_coordinates}[/bold blue]")

    for i, coord in enumerate(coordinates):
        console.log(f"[bold blue]Fetching data for grid {i + 1}/{total_coordinates}: Center at {coord}[/bold blue]")
        data = get_places(api_key, query, coord, radius)
        all_data.extend(data)
        console.log(f"[bold blue]Completed fetching data for grid {i + 1}/{total_coordinates}[/bold blue]")
        # Optionally, add a sleep to respect the API's rate limit
        time.sleep(1)

    if all_data:
        filename = f"{query}_in_Spain.xlsx"
        create_excel_file(all_data, filename)
    else:
        console.log("[bold yellow]No data found.[/bold yellow]")

    console.log("[bold blue]Process completed.[/bold blue]")

if __name__ == "__main__":
    main()
