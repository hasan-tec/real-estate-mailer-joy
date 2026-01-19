import pandas as pd
import os
import sys
import json
import requests
import time
from geopy.distance import geodesic
from jinja2 import Template
from pypdf import PdfWriter
from dotenv import load_dotenv

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Load environment variables
load_dotenv(os.path.join(SCRIPT_DIR, '.env'))

# --- WINDOWS GTK PATCH ---
MSYS2_BIN_PATH = r'D:\msys2\ucrt64\bin' 

if sys.platform == 'win32' and os.path.exists(MSYS2_BIN_PATH):
    os.add_dll_directory(MSYS2_BIN_PATH)

# --- SYSTEM CHECK: WEASYPRINT ---
try:
    from weasyprint import HTML
except (OSError, ImportError) as e:
    print("\n" + "="*60)
    print("ERROR: WEASYPRINT / GTK DEPENDENCIES MISSING")
    print("="*60)
    sys.exit(1)

# --- CONFIGURATION ---
TOMTOM_API_KEY = os.getenv('TOMTOM_API_KEY', '')
MAPBOX_TOKEN = os.getenv('MAPBOX_TOKEN', '')
CLIENT_CSV = 'Clientlist1-25.csv'
SOLD_CSV = 'Justsoldtest2-5.csv'
OUTPUT_DIR = 'output'
MAP_DEBUG_DIR = os.path.join(OUTPUT_DIR, 'debug_maps')
CACHE_FILE = 'geocoding_cache.json'
FINAL_PDF = 'final_mailers_batch.pdf'
SKIPPED_REPORT = 'skipped_addresses.csv'

# Initialize global log
skipped_log = []

# Create output directories
for folder in [OUTPUT_DIR, os.path.join(OUTPUT_DIR, 'individual'), MAP_DEBUG_DIR]:
    if not os.path.exists(folder):
        os.makedirs(folder)

def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r') as f:
                return json.load(f)
        except: return {}
    return {}

def save_cache(cache):
    with open(CACHE_FILE, 'w') as f:
        json.dump(cache, f)

# --- STEP 1: LOAD & CLEAN DATA ---
print("\n[1/6] Loading CSV data...")
try:
    df_clients = pd.read_csv(CLIENT_CSV)
    df_sold = pd.read_csv(SOLD_CSV)
    
    # Clean garbage rows
    df_clients = df_clients[df_clients['Address'].str.contains("The information", na=False) == False]
    df_sold = df_sold[df_sold['Address'].str.contains("The information", na=False) == False]

    user_input = input(f"      How many clients to process? (number or 'all'): ").strip().lower()
    if user_input != 'all' and user_input != '':
        df_clients = df_clients.head(int(user_input))

    print(f"      Ready to process {len(df_clients)} clients.")
except Exception as e:
    print(f"      CRITICAL ERROR: {e}")
    sys.exit(1)

# --- STEP 2: TOMTOM GEOCODING ---
cache = load_cache()

def get_coords_tomtom(row, index, total, list_type="Client"):
    global skipped_log
    address = str(row['Address']).strip()
    city = str(row.get('City', 'Bakersfield')).strip()
    zip_code = str(row.get('ZIP', '')).split('.')[0].strip()
    full_address = f"{address}, {city}, CA {zip_code}"
    
    sys.stdout.write(f"\r      -> Geocoding {list_type} {index + 1}/{total}...")
    sys.stdout.flush()

    if full_address in cache: return cache[full_address]
    
    try:
        url = f"https://api.tomtom.com/search/2/geocode/{requests.utils.quote(full_address)}.json?key={TOMTOM_API_KEY}"
        data = requests.get(url).json()
        if data.get('results'):
            pos = data['results'][0]['position']
            coords = [pos['lat'], pos['lon']]
            cache[full_address] = coords
            save_cache(cache)
            return coords
    except: pass
    skipped_log.append({'Address': full_address, 'Type': list_type})
    return None

print(f"\n[2/6] Geocoding properties...")
df_clients['coords'] = [get_coords_tomtom(row, i, len(df_clients), "Client") for i, row in df_clients.iterrows()]
df_sold['coords'] = [get_coords_tomtom(row, i, len(df_sold), "Sold") for i, row in df_sold.iterrows()]

valid_clients = df_clients.dropna(subset=['coords']).copy()
valid_sold = df_sold.dropna(subset=['coords']).copy()

# --- STEP 3: DISTANCE LOGIC ---
def find_nearest_sold(client_coords, sold_df, n=3):
    sold_pool = sold_df.copy()
    sold_pool['distance'] = sold_pool['coords'].apply(lambda x: geodesic(client_coords, x).miles)
    # Filter out identical address
    return sold_pool[sold_pool['distance'] > 0.005].sort_values('distance').head(n).to_dict('records')

# --- STEP 4: HTML TEMPLATE ---
html_template_str = """
<!DOCTYPE html>
<html>
<head>
    <style>
        @page { size: 6in 4in; margin: 0; }
        body { font-family: 'Helvetica', sans-serif; padding: 25px; color: #333; line-height: 1.2; }
        .map-box { width: 250px; height: 200px; float: right; border: 2px solid #333; border-radius: 8px; overflow: hidden; }
        .header { color: #d32f2f; border-bottom: 3px solid #d32f2f; margin-bottom: 10px; }
        .sold-item { font-size: 11px; margin-bottom: 5px; padding: 5px; background: #f9f9f9; border-left: 4px solid #d32f2f; }
        .footer { position: absolute; bottom: 10px; font-size: 10px; color: #666; border-top: 1px solid #ddd; width: 100%; padding-top: 5px; }
    </style>
</head>
<body>
    <div class="header"><h2>Neighborhood Alert</h2></div>
    <div class="map-box"><img src="{{ map_url }}" style="width:100%; height:100%;"></div>
    <div class="content" style="width: 270px;">
        <p>Hi <b>{{ first_name }}</b>,</p>
        <p style="font-size: 12px;">Real estate is moving near <b>{{ address }}</b>! Recently sold neighbors:</p>
        {% for property in nearby %}
        <div class="sold-item"><b>{{ property.Address }}</b><br>Sold: ${{ "{:,.0f}".format(property['Purchase Amt']) }} ({{ "{:.2f}".format(property.distance) }} mi)</div>
        {% endfor %}
    </div>
    <div class="footer">Legend: Red Pin = Your Home | Green Pins = Neighbors</div>
</body>
</html>
"""
template = Template(html_template_str)

# --- STEP 5: RENDERING & IMAGE SAVE ---
print(f"\n[3/6] Generating Mailers & Exporting Maps...")
pdf_files = []

for index, client in valid_clients.iterrows():
    nearby = find_nearest_sold(client['coords'], valid_sold, n=3)
    lat, lon = client['coords']
    
    # --- MAPBOX STATIC IMAGE WITH MARKERS ---
    # Build marker string: pin-l+COLOR(lon,lat) for large, pin-s+COLOR for small
    # Red marker for client home
    markers = f"pin-l+e74c3c({lon},{lat})"
    
    # Green markers for nearby sold homes
    if nearby:
        for home in nearby:
            h_lat, h_lon = home['coords']
            markers += f",pin-s+27ae60({h_lon},{h_lat})"
    
    # Mapbox Static Images API URL
    map_url = f"https://api.mapbox.com/styles/v1/mapbox/streets-v12/static/{markers}/{lon},{lat},14/500x400?access_token={MAPBOX_TOKEN}"

    # Save map image for manual verification
    try:
        img_response = requests.get(map_url)
        if img_response.status_code == 200:
            safe_name = f"{client['Primary First']}_{client['Primary Last']}".replace(" ", "_")
            with open(os.path.join(MAP_DEBUG_DIR, f"map_{safe_name}.png"), "wb") as f:
                f.write(img_response.content)
    except: pass

    html_out = template.render(
        first_name=str(client['Primary First']).capitalize(),
        address=client['Address'],
        nearby=nearby,
        map_url=map_url
    )
    
    file_path = f"{OUTPUT_DIR}/individual/mailer_{index}.pdf"
    HTML(string=html_out).write_pdf(file_path)
    pdf_files.append(file_path)
    sys.stdout.write(f"\r      -> Created {len(pdf_files)}/{len(valid_clients)} PDFs...")
    sys.stdout.flush()

# --- STEP 6: FINAL MERGE ---
print("\n\n[4/6] Finalizing Output...")
if pdf_files:
    merger = PdfWriter()
    for pdf in pdf_files: merger.append(pdf)
    with open(os.path.join(OUTPUT_DIR, FINAL_PDF), "wb") as f: merger.write(f)
    print(f"      Success! Total pages: {len(pdf_files)}")

print(f"\n[6/6] Check {MAP_DEBUG_DIR} to verify the pins yourself!\n")