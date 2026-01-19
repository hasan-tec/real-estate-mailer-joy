import pandas as pd
import os
import sys
import json
import requests
import time
from geopy.distance import geodesic
from jinja2 import Template
from pypdf import PdfWriter

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
    print(f"Details: {e}")
    sys.exit(1)

# --- CONFIGURATION ---
TOMTOM_API_KEY = '2kruf0YV0Ixe85dxzYx4uR08XPbu7ywo'
CLIENT_CSV = 'Clientlist1-25.csv'
SOLD_CSV = 'Justsoldtest2-5.csv'
OUTPUT_DIR = 'output'
CACHE_FILE = 'geocoding_cache.json'
FINAL_PDF = 'final_mailers_batch.pdf'
SKIPPED_REPORT = 'skipped_addresses.csv'

# Tracking lists for logging
skipped_log = []

# Create output directories
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
if not os.path.exists(f"{OUTPUT_DIR}/individual"):
    os.makedirs(f"{OUTPUT_DIR}/individual")

# --- CACHE HELPERS ---
def load_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_cache(cache):
    with open(CACHE_FILE, 'w') as f:
        json.dump(cache, f)

# --- STEP 1: LOAD & CLEAN DATA ---
print("\n[1/6] Loading CSV data...")
try:
    df_clients = pd.read_csv(CLIENT_CSV)
    df_sold = pd.read_csv(SOLD_CSV)
    
    total_clients_in_file = len(df_clients)
    print(f"      File loaded. Total clients available: {total_clients_in_file}")
    
    user_input = input(f"      How many clients to process? (number or 'all'): ").strip().lower()
    
    if user_input == 'all' or user_input == '':
        pass
    else:
        try:
            limit = int(user_input)
            df_clients = df_clients.head(limit)
        except ValueError:
            df_clients = df_clients.head(10)

    print(f"      Success: Ready to process {len(df_clients)} clients.")
except Exception as e:
    print(f"      CRITICAL ERROR: {e}")
    sys.exit(1)

# --- STEP 2: TOMTOM GEOCODING ---
cache = load_cache()

def get_coords_tomtom(row, index, total, list_type="Client"):
    raw_address = str(row['Address']).strip()
    raw_city = str(row.get('City', 'Bakersfield')).strip()
    raw_zip = str(row.get('ZIP', '')).split('.')[0].strip()
    
    full_address = f"{raw_address}, {raw_city}, CA {raw_zip}"
    
    sys.stdout.write(f"\r      -> TomTom Geocoding {list_type} {index + 1}/{total}...")
    sys.stdout.flush()

    if full_address in cache and cache[full_address] is not None:
        return cache[full_address]
    
    try:
        # TomTom Geocoding API endpoint
        url = f"https://api.tomtom.com/search/2/geocode/{requests.utils.quote(full_address)}.json?key={TOMTOM_API_KEY}"
        response = requests.get(url)
        data = response.json()
        
        if data.get('results'):
            pos = data['results'][0]['position']
            coords = [pos['lat'], pos['lon']]
            cache[full_address] = coords
            # Save cache every 10 addresses
            if len(cache) % 10 == 0: save_cache(cache)
            return coords
        else:
            skipped_log.append({'Address': full_address, 'Reason': 'TomTom: No results', 'Type': list_type, 'Row': index + 2})
    except Exception as e:
        skipped_log.append({'Address': full_address, 'Reason': f'TomTom Error: {str(e)}', 'Type': list_type, 'Row': index + 2})
    
    return None

print(f"\n[2/6] Geocoding properties with TomTom...")
coords_clients = [get_coords_tomtom(row, i, len(df_clients), "Client") for i, row in df_clients.iterrows()]
df_clients['coords'] = coords_clients
save_cache(cache)

print("\n      Sold Properties:")
coords_sold = [get_coords_tomtom(row, i, len(df_sold), "Sold") for i, row in df_sold.iterrows()]
df_sold['coords'] = coords_sold
save_cache(cache)

valid_clients = df_clients.dropna(subset=['coords']).copy()
valid_sold = df_sold.dropna(subset=['coords']).copy()

print(f"\n      Found {len(valid_clients)} valid clients and {len(valid_sold)} valid sales.")

# --- STEP 3: DISTANCE LOGIC ---
def find_nearest_sold(client_coords, sold_df, n=3):
    sold_pool = sold_df.copy()
    sold_pool['distance'] = sold_pool['coords'].apply(lambda x: geodesic(client_coords, x).miles)
    return sold_pool.sort_values('distance').head(n).to_dict('records')

# --- STEP 4: HTML TEMPLATE ---
html_template_str = """
<!DOCTYPE html>
<html>
<head>
    <style>
        @page { size: 6in 4in; margin: 0; }
        body { font-family: 'Arial', sans-serif; margin: 0; padding: 20px; color: #333; }
        .postcard { width: 100%; height: 100%; position: relative; box-sizing: border-box; }
        .header { color: #cc0000; border-bottom: 2px solid #cc0000; margin-bottom: 10px; }
        .map-box { width: 220px; height: 180px; background: #eee; float: right; margin-left: 15px; border: 1px solid #ccc; border-radius: 4px; overflow: hidden; }
        .sold-list { font-size: 11px; margin-top: 10px; }
        .sold-item { margin-bottom: 5px; padding: 5px; background: #fdfdfd; border-left: 4px solid #cc0000; border-bottom: 1px solid #eee; }
        .footer { position: absolute; bottom: 0; font-size: 9px; color: #777; width: 100%; border-top: 1px solid #eee; padding-top: 5px; }
    </style>
</head>
<body>
    <div class="postcard">
        <div class="header">
            <h2 style="margin:0; font-size: 20px;">Neighborhood Sold Alert</h2>
        </div>
        <div class="map-box">
            <img src="{{ map_url }}" style="width:100%; height:100%; object-fit: cover;">
        </div>
        <div class="content" style="width: 280px;">
            <p style="margin-top:0;">Hi Neighbor,</p>
            <p style="font-size: 11px; line-height: 1.4;">Homes near <strong>{{ address }}</strong> are moving! Check out these 3 neighbors who recently sold:</p>
            <div class="sold-list">
                {% for property in nearby %}
                <div class="sold-item">
                    <strong>{{ property.Address }}</strong><br>
                    Sold: ${{ "{:,.0f}".format(property['Purchase Amt']) }} | {{ "{:.2f}".format(property.distance) }} miles away
                </div>
                {% endfor %}
            </div>
        </div>
        <div class="footer">
            Generated using TomTom Mapping Technology. Prepared for {{ address }}.
        </div>
    </div>
</body>
</html>
"""
template = Template(html_template_str)

# --- STEP 5: BATCH GENERATION ---
print(f"\n[3/6] Rendering {len(valid_clients)} Mailers with TomTom Maps...")
pdf_files = []

for index, client in valid_clients.iterrows():
    nearby_homes = find_nearest_sold(client['coords'], valid_sold, n=3)
    lat, lon = client['coords']
    
    # Generate TomTom Static Map URL
    # center is lon,lat for TomTom
    map_url = f"https://api.tomtom.com/map/1/staticimage?key={TOMTOM_API_KEY}&zoom=15&center={lon},{lat}&format=png&layer=basic&style=main&width=440&height=360"
    
    # Add Marker for client
    map_url += f"&pins=default|red|{lon},{lat}"
    
    html_out = template.render(
        address=client['Address'],
        nearby=nearby_homes,
        map_url=map_url
    )
    
    file_path = f"{OUTPUT_DIR}/individual/mailer_{index}.pdf"
    try:
        HTML(string=html_out).write_pdf(file_path)
        pdf_files.append(file_path)
        sys.stdout.write(f"\r      -> Created {len(pdf_files)}/{len(valid_clients)} PDFs...")
        sys.stdout.flush()
    except Exception as e:
        skipped_log.append({'Address': client['Address'], 'Reason': f'PDF Render Error: {str(e)}', 'Type': 'Client', 'Row': index + 2})

# --- STEP 6: MERGE & AUDIT ---
print("\n\n[4/6] Finalizing Output...")
if pdf_files:
    merger = PdfWriter()
    for pdf in pdf_files: merger.append(pdf)
    with open(f"{OUTPUT_DIR}/{FINAL_PDF}", "wb") as f: merger.write(f)
    print(f"      Success: Generated {FINAL_PDF}")

if skipped_log:
    pd.DataFrame(skipped_log).to_csv(f"{OUTPUT_DIR}/{SKIPPED_REPORT}", index=False)
    print(f"      Audit: {len(skipped_log)} items skipped. Check {SKIPPED_REPORT}")

print(f"\n[6/6] Done! TomTom batch processing complete.\n")