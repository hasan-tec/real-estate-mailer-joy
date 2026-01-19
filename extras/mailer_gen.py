import pandas as pd
import os
import sys
import json
from geopy.distance import geodesic
from jinja2 import Template
from pypdf import PdfWriter

# --- WINDOWS GTK PATCH ---
# Updated to your specific D: drive path and UCRT64 folder
MSYS2_BIN_PATH = r'D:\msys2\ucrt64\bin' 

if sys.platform == 'win32' and os.path.exists(MSYS2_BIN_PATH):
    # This tells Python where to find the DLLs you installed via MSYS2
    os.add_dll_directory(MSYS2_BIN_PATH)

# --- SYSTEM CHECK: WEASYPRINT ---
try:
    from weasyprint import HTML
except (OSError, ImportError) as e:
    print("\n" + "="*60)
    print("ERROR: WEASYPRINT / GTK DEPENDENCIES MISSING")
    print("="*60)
    print(f"Details: {e}")
    print("\nPRO TIP FROM R&D:")
    print(f"Go check if this folder exists and has files: {MSYS2_BIN_PATH}")
    print("If it's empty, you might have installed GTK in a different MSYS2 environment.")
    print("Check 'D:\\msys2\\mingw64\\bin' or 'C:\\msys64\\ucrt64\\bin' as well.")
    print("="*60 + "\n")
    sys.exit(1)

# --- CONFIGURATION ---
CLIENT_CSV = 'Clientlist1-25.csv'
SOLD_CSV = 'Justsoldtest2-5.csv'
OUTPUT_DIR = 'output'
FINAL_PDF = 'final_mailers_batch.pdf'

# Create output directories
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
if not os.path.exists(f"{OUTPUT_DIR}/individual"):
    os.makedirs(f"{OUTPUT_DIR}/individual")

# --- STEP 1: LOAD & CLEAN DATA ---
print("--- Loading CSV data...")
try:
    df_clients = pd.read_csv(CLIENT_CSV)
    df_sold = pd.read_csv(SOLD_CSV)
except FileNotFoundError as e:
    print(f"Error: Could not find one of the CSV files. {e}")
    sys.exit(1)

# Basic cleaning
df_clients = df_clients.dropna(subset=['Address'])
df_sold = df_sold.dropna(subset=['Address'])

# --- STEP 2: MOCK GEOCODING ---
def get_mock_coords(row):
    import random
    return (35.3733 + random.uniform(-0.05, 0.05), -119.0187 + random.uniform(-0.05, 0.05))

print("--- Geocoding properties (Simulated)...")
df_clients['coords'] = df_clients.apply(get_mock_coords, axis=1)
df_sold['coords'] = df_sold.apply(get_mock_coords, axis=1)

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
        .postcard { width: 100%; height: 100%; border: 1px solid #eee; position: relative; box-sizing: border-box; }
        .header { color: #2c3e50; border-bottom: 2px solid #2c3e50; margin-bottom: 10px; }
        .map-box { width: 200px; height: 150px; background: #ddd; float: right; margin-left: 15px; text-align: center; line-height: 150px; font-size: 12px; }
        .sold-list { font-size: 11px; margin-top: 10px; }
        .sold-item { margin-bottom: 5px; padding: 3px; background: #f9f9f9; border-left: 3px solid #2c3e50; }
        .footer { position: absolute; bottom: 10px; font-size: 10px; color: #7f8c8d; }
    </style>
</head>
<body>
    <div class="postcard">
        <div class="header">
            <h2 style="margin:0;">Recent Sales Near You!</h2>
        </div>
        <div class="map-box">
            <img src="https://via.placeholder.com/200x150.png?text=Map+of+{{ client_zip }}" style="width:100%; height:100%;">
        </div>
        <div class="content">
            <p>Hi <strong>{{ first_name }}</strong>,</p>
            <p>Homes near <strong>{{ address }}</strong> are selling fast. Check out these recent neighbors:</p>
            <div class="sold-list">
                {% for property in nearby %}
                <div class="sold-item">
                    <strong>{{ property.Address }}</strong> - Sold for ${{ "{:,.0f}".format(property['Purchase Amt']) }} ({{ "{:.2f}".format(property.distance) }} miles away)
                </div>
                {% endfor %}
            </div>
        </div>
        <div class="footer">
            Generated for {{ first_name }} {{ last_name }} | {{ client_city }}, {{ client_zip }}
        </div>
    </div>
</body>
</html>
"""
template = Template(html_template_str)

# --- STEP 5: BATCH PROCESSING ---
print(f"--- Generating {len(df_clients)} mailers...")
pdf_files = []

# Testing first 10
for index, client in df_clients.head(10).iterrows():
    nearby_homes = find_nearest_sold(client['coords'], df_sold, n=3)
    
    html_out = template.render(
        first_name=client['Primary First'],
        last_name=client['Primary Last'],
        address=client['Address'],
        client_city=client['City'],
        client_zip=int(client['ZIP']),
        nearby=nearby_homes
    )
    
    file_path = f"{OUTPUT_DIR}/individual/mailer_{index}.pdf"
    try:
        HTML(string=html_out).write_pdf(file_path)
        pdf_files.append(file_path)
    except Exception as e:
        print(f"      Error generating PDF for {client['Address']}: {e}")
    
    if (index + 1) % 5 == 0:
        print(f"    Processed {index + 1} mailers...")

# --- STEP 6: MERGE ALL PDFS ---
if pdf_files:
    print("--- Merging batch into final file...")
    merger = PdfWriter()
    for pdf in pdf_files:
        merger.append(pdf)

    with open(f"{OUTPUT_DIR}/{FINAL_PDF}", "wb") as f:
        merger.write(f)
    print(f"--- SUCCESS! Final batch saved to {OUTPUT_DIR}/{FINAL_PDF}")
else:
    print("--- FAILED: No PDFs were generated.")