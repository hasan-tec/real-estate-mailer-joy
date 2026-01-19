"""
Real Estate Mailer Generator
A simple GUI app to generate real estate mailers from CSV data.
"""

import pandas as pd
import os
import sys
import json
import requests
import time
import threading
from datetime import datetime
from geopy.distance import geodesic
from jinja2 import Template
from pypdf import PdfWriter
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# --- GET SCRIPT DIRECTORY (for proper path handling) ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

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

# --- CONFIGURATION (using absolute paths) ---
DEFAULT_TOMTOM_API_KEY = '2kruf0YV0Ixe85dxzYx4uR08XPbu7ywo'
DEFAULT_MAPBOX_TOKEN = 'pk.eyJ1IjoiaGFzYW5hbmFzIiwiYSI6ImNta2wyajVkaTAwMmszZXIxaWczMmRsYjAifQ._-D6cOa84DMgXyU0a9JPaQ'
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'output')
INDIVIDUAL_DIR = os.path.join(OUTPUT_DIR, 'individual')
MAP_DEBUG_DIR = os.path.join(OUTPUT_DIR, 'debug_maps')
CACHE_FILE = os.path.join(SCRIPT_DIR, 'geocoding_cache.json')
FINAL_PDF = 'final_mailers_batch.pdf'
SKIPPED_REPORT = 'skipped_addresses.csv'


def ensure_directories():
    """Create output directories if they don't exist"""
    for folder in [OUTPUT_DIR, INDIVIDUAL_DIR, MAP_DEBUG_DIR]:
        if not os.path.exists(folder):
            os.makedirs(folder)


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


# --- IMPROVED HTML TEMPLATE ---
html_template_str = """
<!DOCTYPE html>
<html>
<head>
    <style>
        @page { 
            size: 6in 4in; 
            margin: 0; 
        }
        * {
            box-sizing: border-box;
        }
        body { 
            font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
            margin: 0;
            padding: 0;
            width: 6in;
            height: 4in;
            background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
            color: #2c3e50;
            position: relative;
        }
        .container {
            padding: 18px 22px;
            height: 100%;
            display: flex;
            flex-direction: column;
        }
        .header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 12px;
            padding-bottom: 10px;
            border-bottom: 3px solid #c0392b;
        }
        .header-left h1 {
            color: #c0392b;
            font-size: 22px;
            margin: 0 0 2px 0;
            font-weight: 700;
            letter-spacing: -0.5px;
        }
        .header-left .subtitle {
            font-size: 10px;
            color: #7f8c8d;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        .content-wrapper {
            display: flex;
            flex: 1;
            gap: 15px;
        }
        .left-content {
            flex: 1;
            display: flex;
            flex-direction: column;
        }
        .greeting {
            font-size: 13px;
            margin-bottom: 8px;
        }
        .greeting b {
            color: #c0392b;
        }
        .intro-text {
            font-size: 11px;
            color: #555;
            margin-bottom: 10px;
            line-height: 1.4;
        }
        .intro-text b {
            color: #2c3e50;
        }
        .sold-list {
            flex: 1;
        }
        .sold-item {
            background: #fff;
            border-left: 4px solid #27ae60;
            padding: 6px 10px;
            margin-bottom: 6px;
            border-radius: 0 6px 6px 0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        .sold-item .address {
            font-size: 11px;
            font-weight: 600;
            color: #2c3e50;
            margin-bottom: 2px;
        }
        .sold-item .details {
            font-size: 10px;
            color: #666;
        }
        .sold-item .price {
            color: #27ae60;
            font-weight: 600;
        }
        .sold-item .distance {
            color: #95a5a6;
        }
        .map-section {
            width: 220px;
            display: flex;
            flex-direction: column;
        }
        .map-box {
            width: 220px;
            height: 180px;
            border: 2px solid #bdc3c7;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        .map-box img {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }
        .map-legend {
            margin-top: 6px;
            font-size: 9px;
            color: #7f8c8d;
            text-align: center;
        }
        .map-legend .red { color: #c0392b; }
        .map-legend .green { color: #27ae60; }
        .footer {
            position: absolute;
            bottom: 0;
            left: 0;
            right: 0;
            background: #2c3e50;
            color: #fff;
            padding: 8px 22px;
            font-size: 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .footer-cta {
            font-weight: 600;
        }
        .footer-contact {
            color: #bdc3c7;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-left">
                <h1>🏠 Neighborhood Market Alert</h1>
                <div class="subtitle">Real Estate Activity Near You</div>
            </div>
        </div>
        
        <div class="content-wrapper">
            <div class="left-content">
                <div class="greeting">Hi <b>{{ first_name }}</b>,</div>
                <div class="intro-text">
                    Homes are selling near <b>{{ address }}</b>! Here are <b>{{ nearby|length }}</b> recent sales in your neighborhood:
                </div>
                
                <div class="sold-list">
                    {% for property in nearby %}
                    <div class="sold-item">
                        <div class="address">{{ property.Address }}</div>
                        <div class="details">
                            <span class="price">${{ "{:,.0f}".format(property['Purchase Amt']) }}</span>
                            <span class="distance">• {{ "{:.2f}".format(property.distance) }} miles away</span>
                            {% if property.get('Beds') and property.get('Baths') %}
                            <span>• {{ property.Beds }}bd/{{ property.Baths }}ba</span>
                            {% endif %}
                        </div>
                    </div>
                    {% endfor %}
                </div>
            </div>
            
            <div class="map-section">
                <div class="map-box">
                    <img src="{{ map_url }}" alt="Neighborhood Map">
                </div>
                <div class="map-legend">
                    <span class="red">●</span> Your Home &nbsp;|&nbsp; <span class="green">●</span> Recent Sales
                </div>
            </div>
        </div>
    </div>
    
    <div class="footer">
        <span class="footer-cta">Curious what your home is worth? Let's talk!</span>
        <span class="footer-contact">📞 Call for a free home valuation</span>
    </div>
</body>
</html>
"""


class MailerGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Real Estate Mailer Generator")
        self.root.geometry("750x750")
        self.root.resizable(True, True)
        
        # Variables
        self.client_csv_path = tk.StringVar()
        self.sold_csv_path = tk.StringVar()
        self.num_nearby = tk.IntVar(value=3)
        self.num_clients = tk.StringVar(value="all")
        self.tomtom_api_key = tk.StringVar(value=DEFAULT_TOMTOM_API_KEY)
        self.mapbox_token = tk.StringVar(value=DEFAULT_MAPBOX_TOKEN)
        self.cache = load_cache()
        self.skipped_log = []
        
        self.setup_ui()
        self.log("Application started")
        self.log(f"Working directory: {SCRIPT_DIR}")
        self.log(f"Output directory: {OUTPUT_DIR}")
    
    def setup_ui(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="🏠 Real Estate Mailer Generator", 
                                font=('Segoe UI', 16, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # --- File Selection Frame ---
        file_frame = ttk.LabelFrame(main_frame, text="📁 CSV Files", padding="10")
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Client CSV
        ttk.Label(file_frame, text="Client List (targets):").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.client_csv_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_client_csv).grid(row=0, column=2)
        
        # Sold CSV
        ttk.Label(file_frame, text="Sold Homes List:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(file_frame, textvariable=self.sold_csv_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_sold_csv).grid(row=1, column=2)
        
        # --- API Keys Frame ---
        api_frame = ttk.LabelFrame(main_frame, text="🔑 API Keys", padding="10")
        api_frame.pack(fill=tk.X, pady=(0, 10))
        
        # TomTom API Key
        ttk.Label(api_frame, text="TomTom API Key:").grid(row=0, column=0, sticky=tk.W, pady=5)
        tomtom_entry = ttk.Entry(api_frame, textvariable=self.tomtom_api_key, width=60, show="*")
        tomtom_entry.grid(row=0, column=1, padx=5)
        ttk.Button(api_frame, text="👁", width=3, command=lambda: self.toggle_visibility(tomtom_entry)).grid(row=0, column=2)
        
        # Mapbox Token
        ttk.Label(api_frame, text="Mapbox Token:").grid(row=1, column=0, sticky=tk.W, pady=5)
        mapbox_entry = ttk.Entry(api_frame, textvariable=self.mapbox_token, width=60, show="*")
        mapbox_entry.grid(row=1, column=1, padx=5)
        ttk.Button(api_frame, text="👁", width=3, command=lambda: self.toggle_visibility(mapbox_entry)).grid(row=1, column=2)
        
        # --- Settings Frame ---
        settings_frame = ttk.LabelFrame(main_frame, text="⚙️ Settings", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Number of nearby properties
        nearby_frame = ttk.Frame(settings_frame)
        nearby_frame.pack(fill=tk.X, pady=5)
        ttk.Label(nearby_frame, text="Nearby sold properties to show:").pack(side=tk.LEFT)
        nearby_slider = ttk.Scale(nearby_frame, from_=3, to=5, variable=self.num_nearby, 
                                   orient=tk.HORIZONTAL, length=150, command=self.update_nearby_label)
        nearby_slider.pack(side=tk.LEFT, padx=10)
        self.nearby_label = ttk.Label(nearby_frame, text="3", font=('Segoe UI', 11, 'bold'))
        self.nearby_label.pack(side=tk.LEFT)
        
        # Number of clients to process
        clients_frame = ttk.Frame(settings_frame)
        clients_frame.pack(fill=tk.X, pady=5)
        ttk.Label(clients_frame, text="Clients to process:").pack(side=tk.LEFT)
        ttk.Entry(clients_frame, textvariable=self.num_clients, width=10).pack(side=tk.LEFT, padx=10)
        ttk.Label(clients_frame, text="(enter number or 'all')", 
                  font=('Segoe UI', 9), foreground='gray').pack(side=tk.LEFT)
        
        # --- Progress Frame ---
        progress_frame = ttk.LabelFrame(main_frame, text="📊 Progress", padding="10")
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.status_label = ttk.Label(progress_frame, text="Ready to generate mailers", 
                                       font=('Segoe UI', 10))
        self.status_label.pack(anchor=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=680)
        self.progress_bar.pack(fill=tk.X, pady=10)
        
        self.detail_label = ttk.Label(progress_frame, text="", font=('Segoe UI', 9), foreground='gray')
        self.detail_label.pack(anchor=tk.W)
        
        # --- Buttons Frame ---
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(5, 10))
        
        self.generate_btn = ttk.Button(button_frame, text="🚀 Generate Mailers", 
                                        command=self.start_generation)
        self.generate_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="📂 Open Output Folder", 
                   command=self.open_output_folder).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="🗑️ Clear Cache", 
                   command=self.clear_cache).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="📋 Clear Log", 
                   command=self.clear_log).pack(side=tk.RIGHT, padx=5)
        
        # --- Log Frame ---
        log_frame = ttk.LabelFrame(main_frame, text="📝 Detailed Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create scrolled text widget for logs
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, font=('Consolas', 9),
                                                   wrap=tk.WORD, state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Configure log colors
        self.log_text.tag_configure('INFO', foreground='black')
        self.log_text.tag_configure('SUCCESS', foreground='green')
        self.log_text.tag_configure('WARNING', foreground='orange')
        self.log_text.tag_configure('ERROR', foreground='red')
        self.log_text.tag_configure('TIMESTAMP', foreground='gray')
        
        # --- Stats Frame ---
        stats_frame = ttk.LabelFrame(main_frame, text="📈 Results", padding="10")
        stats_frame.pack(fill=tk.X)
        
        self.stats_label = ttk.Label(stats_frame, text="No mailers generated yet", 
                                      font=('Segoe UI', 10))
        self.stats_label.pack(anchor=tk.W)
    
    def log(self, message, level='INFO'):
        """Add a message to the log with timestamp"""
        self.log_text.config(state='normal')
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert(tk.END, f"[{timestamp}] ", 'TIMESTAMP')
        self.log_text.insert(tk.END, f"{message}\n", level)
        self.log_text.see(tk.END)  # Auto-scroll to bottom
        self.log_text.config(state='disabled')
        self.root.update_idletasks()
    
    def clear_log(self):
        """Clear the log text area"""
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        self.log("Log cleared")
    
    def update_nearby_label(self, val):
        self.nearby_label.config(text=str(int(float(val))))
    
    def toggle_visibility(self, entry_widget):
        """Toggle password visibility for API key fields"""
        if entry_widget['show'] == '*':
            entry_widget.config(show='')
        else:
            entry_widget.config(show='*')
    
    def browse_client_csv(self):
        path = filedialog.askopenfilename(
            title="Select Client List CSV",
            initialdir=SCRIPT_DIR,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self.client_csv_path.set(path)
            self.log(f"Selected client CSV: {os.path.basename(path)}")
    
    def browse_sold_csv(self):
        path = filedialog.askopenfilename(
            title="Select Sold Homes CSV",
            initialdir=SCRIPT_DIR,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            self.sold_csv_path.set(path)
            self.log(f"Selected sold homes CSV: {os.path.basename(path)}")
    
    def open_output_folder(self):
        ensure_directories()
        if sys.platform == 'win32':
            os.startfile(OUTPUT_DIR)
        elif sys.platform == 'darwin':
            os.system(f'open "{OUTPUT_DIR}"')
        else:
            os.system(f'xdg-open "{OUTPUT_DIR}"')
        self.log(f"Opened output folder: {OUTPUT_DIR}")
    
    def clear_cache(self):
        if messagebox.askyesno("Clear Cache", "Clear the geocoding cache? This will require re-geocoding all addresses."):
            if os.path.exists(CACHE_FILE):
                os.remove(CACHE_FILE)
            self.cache = {}
            self.log("Geocoding cache cleared", 'WARNING')
            messagebox.showinfo("Cache Cleared", "Geocoding cache has been cleared.")
    
    def start_generation(self):
        # Validate inputs
        if not self.client_csv_path.get():
            messagebox.showerror("Error", "Please select a Client List CSV file.")
            self.log("Error: No client CSV selected", 'ERROR')
            return
        if not self.sold_csv_path.get():
            messagebox.showerror("Error", "Please select a Sold Homes CSV file.")
            self.log("Error: No sold homes CSV selected", 'ERROR')
            return
        
        # Disable button during processing
        self.generate_btn.config(state='disabled')
        self.log("=" * 50)
        self.log("Starting mailer generation...")
        
        # Start processing in a separate thread
        thread = threading.Thread(target=self.generate_mailers)
        thread.start()
    
    def update_status(self, text):
        self.status_label.config(text=text)
        self.root.update_idletasks()
    
    def update_detail(self, text):
        self.detail_label.config(text=text)
        self.root.update_idletasks()
    
    def update_progress(self, value, maximum=100):
        self.progress_bar['maximum'] = maximum
        self.progress_bar['value'] = value
        self.root.update_idletasks()
    
    def get_coords_tomtom(self, row, index, total, list_type="Client"):
        address = str(row['Address']).strip()
        city = str(row.get('City', 'Bakersfield')).strip()
        zip_code = str(row.get('ZIP', '')).split('.')[0].strip()
        full_address = f"{address}, {city}, CA {zip_code}"
        
        self.update_detail(f"Geocoding {list_type} {index + 1}/{total}: {address[:40]}...")
        
        if full_address in self.cache:
            return self.cache[full_address]
        
        try:
            url = f"https://api.tomtom.com/search/2/geocode/{requests.utils.quote(full_address)}.json?key={self.tomtom_api_key.get()}"
            data = requests.get(url, timeout=10).json()
            if data.get('results'):
                pos = data['results'][0]['position']
                coords = [pos['lat'], pos['lon']]
                self.cache[full_address] = coords
                save_cache(self.cache)
                self.log(f"  Geocoded: {address[:35]}... -> ({coords[0]:.4f}, {coords[1]:.4f})")
                return coords
        except Exception as e:
            self.log(f"  Failed to geocode: {address[:35]}... ({str(e)})", 'WARNING')
        
        self.skipped_log.append({'Address': full_address, 'Type': list_type, 'Reason': 'Geocoding failed'})
        return None
    
    def find_nearest_sold(self, client_coords, sold_df, n=3):
        sold_pool = sold_df.copy()
        sold_pool['distance'] = sold_pool['coords'].apply(lambda x: geodesic(client_coords, x).miles)
        return sold_pool[sold_pool['distance'] > 0.005].sort_values('distance').head(n).to_dict('records')
    
    def generate_mailers(self):
        try:
            self.skipped_log = []
            
            # --- Ensure directories exist ---
            self.log("Creating output directories...")
            ensure_directories()
            self.log(f"  Output: {OUTPUT_DIR}", 'SUCCESS')
            self.log(f"  Individual PDFs: {INDIVIDUAL_DIR}", 'SUCCESS')
            self.log(f"  Debug Maps: {MAP_DEBUG_DIR}", 'SUCCESS')
            
            # --- STEP 1: Load Data ---
            self.update_status("📥 Loading CSV data...")
            self.update_progress(0)
            self.log("Loading CSV files...")
            
            df_clients = pd.read_csv(self.client_csv_path.get())
            df_sold = pd.read_csv(self.sold_csv_path.get())
            
            # Clean garbage rows
            df_clients = df_clients[df_clients['Address'].str.contains("The information", na=False) == False]
            df_sold = df_sold[df_sold['Address'].str.contains("The information", na=False) == False]
            
            # Limit clients if specified
            num_clients_str = self.num_clients.get().strip().lower()
            if num_clients_str != 'all' and num_clients_str != '':
                try:
                    limit = int(num_clients_str)
                    df_clients = df_clients.head(limit)
                    self.log(f"  Limited to first {limit} clients")
                except:
                    pass
            
            total_clients = len(df_clients)
            total_sold = len(df_sold)
            
            self.log(f"  Loaded {total_clients} clients from {os.path.basename(self.client_csv_path.get())}", 'SUCCESS')
            self.log(f"  Loaded {total_sold} sold properties from {os.path.basename(self.sold_csv_path.get())}", 'SUCCESS')
            self.update_detail(f"Loaded {total_clients} clients and {total_sold} sold properties")
            
            # --- STEP 2: Geocoding ---
            self.update_status("🗺️ Geocoding addresses...")
            self.log("Geocoding client addresses...")
            
            # Geocode clients
            client_coords = []
            cached_count = 0
            for i, row in df_clients.iterrows():
                full_address = f"{row['Address']}, {row.get('City', 'Bakersfield')}, CA {str(row.get('ZIP', '')).split('.')[0]}"
                was_cached = full_address in self.cache
                coords = self.get_coords_tomtom(row, len(client_coords), total_clients, "Client")
                client_coords.append(coords)
                if was_cached:
                    cached_count += 1
                self.update_progress(len(client_coords), total_clients + total_sold)
            df_clients['coords'] = client_coords
            self.log(f"  Client geocoding complete ({cached_count} from cache)", 'SUCCESS')
            
            # Geocode sold properties
            self.log("Geocoding sold property addresses...")
            sold_coords = []
            cached_count = 0
            for i, row in df_sold.iterrows():
                full_address = f"{row['Address']}, {row.get('City', 'Bakersfield')}, CA {str(row.get('ZIP', '')).split('.')[0]}"
                was_cached = full_address in self.cache
                coords = self.get_coords_tomtom(row, len(sold_coords), total_sold, "Sold")
                sold_coords.append(coords)
                if was_cached:
                    cached_count += 1
                self.update_progress(total_clients + len(sold_coords), total_clients + total_sold)
            df_sold['coords'] = sold_coords
            self.log(f"  Sold property geocoding complete ({cached_count} from cache)", 'SUCCESS')
            
            valid_clients = df_clients.dropna(subset=['coords']).copy()
            valid_sold = df_sold.dropna(subset=['coords']).copy()
            
            self.log(f"Valid addresses: {len(valid_clients)} clients, {len(valid_sold)} sold properties")
            
            # --- STEP 3: Generate PDFs ---
            self.update_status("📄 Generating mailers...")
            self.update_progress(0)
            self.log("Generating PDF mailers...")
            
            template = Template(html_template_str)
            pdf_files = []
            num_nearby = int(self.num_nearby.get())
            
            for idx, (index, client) in enumerate(valid_clients.iterrows()):
                nearby = self.find_nearest_sold(client['coords'], valid_sold, n=num_nearby)
                lat, lon = client['coords']
                
                # Build Mapbox URL with markers
                markers = f"pin-l+e74c3c({lon},{lat})"
                if nearby:
                    for home in nearby:
                        h_lat, h_lon = home['coords']
                        markers += f",pin-s+27ae60({h_lon},{h_lat})"
                
                map_url = f"https://api.mapbox.com/styles/v1/mapbox/streets-v12/static/{markers}/{lon},{lat},14/500x400?access_token={self.mapbox_token.get()}"
                
                # Save map image for verification
                try:
                    img_response = requests.get(map_url, timeout=15)
                    if img_response.status_code == 200:
                        first_name = str(client.get('Primary First', 'Unknown')).strip()
                        last_name = str(client.get('Primary Last', '')).strip()
                        safe_name = f"{first_name}_{last_name}".replace(" ", "_")
                        map_path = os.path.join(MAP_DEBUG_DIR, f"map_{safe_name}.png")
                        with open(map_path, "wb") as f:
                            f.write(img_response.content)
                except Exception as e:
                    self.log(f"  Warning: Could not save map image for {client.get('Primary First', 'Unknown')}: {e}", 'WARNING')
                
                # Get first name, handle empty
                first_name = str(client.get('Primary First', '')).strip()
                if not first_name or first_name.lower() == 'nan':
                    first_name = 'Neighbor'
                else:
                    first_name = first_name.capitalize()
                
                html_out = template.render(
                    first_name=first_name,
                    address=client['Address'],
                    nearby=nearby,
                    map_url=map_url
                )
                
                # Use absolute path for PDF output
                file_path = os.path.join(INDIVIDUAL_DIR, f"mailer_{idx}.pdf")
                HTML(string=html_out).write_pdf(file_path)
                pdf_files.append(file_path)
                
                # Log each mailer
                nearby_addrs = [p['Address'][:25] for p in nearby[:2]]
                self.log(f"  [{idx+1}/{len(valid_clients)}] {first_name} @ {client['Address'][:30]}...")
                
                self.update_detail(f"Created mailer {len(pdf_files)}/{len(valid_clients)}: {first_name}")
                self.update_progress(len(pdf_files), len(valid_clients))
            
            self.log(f"Generated {len(pdf_files)} individual PDFs", 'SUCCESS')
            
            # --- STEP 4: Merge PDFs ---
            self.update_status("📑 Merging PDFs...")
            self.log("Merging PDFs into single file...")
            if pdf_files:
                merger = PdfWriter()
                for pdf in pdf_files:
                    merger.append(pdf)
                final_path = os.path.join(OUTPUT_DIR, FINAL_PDF)
                with open(final_path, "wb") as f:
                    merger.write(f)
                self.log(f"  Merged PDF: {final_path}", 'SUCCESS')
            
            # --- STEP 5: Generate Error Report ---
            if self.skipped_log:
                skipped_df = pd.DataFrame(self.skipped_log)
                skipped_path = os.path.join(OUTPUT_DIR, SKIPPED_REPORT)
                skipped_df.to_csv(skipped_path, index=False)
                self.log(f"  Skipped addresses report: {skipped_path}", 'WARNING')
            
            # --- Done ---
            self.update_status("✅ Complete!")
            self.update_progress(100)
            self.update_detail("")
            
            self.log("=" * 50)
            self.log(f"COMPLETE: Generated {len(pdf_files)} mailers!", 'SUCCESS')
            
            stats_text = f"✅ Generated {len(pdf_files)} mailers\n"
            stats_text += f"📄 Output: {os.path.join(OUTPUT_DIR, FINAL_PDF)}\n"
            if self.skipped_log:
                stats_text += f"⚠️ Skipped {len(self.skipped_log)} addresses (see {SKIPPED_REPORT})"
            
            self.stats_label.config(text=stats_text)
            
            messagebox.showinfo("Success!", f"Successfully generated {len(pdf_files)} mailers!\n\nOutput saved to:\n{OUTPUT_DIR}")
            
        except Exception as e:
            import traceback
            error_msg = str(e)
            self.update_status(f"❌ Error: {error_msg}")
            self.log(f"ERROR: {error_msg}", 'ERROR')
            self.log(traceback.format_exc(), 'ERROR')
            messagebox.showerror("Error", f"An error occurred:\n{error_msg}")
        
        finally:
            self.generate_btn.config(state='normal')


def main():
    root = tk.Tk()
    
    # Set theme
    style = ttk.Style()
    if 'vista' in style.theme_names():
        style.theme_use('vista')
    elif 'clam' in style.theme_names():
        style.theme_use('clam')
    
    app = MailerGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
