"""
Real Estate Mailer Generator - Mapbox Edition
Uses Mapbox for both geocoding and static maps
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
from dotenv import load_dotenv

# --- GET SCRIPT DIRECTORY ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# --- LOAD ENVIRONMENT VARIABLES ---
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
DEFAULT_MAPBOX_TOKEN = os.getenv('MAPBOX_TOKEN', '')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'output')
INDIVIDUAL_DIR = os.path.join(OUTPUT_DIR, 'individual')
MAP_DEBUG_DIR = os.path.join(OUTPUT_DIR, 'debug_maps')
CACHE_FILE = os.path.join(SCRIPT_DIR, 'geocoding_cache_mapbox.json')
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


# --- PROFESSIONAL BRANDED TEMPLATE (ONE PAGE) ---
html_template_str = """
<!DOCTYPE html>
<html>
<head>
    <style>
        @page { 
            size: 6in 9in; 
            margin: 0; 
        }
        * {
            box-sizing: border-box;
        }
        body { 
            font-family: 'Georgia', 'Times New Roman', serif;
            margin: 0;
            padding: 0;
            width: 6in;
            height: 9in;
            background: #ffffff;
            color: #2c2c2c;
            display: flex;
            flex-direction: column;
        }
        
        /* Header with branding */
        .header {
            background: linear-gradient(135deg, #8B4513 0%, #A0522D 50%, #D4AF37 100%);
            padding: 12px 20px;
            text-align: center;
            color: white;
        }
        .logo-text {
            font-size: 22px;
            font-weight: bold;
            letter-spacing: 2px;
            margin: 0;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        .logo-subtitle {
            font-size: 9px;
            letter-spacing: 1.5px;
            margin-top: 2px;
            color: #f0e68c;
        }
        
        /* Main content area */
        .content {
            flex: 1;
            padding: 15px 20px;
            display: flex;
            flex-direction: column;
        }
        
        .greeting-section {
            margin-bottom: 10px;
            border-left: 3px solid #8B4513;
            padding-left: 10px;
        }
        .greeting {
            font-size: 16px;
            color: #8B4513;
            margin: 0 0 5px 0;
            font-weight: bold;
        }
        .intro-text {
            font-size: 10px;
            line-height: 1.4;
            color: #444;
        }
        .address-highlight {
            color: #8B4513;
            font-weight: bold;
        }
        
        /* Map section */
        .map-container {
            margin: 10px 0;
            text-align: center;
            background: #f8f8f8;
            padding: 8px;
            border-radius: 6px;
            border: 2px solid #D4AF37;
        }
        .map-box {
            width: 100%;
            height: 180px;
            border-radius: 5px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        }
        .map-box img {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }
        .map-legend {
            margin-top: 5px;
            font-size: 8px;
            color: #666;
            font-style: italic;
        }
        .legend-red { color: #c0392b; font-weight: bold; }
        .legend-green { color: #27ae60; font-weight: bold; }
        
        /* Property listings */
        .section-title {
            font-size: 11px;
            color: #8B4513;
            font-weight: bold;
            margin: 10px 0 6px 0;
            padding-bottom: 4px;
            border-bottom: 1.5px solid #D4AF37;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        .property-list {
            margin-bottom: 8px;
        }
        .property-item {
            background: linear-gradient(to right, #f9f9f9 0%, #ffffff 100%);
            border-left: 3px solid #27ae60;
            padding: 6px 10px;
            margin-bottom: 5px;
            border-radius: 0 5px 5px 0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        .property-address {
            font-size: 10px;
            font-weight: bold;
            color: #2c2c2c;
            margin-bottom: 3px;
        }
        .property-details {
            font-size: 8px;
            color: #555;
            line-height: 1.3;
        }
        .property-price {
            color: #27ae60;
            font-weight: bold;
            font-size: 10px;
        }
        .property-distance {
            color: #888;
            font-style: italic;
        }
        
        /* Market insight box */
        .insight-box {
            background: linear-gradient(135deg, #f0e68c 0%, #ffd700 100%);
            padding: 8px 10px;
            border-radius: 6px;
            margin: 8px 0;
            border: 1.5px solid #D4AF37;
        }
        .insight-title {
            font-size: 10px;
            font-weight: bold;
            color: #8B4513;
            margin: 0 0 4px 0;
        }
        .insight-text {
            font-size: 8px;
            color: #2c2c2c;
            line-height: 1.3;
        }
        
        /* Footer/Contact section */
        .footer {
            background: linear-gradient(135deg, #2c2c2c 0%, #4a4a4a 100%);
            color: white;
            padding: 12px 20px;
            margin-top: auto;
        }
        .agent-name {
            font-size: 14px;
            font-weight: bold;
            color: #D4AF37;
            margin: 0 0 2px 0;
        }
        .company-name {
            font-size: 9px;
            color: #f0e68c;
            margin: 0 0 6px 0;
            letter-spacing: 1px;
        }
        .contact-details {
            font-size: 8px;
            line-height: 1.4;
        }
        .contact-item {
            margin: 2px 0;
            display: inline-block;
            width: 48%;
        }
        .contact-icon {
            color: #D4AF37;
            margin-right: 3px;
        }
        .cta-text {
            text-align: center;
            font-size: 10px;
            margin-top: 8px;
            padding-top: 8px;
            border-top: 1px solid #666;
            color: #D4AF37;
            font-weight: bold;
        }
    </style>
</head>
<body>
    <!-- Header -->
    <div class="header">
        <div class="logo-text">GEBARAH</div>
        <div class="logo-subtitle">REAL ESTATE GROUP</div>
    </div>
    
    <!-- Main Content -->
    <div class="content">
        <!-- Greeting -->
        <div class="greeting-section">
            <div class="greeting">Hello {{ first_name }},</div>
            <div class="intro-text">
                Your neighborhood is experiencing significant real estate activity! 
                Homes near <span class="address-highlight">{{ address }}</span> are selling quickly.
            </div>
        </div>
        
        <!-- Map -->
        <div class="map-container">
            <div class="map-box">
                <img src="{{ map_url }}" alt="Neighborhood Map">
            </div>
            <div class="map-legend">
                <span class="legend-red">● Your Home</span> | 
                <span class="legend-green">● Recent Sales</span>
            </div>
        </div>
        
        <!-- Recent Sales -->
        <div class="section-title">📊 Recent Neighborhood Sales</div>
        <div class="property-list">
            {% for property in nearby %}
            <div class="property-item">
                <div class="property-address">{{ property.Address }}</div>
                <div class="property-details">
                    <span class="property-price">${{ "{:,.0f}".format(property['Purchase Amt']) }}</span>
                    <span class="property-distance"> • {{ "{:.2f}".format(property.distance) }} mi away</span>
                    {% if property.get('Beds') and property.get('Baths') and property.get('Sq Ft') %}
                    <span> • {{ property.Beds }}bd/{{ property.Baths }}ba, {{ "{:,.0f}".format(property['Sq Ft']) }} sqft</span>
                    {% endif %}
                </div>
            </div>
            {% endfor %}
        </div>
        
        <!-- Market Insight -->
        <div class="insight-box">
            <div class="insight-title">💡 What Does This Mean For You?</div>
            <div class="insight-text">
                With {{ nearby|length }} recent sales in your immediate area, now is an excellent time to 
                understand your home's current market value. Contact me for a complimentary market analysis!
            </div>
        </div>
    </div>
    
    <!-- Footer -->
    <div class="footer">
        <div>
            <div class="agent-name">Joy Gebarah</div>
            <div class="company-name">GEBARAH REAL ESTATE GROUP</div>
            <div class="contact-details">
                <div class="contact-item">
                    <span class="contact-icon">📞</span> Licensed Professional
                </div>
                <div class="contact-item">
                    <span class="contact-icon">⭐</span> 5.0 Rating
                </div>
                <div class="contact-item">
                    <span class="contact-icon">🏆</span> Bakersfield Specialist
                </div>
                <div class="contact-item">
                    <span class="contact-icon">🗣️</span> Multilingual Services
                </div>
            </div>
        </div>
        <div class="cta-text">
            📧 Contact me today for your FREE home valuation!
        </div>
    </div>
</body>
</html>
"""


class MailerGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Real Estate Mailer Generator - Mapbox Edition")
        self.root.geometry("750x750")
        self.root.resizable(True, True)
        
        # Variables
        self.client_csv_path = tk.StringVar()
        self.sold_csv_path = tk.StringVar()
        self.num_nearby = tk.IntVar(value=3)
        self.num_clients = tk.StringVar(value="all")
        self.mapbox_token = tk.StringVar(value=DEFAULT_MAPBOX_TOKEN)
        self.cache = load_cache()
        self.skipped_log = []
        
        self.setup_ui()
        self.log("Mapbox Edition - Application started")
        self.log(f"Working directory: {SCRIPT_DIR}")
        self.log(f"Output directory: {OUTPUT_DIR}")
    
    def setup_ui(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="🏠 Real Estate Mailer Generator (Mapbox)", 
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
        
        # --- API Key Frame ---
        api_frame = ttk.LabelFrame(main_frame, text="🔑 Mapbox API Token", padding="10")
        api_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Mapbox Token
        ttk.Label(api_frame, text="Mapbox Token:").grid(row=0, column=0, sticky=tk.W, pady=5)
        mapbox_entry = ttk.Entry(api_frame, textvariable=self.mapbox_token, width=60, show="*")
        mapbox_entry.grid(row=0, column=1, padx=5)
        ttk.Button(api_frame, text="👁", width=3, command=lambda: self.toggle_visibility(mapbox_entry)).grid(row=0, column=2)
        
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
        self.log_text.see(tk.END)
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
        self.log("Starting mailer generation with Mapbox...")
        
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
    
    def get_coords_mapbox(self, row, index, total, list_type="Client"):
        """Geocode using Mapbox Geocoding API"""
        address = str(row['Address']).strip()
        city = str(row.get('City', 'Bakersfield')).strip()
        zip_code = str(row.get('ZIP', '')).split('.')[0].strip()
        full_address = f"{address}, {city}, CA {zip_code}"
        
        self.update_detail(f"Geocoding {list_type} {index + 1}/{total}: {address[:40]}...")
        
        if full_address in self.cache:
            return self.cache[full_address]
        
        try:
            # Mapbox Geocoding API
            url = f"https://api.mapbox.com/geocoding/v5/mapbox.places/{requests.utils.quote(full_address)}.json"
            params = {
                'access_token': self.mapbox_token.get(),
                'limit': 1,
                'country': 'US'
            }
            response = requests.get(url, params=params, timeout=10)
            data = response.json()
            
            if data.get('features'):
                coords_long_lat = data['features'][0]['center']  # [longitude, latitude]
                coords = [coords_long_lat[1], coords_long_lat[0]]  # Convert to [latitude, longitude]
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
            
            # --- STEP 2: Geocoding with Mapbox ---
            self.update_status("🗺️ Geocoding addresses with Mapbox...")
            self.log("Geocoding client addresses using Mapbox...")
            
            # Geocode clients
            client_coords = []
            cached_count = 0
            for i, row in df_clients.iterrows():
                full_address = f"{row['Address']}, {row.get('City', 'Bakersfield')}, CA {str(row.get('ZIP', '')).split('.')[0]}"
                was_cached = full_address in self.cache
                coords = self.get_coords_mapbox(row, len(client_coords), total_clients, "Client")
                client_coords.append(coords)
                if was_cached:
                    cached_count += 1
                self.update_progress(len(client_coords), total_clients + total_sold)
            df_clients['coords'] = client_coords
            self.log(f"  Client geocoding complete ({cached_count} from cache)", 'SUCCESS')
            
            # Geocode sold properties
            self.log("Geocoding sold property addresses using Mapbox...")
            sold_coords = []
            cached_count = 0
            for i, row in df_sold.iterrows():
                full_address = f"{row['Address']}, {row.get('City', 'Bakersfield')}, CA {str(row.get('ZIP', '')).split('.')[0]}"
                was_cached = full_address in self.cache
                coords = self.get_coords_mapbox(row, len(sold_coords), total_sold, "Sold")
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
            self.update_status("📄 Generating professional branded mailers...")
            self.update_progress(0)
            self.log("Generating PDF mailers with Joy Gebarah branding...")
            
            template = Template(html_template_str)
            pdf_files = []
            num_nearby = int(self.num_nearby.get())
            
            for idx, (index, client) in enumerate(valid_clients.iterrows()):
                nearby = self.find_nearest_sold(client['coords'], valid_sold, n=num_nearby)
                lat, lon = client['coords']
                
                # Build Mapbox Static Images API URL with markers
                markers = f"pin-l+c0392b({lon},{lat})"  # Red pin for client
                if nearby:
                    for home in nearby:
                        h_lat, h_lon = home['coords']
                        markers += f",pin-s+27ae60({h_lon},{h_lat})"  # Green pins for sold homes
                
                map_url = f"https://api.mapbox.com/styles/v1/mapbox/streets-v12/static/{markers}/{lon},{lat},14,0/500x400@2x?access_token={self.mapbox_token.get()}"
                
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
            self.log(f"COMPLETE: Generated {len(pdf_files)} professional mailers!", 'SUCCESS')
            
            stats_text = f"✅ Generated {len(pdf_files)} mailers\n"
            stats_text += f"📄 Output: {os.path.join(OUTPUT_DIR, FINAL_PDF)}\n"
            if self.skipped_log:
                stats_text += f"⚠️ Skipped {len(self.skipped_log)} addresses (see {SKIPPED_REPORT})"
            
            self.stats_label.config(text=stats_text)
            
            messagebox.showinfo("Success!", f"Successfully generated {len(pdf_files)} professional mailers!\n\nOutput saved to:\n{OUTPUT_DIR}")
            
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
