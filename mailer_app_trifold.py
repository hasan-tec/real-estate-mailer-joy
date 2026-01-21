"""
Real Estate Mailer Generator - Tri-Fold Edition
8.5" x 11" Letter format for #10 Window Envelopes
Uses Mapbox for both geocoding and static maps
Supports custom banner uploads
"""

import pandas as pd
import os
import sys
import json
import requests
import time
import threading
import base64
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
FINAL_PDF = 'final_mailers_trifold.pdf'
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


def image_to_base64(image_path):
    """Convert an image file to base64 data URI"""
    if not image_path or not os.path.exists(image_path):
        return None
    
    # Determine MIME type based on extension
    ext = os.path.splitext(image_path)[1].lower()
    mime_types = {
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.webp': 'image/webp'
    }
    mime_type = mime_types.get(ext, 'image/png')
    
    with open(image_path, 'rb') as f:
        encoded = base64.b64encode(f.read()).decode('utf-8')
    
    return f"data:{mime_type};base64,{encoded}"


# --- TRI-FOLD TEMPLATE (8.5" x 11" Letter) ---
# Layout: Address panel (top) | Upload #1 + Map/Table | Upload #2 (bottom)
html_template_str = """
<!DOCTYPE html>
<html>
<head>
    <style>
        @page { 
            size: 8.5in 11in; 
            margin: 0; 
        }
        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }
        body { 
            font-family: 'Arial', 'Helvetica', sans-serif;
            width: 8.5in;
            height: 11in;
            background: #ffffff;
            color: #2c2c2c;
        }
        
        /* ============================================
           TRI-FOLD LAYOUT - Three panels fitting on one page
           Panel 1 (TOP): Address for window envelope
           Panel 2 (MIDDLE): Upload #1 + Map + Table
           Panel 3 (BOTTOM): Upload #2
           ============================================ */
        
        .panel {
            width: 100%;
            overflow: hidden;
            position: relative;
        }
        
        .panel-address {
            height: 3.5in;
        }
        
        .panel-content {
            height: 4.2in;
        }
        
        .panel-bottom {
            height: 3.3in;
        }
        
        /* ============================================
           PANEL 1: ADDRESS PANEL (TOP)
           Shows through #10 window envelope
           ============================================ */
        .panel.panel-address {
            background: #ffffff;
            display: flex;
            flex-direction: column;
            padding: 0.25in;
        }
        
        /* Small branding bar at very top */
        .mini-header {
            background: linear-gradient(135deg, #8B4513 0%, #A0522D 50%, #D4AF37 100%);
            color: white;
            padding: 8px 15px;
            text-align: center;
            font-size: 11px;
            font-weight: bold;
            letter-spacing: 2px;
            border-radius: 4px;
            margin-bottom: 0.15in;
            position: relative;
        }
        
        /* Address box - positioned for #10 window envelope */
        /* Window typically: 4.5" x 1.125", positioned ~0.875" from left, ~0.5" from bottom of panel */
        .address-container {
            margin-top: 0.4in;
            margin-left: 0.5in;
            width: 4in;
            height: 1.2in;
            padding: 10px 15px;
        }
        
        .address-line {
            font-size: 12pt;
            line-height: 1.5;
            color: #000;
        }
        
        .address-name {
            font-weight: bold;
            text-transform: uppercase;
        }
        
        /* Return address in corner of header */
        .return-address {
            position: absolute;
            top: 6px;
            right: 15px;
            font-size: 7pt;
            color: #ffd700;
            text-align: right;
            line-height: 1.3;
            max-width: 180px;
        }
        
        /* ============================================
           PANEL 2: MAIN CONTENT (MIDDLE)
           Upload #1 banner + Map + Property table
           ============================================ */
        .panel.panel-content {
            background: #ffffff;
            padding: 0.15in 0.25in;
            display: flex;
            flex-direction: column;
        }
        
        /* Top banner upload area */
        .top-banner {
            width: 100%;
            height: 1.1in;
            background: #f5f5f5;
            border: 2px dashed #D4AF37;
            border-radius: 6px;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
            margin-bottom: 0.1in;
        }
        
        .top-banner img {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }
        
        .top-banner.placeholder {
            color: #999;
            font-size: 11px;
            font-style: italic;
        }
        
        /* Map and Table side by side */
        .content-row {
            display: flex;
            gap: 0.15in;
            flex: 1;
        }
        
        .map-section {
            width: 3.5in;
            display: flex;
            flex-direction: column;
        }
        
        .map-box {
            flex: 1;
            border-radius: 6px;
            overflow: hidden;
            border: 2px solid #D4AF37;
            background: #f8f8f8;
        }
        
        .map-box img {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }
        
        .map-legend {
            font-size: 7pt;
            color: #666;
            text-align: center;
            margin-top: 3px;
        }
        
        .legend-red { color: #c0392b; font-weight: bold; }
        .legend-green { color: #27ae60; font-weight: bold; }
        
        /* Property table */
        .table-section {
            flex: 1;
            display: flex;
            flex-direction: column;
        }
        
        .section-title {
            font-size: 9pt;
            color: #8B4513;
            font-weight: bold;
            padding-bottom: 4px;
            border-bottom: 1.5px solid #D4AF37;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 6px;
        }
        
        .property-table {
            font-size: 7.5pt;
            width: 100%;
            border-collapse: collapse;
        }
        
        .property-table th {
            background: #8B4513;
            color: white;
            padding: 4px 6px;
            text-align: left;
            font-weight: bold;
        }
        
        .property-table td {
            padding: 4px 6px;
            border-bottom: 1px solid #eee;
        }
        
        .property-table tr:nth-child(even) {
            background: #f9f9f9;
        }
        
        .price-cell {
            color: #27ae60;
            font-weight: bold;
        }
        
        /* ============================================
           PANEL 3: BOTTOM BANNER (BOTTOM)
           Upload #2 - Full custom area
           ============================================ */
        .panel.panel-bottom {
            background: #ffffff;
            padding: 0.15in 0.25in;
            display: flex;
            flex-direction: column;
        }
        
        .bottom-banner {
            width: 100%;
            flex: 1;
            background: #f5f5f5;
            border: 2px dashed #D4AF37;
            border-radius: 6px;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: hidden;
        }
        
        .bottom-banner img {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }
        
        .bottom-banner.placeholder {
            color: #999;
            font-size: 14px;
            font-style: italic;
        }
        
        /* Small footer with agent info */
        .mini-footer {
            margin-top: 0.1in;
            background: linear-gradient(135deg, #2c2c2c 0%, #4a4a4a 100%);
            color: white;
            padding: 8px 15px;
            border-radius: 4px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 8pt;
        }
        
        .mini-footer .agent-name {
            color: #D4AF37;
            font-weight: bold;
            font-size: 10pt;
        }
        
        .mini-footer .contact-info {
            color: #ccc;
        }
        
        /* Fold guide lines (for printing reference - optional) */
        .fold-line {
            position: absolute;
            left: 0;
            right: 0;
            border-top: 1px dashed #ccc;
            opacity: 0.3;
        }
        
    </style>
</head>
<body>
    <!-- ==========================================
         PANEL 1: ADDRESS (TOP) - Shows through window
         ========================================== -->
    <div class="panel panel-address">
        <div class="mini-header">GEBARAH REAL ESTATE GROUP</div>
        
        <div class="return-address">
            Joy Gebarah<br>
            Gebarah Real Estate Group<br>
            Bakersfield, CA
        </div>
        
        <div class="address-container">
            <div class="address-line address-name">{{ first_name }} {{ last_name }}</div>
            <div class="address-line">{{ address }}</div>
            <div class="address-line">{{ city }}, CA {{ zip_code }}</div>
        </div>
    </div>
    
    <!-- ==========================================
         PANEL 2: MAIN CONTENT (MIDDLE)
         ========================================== -->
    <div class="panel panel-content">
        <!-- Top Banner Upload Area -->
        {% if top_banner_img %}
        <div class="top-banner">
            <img src="{{ top_banner_img }}" alt="Top Banner">
        </div>
        {% else %}
        <div class="top-banner placeholder">
            [Top Banner - Upload your custom Canva design]
        </div>
        {% endif %}
        
        <!-- Map and Property Table -->
        <div class="content-row">
            <!-- Map -->
            <div class="map-section">
                <div class="map-box">
                    <img src="{{ map_url }}" alt="Neighborhood Map">
                </div>
                <div class="map-legend">
                    <span class="legend-red">● Your Home</span> &nbsp;|&nbsp; 
                    <span class="legend-green">● Recent Sales</span>
                </div>
            </div>
            
            <!-- Property Table -->
            <div class="table-section">
                <div class="section-title">📊 Recent Nearby Sales</div>
                <table class="property-table">
                    <tr>
                        <th>Address</th>
                        <th>Price</th>
                        <th>Bed/Bath</th>
                        <th>Sq Ft</th>
                    </tr>
                    {% for property in nearby %}
                    <tr>
                        <td>{{ property.Address[:25] }}{% if property.Address|length > 25 %}...{% endif %}</td>
                        <td class="price-cell">${{ "{:,.0f}".format(property['Purchase Amt'] / 1000) }}k</td>
                        <td>{{ property.Beds }}/{{ property.Baths }}</td>
                        <td>{{ "{:,.0f}".format(property['Sq Ft']) }}</td>
                    </tr>
                    {% endfor %}
                </table>
            </div>
        </div>
    </div>
    
    <!-- ==========================================
         PANEL 3: BOTTOM BANNER (BOTTOM)
         ========================================== -->
    <div class="panel panel-bottom">
        {% if bottom_banner_img %}
        <div class="bottom-banner">
            <img src="{{ bottom_banner_img }}" alt="Bottom Banner">
        </div>
        {% else %}
        <div class="bottom-banner placeholder">
            [Bottom Banner - Upload your custom Canva design / Call-to-Action]
        </div>
        {% endif %}
        
        <div class="mini-footer">
            <div>
                <span class="agent-name">Joy Gebarah</span> | Gebarah Real Estate Group
            </div>
            <div class="contact-info">
                📞 Licensed Professional &nbsp;|&nbsp; ⭐ 5.0 Rating &nbsp;|&nbsp; 🏆 Bakersfield Specialist
            </div>
        </div>
    </div>
</body>
</html>
"""


class MailerGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Real Estate Mailer Generator - Tri-Fold Edition")
        self.root.geometry("800x850")
        self.root.resizable(True, True)
        
        # Variables
        self.client_csv_path = tk.StringVar()
        self.sold_csv_path = tk.StringVar()
        self.top_banner_path = tk.StringVar()
        self.bottom_banner_path = tk.StringVar()
        self.num_nearby = tk.IntVar(value=3)
        self.num_clients = tk.StringVar(value="all")
        self.mapbox_token = tk.StringVar(value=DEFAULT_MAPBOX_TOKEN)
        self.cache = load_cache()
        self.skipped_log = []
        
        self.setup_ui()
        self.log("Tri-Fold Edition - Application started")
        self.log(f"Output: 8.5\" x 11\" letter format for #10 window envelopes")
        self.log(f"Working directory: {SCRIPT_DIR}")
    
    def setup_ui(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="🏠 Real Estate Mailer Generator (Tri-Fold)", 
                                font=('Segoe UI', 16, 'bold'))
        title_label.pack(pady=(0, 10))
        
        subtitle = ttk.Label(main_frame, text="8.5\" x 11\" Letter | #10 Window Envelope Compatible",
                            font=('Segoe UI', 10), foreground='gray')
        subtitle.pack(pady=(0, 15))
        
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
        
        # --- Banner Upload Frame ---
        banner_frame = ttk.LabelFrame(main_frame, text="🎨 Custom Banners (Canva Uploads)", padding="10")
        banner_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Top Banner
        ttk.Label(banner_frame, text="Top Banner Image:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(banner_frame, textvariable=self.top_banner_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(banner_frame, text="Browse", command=self.browse_top_banner).grid(row=0, column=2)
        ttk.Button(banner_frame, text="Clear", command=lambda: self.top_banner_path.set("")).grid(row=0, column=3, padx=5)
        
        # Bottom Banner
        ttk.Label(banner_frame, text="Bottom Banner Image:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(banner_frame, textvariable=self.bottom_banner_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(banner_frame, text="Browse", command=self.browse_bottom_banner).grid(row=1, column=2)
        ttk.Button(banner_frame, text="Clear", command=lambda: self.bottom_banner_path.set("")).grid(row=1, column=3, padx=5)
        
        # Banner tips
        tip_label = ttk.Label(banner_frame, 
                              text="💡 Tip: Create banners at 8\" x 1\" (top) and 8\" x 3\" (bottom) for best results",
                              font=('Segoe UI', 9), foreground='#666')
        tip_label.grid(row=2, column=0, columnspan=4, pady=(5, 0), sticky=tk.W)
        
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
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=730)
        self.progress_bar.pack(fill=tk.X, pady=10)
        
        self.detail_label = ttk.Label(progress_frame, text="", font=('Segoe UI', 9), foreground='gray')
        self.detail_label.pack(anchor=tk.W)
        
        # --- Buttons Frame ---
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(5, 10))
        
        self.generate_btn = ttk.Button(button_frame, text="🚀 Generate Tri-Fold Mailers", 
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
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, font=('Consolas', 9),
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
    
    def browse_top_banner(self):
        path = filedialog.askopenfilename(
            title="Select Top Banner Image",
            initialdir=SCRIPT_DIR,
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.webp"), ("All files", "*.*")]
        )
        if path:
            self.top_banner_path.set(path)
            self.log(f"Selected top banner: {os.path.basename(path)}", 'SUCCESS')
    
    def browse_bottom_banner(self):
        path = filedialog.askopenfilename(
            title="Select Bottom Banner Image",
            initialdir=SCRIPT_DIR,
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.webp"), ("All files", "*.*")]
        )
        if path:
            self.bottom_banner_path.set(path)
            self.log(f"Selected bottom banner: {os.path.basename(path)}", 'SUCCESS')
    
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
        self.log("Starting TRI-FOLD mailer generation...")
        
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
            
            # --- Load banner images ---
            top_banner_img = None
            bottom_banner_img = None
            
            if self.top_banner_path.get() and os.path.exists(self.top_banner_path.get()):
                top_banner_img = image_to_base64(self.top_banner_path.get())
                self.log(f"  Loaded top banner: {os.path.basename(self.top_banner_path.get())}", 'SUCCESS')
            else:
                self.log("  No top banner selected (will show placeholder)", 'WARNING')
            
            if self.bottom_banner_path.get() and os.path.exists(self.bottom_banner_path.get()):
                bottom_banner_img = image_to_base64(self.bottom_banner_path.get())
                self.log(f"  Loaded bottom banner: {os.path.basename(self.bottom_banner_path.get())}", 'SUCCESS')
            else:
                self.log("  No bottom banner selected (will show placeholder)", 'WARNING')
            
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
            
            self.log(f"  Loaded {total_clients} clients", 'SUCCESS')
            self.log(f"  Loaded {total_sold} sold properties", 'SUCCESS')
            
            # --- STEP 2: Geocoding with Mapbox ---
            self.update_status("🗺️ Geocoding addresses with Mapbox...")
            self.log("Geocoding addresses...")
            
            # Geocode clients
            client_coords = []
            for i, row in df_clients.iterrows():
                coords = self.get_coords_mapbox(row, len(client_coords), total_clients, "Client")
                client_coords.append(coords)
                self.update_progress(len(client_coords), total_clients + total_sold)
            df_clients['coords'] = client_coords
            
            # Geocode sold properties
            sold_coords = []
            for i, row in df_sold.iterrows():
                coords = self.get_coords_mapbox(row, len(sold_coords), total_sold, "Sold")
                sold_coords.append(coords)
                self.update_progress(total_clients + len(sold_coords), total_clients + total_sold)
            df_sold['coords'] = sold_coords
            
            valid_clients = df_clients.dropna(subset=['coords']).copy()
            valid_sold = df_sold.dropna(subset=['coords']).copy()
            
            self.log(f"Valid addresses: {len(valid_clients)} clients, {len(valid_sold)} sold", 'SUCCESS')
            
            # --- STEP 3: Generate PDFs ---
            self.update_status("📄 Generating TRI-FOLD mailers...")
            self.update_progress(0)
            self.log("Generating PDF mailers (8.5\" x 11\" tri-fold format)...")
            
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
                
                # Get client info
                first_name = str(client.get('Primary First', '')).strip()
                last_name = str(client.get('Primary Last', '')).strip()
                if not first_name or first_name.lower() == 'nan':
                    first_name = 'Neighbor'
                else:
                    first_name = first_name.upper()
                if last_name.lower() == 'nan':
                    last_name = ''
                else:
                    last_name = last_name.upper()
                
                address = str(client.get('Address', '')).strip()
                city = str(client.get('City', 'BAKERSFIELD')).strip().upper()
                zip_code = str(client.get('ZIP', '')).split('.')[0].strip()
                
                html_out = template.render(
                    first_name=first_name,
                    last_name=last_name,
                    address=address,
                    city=city,
                    zip_code=zip_code,
                    nearby=nearby,
                    map_url=map_url,
                    top_banner_img=top_banner_img,
                    bottom_banner_img=bottom_banner_img
                )
                
                # Save map image for verification
                try:
                    img_response = requests.get(map_url, timeout=15)
                    if img_response.status_code == 200:
                        safe_name = f"{first_name}_{last_name}".replace(" ", "_")
                        map_path = os.path.join(MAP_DEBUG_DIR, f"map_{safe_name}.png")
                        with open(map_path, "wb") as f:
                            f.write(img_response.content)
                except Exception as e:
                    pass  # Non-critical, continue
                
                # Generate PDF
                file_path = os.path.join(INDIVIDUAL_DIR, f"mailer_trifold_{idx}.pdf")
                HTML(string=html_out).write_pdf(file_path)
                pdf_files.append(file_path)
                
                self.log(f"  [{idx+1}/{len(valid_clients)}] {first_name} {last_name} @ {address[:30]}...")
                self.update_detail(f"Created mailer {len(pdf_files)}/{len(valid_clients)}")
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
            self.log(f"COMPLETE: Generated {len(pdf_files)} tri-fold mailers!", 'SUCCESS')
            
            stats_text = f"✅ Generated {len(pdf_files)} tri-fold mailers (8.5\" x 11\")\n"
            stats_text += f"📄 Output: {os.path.join(OUTPUT_DIR, FINAL_PDF)}\n"
            if self.skipped_log:
                stats_text += f"⚠️ Skipped {len(self.skipped_log)} addresses"
            
            self.stats_label.config(text=stats_text)
            
            messagebox.showinfo("Success!", f"Successfully generated {len(pdf_files)} tri-fold mailers!\n\nOutput saved to:\n{OUTPUT_DIR}")
            
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
