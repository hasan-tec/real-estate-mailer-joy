# Real Estate Mailer Generator - Web App

A Flask web application for generating real estate tri-fold mailers with personalized maps and nearby sales data.

## Features

- 📄 Generates 8.5" x 11" tri-fold mailer PDFs
- 🗺️ Mapbox integration for geocoding and static maps
- 📊 Shows nearby sold properties with prices
- 🎨 Custom banner image uploads
- 📦 Downloads all mailers as a ZIP file

## Quick Start (Local)

### 1. Install Dependencies

```bash
cd webapp
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Set Up Environment

```bash
cp .env.example .env
# Edit .env and add your Mapbox token
```

### 3. Run the App

```bash
python app.py
```

Open http://localhost:5000 in your browser.

## Deployment (Railway)

### 1. Push to GitHub

```bash
cd webapp
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/mailer-webapp.git
git push -u origin main
```

### 2. Deploy on Railway

1. Go to [railway.app](https://railway.app)
2. Click "New Project" → "Deploy from GitHub repo"
3. Select your repository
4. Add environment variables:
   - `MAPBOX_TOKEN` = your Mapbox token
   - `SECRET_KEY` = random secret string
5. Railway will auto-detect the Dockerfile and deploy

Your app will be live at `https://your-project.railway.app`

## File Structure

```
webapp/
├── app.py                 # Flask routes
├── mailer_generator.py    # PDF generation logic
├── templates/
│   └── index.html         # Form UI
├── static/
│   └── style.css          # Styling
├── requirements.txt       # Python dependencies
├── Dockerfile             # For deployment
└── .env.example           # Environment template
```

## CSV Format

### Client List CSV
Required columns:
- `Address`
- `City`
- `ZIP`
- `Primary First` (optional)
- `Primary Last` (optional)

### Sold Homes CSV
Required columns:
- `Address`
- `City`
- `ZIP`
- `Purchase Amt`
- `Beds`
- `Baths`
- `Sq Ft`

## Notes

- WeasyPrint requires GTK libraries (handled by Docker)
- Processing ~250 addresses takes 3-5 minutes
- Free Railway tier allows 500 hours/month
