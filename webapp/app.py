"""
Real Estate Mailer Generator - Flask Web App
Asynchronous background processing with progress tracking
"""

import os
import sys
import shutil
import tempfile
import threading
import time
import uuid
import zipfile
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from dotenv import load_dotenv

# --- WINDOWS GTK PATCH (same as mailer_app_trifold.py) ---
MSYS2_BIN_PATH = r'D:\msys2\ucrt64\bin'
if sys.platform == 'win32' and os.path.exists(MSYS2_BIN_PATH):
    os.add_dll_directory(MSYS2_BIN_PATH)

from mailer_generator import generate_mailers

# Load environment variables
load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max upload
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-key-change-in-production')

# Allowed file extensions
ALLOWED_CSV = {'csv'}
ALLOWED_IMAGES = {'png', 'jpg', 'jpeg', 'gif', 'webp'}

# In-memory job store  { job_id: { status, progress, total, message, result, ... } }
jobs = {}

# Auto-cleanup: remove finished jobs older than 1 hour
JOB_TTL_SECONDS = 3600


def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions


def cleanup_old_jobs():
    """Remove jobs older than JOB_TTL_SECONDS"""
    now = time.time()
    expired = [jid for jid, j in jobs.items()
               if j.get('finished_at') and (now - j['finished_at']) > JOB_TTL_SECONDS]
    for jid in expired:
        job = jobs.pop(jid, None)
        if job and job.get('job_dir'):
            shutil.rmtree(job['job_dir'], ignore_errors=True)


def run_generation(job_id, job_dir, params):
    """Background worker that generates mailers and updates job progress"""
    job = jobs[job_id]

    def progress_callback(current, total, message):
        job['progress'] = current
        job['total'] = total
        job['message'] = message

    try:
        job['status'] = 'running'
        job['message'] = 'Starting generation...'

        output_dir = os.path.join(job_dir, 'output')
        os.makedirs(output_dir, exist_ok=True)

        result = generate_mailers(
            client_csv_path=params['client_csv_path'],
            sold_csv_path=params['sold_csv_path'],
            output_dir=output_dir,
            mapbox_token=params['mapbox_token'],
            top_banner_path=params.get('top_banner_path'),
            bottom_banner_path=params.get('bottom_banner_path'),
            right_side_image_path=params.get('right_side_path'),
            num_nearby=params.get('num_nearby', 3),
            num_clients=params.get('num_clients', 'all'),
            progress_callback=progress_callback
        )

        if not result['success']:
            job['status'] = 'failed'
            job['message'] = result.get('error', 'Generation failed')
            job['finished_at'] = time.time()
            return

        # Create ZIP
        job['message'] = 'Creating ZIP archive...'
        zip_path = os.path.join(job_dir, 'mailers.zip')
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for pdf_path in result['pdf_files']:
                arcname = os.path.join('individual', os.path.basename(pdf_path))
                zipf.write(pdf_path, arcname)
            if result['merged_pdf'] and os.path.exists(result['merged_pdf']):
                zipf.write(result['merged_pdf'], os.path.basename(result['merged_pdf']))

        job['status'] = 'done'
        job['zip_path'] = zip_path
        job['pdf_count'] = len(result['pdf_files'])
        job['skipped_count'] = len(result.get('skipped', []))
        job['message'] = f"Done! Generated {len(result['pdf_files'])} mailers."
        job['progress'] = job['total']
        job['finished_at'] = time.time()

    except Exception as e:
        import traceback
        job['status'] = 'failed'
        job['message'] = str(e)
        job['traceback'] = traceback.format_exc()
        job['finished_at'] = time.time()


# ─── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    """Render the main form"""
    default_mapbox_token = os.getenv('MAPBOX_TOKEN', '')
    return render_template('index.html', default_mapbox_token=default_mapbox_token)


@app.route('/generate', methods=['POST'])
def generate():
    """Accept uploads, start a background job, return job_id immediately"""
    cleanup_old_jobs()

    job_id = str(uuid.uuid4())
    job_dir = tempfile.mkdtemp(prefix='mailer_job_')
    uploads_dir = os.path.join(job_dir, 'uploads')
    os.makedirs(uploads_dir, exist_ok=True)

    try:
        # ── Validate uploads ──
        if 'client_csv' not in request.files or 'sold_csv' not in request.files:
            shutil.rmtree(job_dir, ignore_errors=True)
            return jsonify({'error': 'Both CSV files are required'}), 400

        client_csv = request.files['client_csv']
        sold_csv = request.files['sold_csv']

        if client_csv.filename == '' or sold_csv.filename == '':
            shutil.rmtree(job_dir, ignore_errors=True)
            return jsonify({'error': 'Both CSV files are required'}), 400

        if not allowed_file(client_csv.filename, ALLOWED_CSV):
            shutil.rmtree(job_dir, ignore_errors=True)
            return jsonify({'error': 'Client file must be a CSV'}), 400

        if not allowed_file(sold_csv.filename, ALLOWED_CSV):
            shutil.rmtree(job_dir, ignore_errors=True)
            return jsonify({'error': 'Sold homes file must be a CSV'}), 400

        # ── Save files ──
        client_csv_path = os.path.join(uploads_dir, secure_filename(client_csv.filename))
        sold_csv_path = os.path.join(uploads_dir, secure_filename(sold_csv.filename))
        client_csv.save(client_csv_path)
        sold_csv.save(sold_csv_path)

        top_banner_path = None
        bottom_banner_path = None
        right_side_path = None

        if 'top_banner' in request.files:
            f = request.files['top_banner']
            if f.filename != '' and allowed_file(f.filename, ALLOWED_IMAGES):
                top_banner_path = os.path.join(uploads_dir, 'top_' + secure_filename(f.filename))
                f.save(top_banner_path)

        if 'bottom_banner' in request.files:
            f = request.files['bottom_banner']
            if f.filename != '' and allowed_file(f.filename, ALLOWED_IMAGES):
                bottom_banner_path = os.path.join(uploads_dir, 'bottom_' + secure_filename(f.filename))
                f.save(bottom_banner_path)

        if 'right_side_image' in request.files:
            f = request.files['right_side_image']
            if f.filename != '' and allowed_file(f.filename, ALLOWED_IMAGES):
                right_side_path = os.path.join(uploads_dir, 'right_' + secure_filename(f.filename))
                f.save(right_side_path)

        mapbox_token = request.form.get('mapbox_token', os.getenv('MAPBOX_TOKEN', ''))
        num_nearby = int(request.form.get('num_nearby', 3))
        num_clients = request.form.get('num_clients', 'all')

        if not mapbox_token:
            shutil.rmtree(job_dir, ignore_errors=True)
            return jsonify({'error': 'Mapbox API token is required'}), 400

        # ── Create job record ──
        jobs[job_id] = {
            'status': 'queued',
            'progress': 0,
            'total': 1,
            'message': 'Job queued...',
            'job_dir': job_dir,
            'zip_path': None,
            'pdf_count': 0,
            'skipped_count': 0,
            'finished_at': None,
        }

        params = {
            'client_csv_path': client_csv_path,
            'sold_csv_path': sold_csv_path,
            'mapbox_token': mapbox_token,
            'top_banner_path': top_banner_path,
            'bottom_banner_path': bottom_banner_path,
            'right_side_path': right_side_path,
            'num_nearby': num_nearby,
            'num_clients': num_clients,
        }

        # ── Launch background thread ──
        thread = threading.Thread(target=run_generation, args=(job_id, job_dir, params), daemon=True)
        thread.start()

        return jsonify({'job_id': job_id}), 202

    except Exception as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        return jsonify({'error': str(e)}), 500


@app.route('/status/<job_id>')
def job_status(job_id):
    """Poll endpoint — returns current progress of a background job"""
    job = jobs.get(job_id)
    if not job:
        return jsonify({'error': 'Job not found'}), 404

    return jsonify({
        'status': job['status'],
        'progress': job['progress'],
        'total': job['total'],
        'message': job['message'],
        'pdf_count': job.get('pdf_count', 0),
        'skipped_count': job.get('skipped_count', 0),
    })


@app.route('/download/<job_id>')
def download(job_id):
    """Download the finished ZIP for a completed job"""
    job = jobs.get(job_id)
    if not job:
        return jsonify({'error': 'Job not found'}), 404
    if job['status'] != 'done':
        return jsonify({'error': 'Job not finished yet'}), 400
    if not job.get('zip_path') or not os.path.exists(job['zip_path']):
        return jsonify({'error': 'ZIP file not found'}), 404

    response = send_file(
        job['zip_path'],
        mimetype='application/zip',
        as_attachment=True,
        download_name='mailers.zip'
    )

    # Cleanup after download
    @response.call_on_close
    def cleanup():
        j = jobs.pop(job_id, None)
        if j and j.get('job_dir'):
            shutil.rmtree(j['job_dir'], ignore_errors=True)

    return response


if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    debug = os.getenv('FLASK_DEBUG', 'false').lower() == 'true'
    app.run(host='0.0.0.0', port=port, debug=debug)
