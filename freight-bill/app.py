import uuid
import shutil
import hashlib
import threading
from datetime import date
from pathlib import Path

import pandas as pd
from flask import (Flask, abort, render_template, request, jsonify,
                   send_file, session, after_this_request)

from config import (APPLICATION_ROOT, SECRET_KEY,
                    UPLOAD_DIR, OUTPUT_DIR,
                    CARRIER_BILLS_DIR, CARRIER_BILLS_URL,
                    CARRIER_SOURCE_DIR, CARRIER_AUDIT_DIR,
                    get_acumatica_config)
from processor import (analyze_carrier_file, classify_file,
                       process_freight_bills,
                       fetch_acumatica_shipments,
                       validate_carrier_file,
                       validate_priority1_file,
                       validate_pacejet_file,
                       validate_wwex_raw_file,
                       append_wwex_raw_to_master,
                       WWEX_SMALL_PARCEL_ENABLED)

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.config['APPLICATION_ROOT'] = APPLICATION_ROOT
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB
app.config['TEMPLATES_AUTO_RELOAD'] = True

# ==============================================================================
# In-memory stores
# ==============================================================================
_jobs:   dict = {}  # {sid: {status, thread, data, error}}
_hashes: dict = {}  # {sid: {file_type: md5_hash}}
_lock = threading.Lock()
_master_lock = threading.Lock()  # serializes writes to the shared master Carrier Import File


def _get_sid() -> str | None:
    """Read sid from form data, query string, or cookie — works in iframes too."""
    return (request.form.get('sid') or
            request.args.get('sid') or
            session.get('sid'))


def _find_master_carrier_file():
    """Locate the live 'Carrier Import File*.xlsx' in the Iris shared folder."""
    if not CARRIER_SOURCE_DIR.exists():
        return None
    xlsx_files = [f for f in CARRIER_SOURCE_DIR.glob('*.xlsx')
                  if f.name.startswith('Carrier Import File') and not f.name.startswith('~$')]
    return xlsx_files[0] if xlsx_files else None

@app.context_processor
def inject_base():
    return {'base_url': APPLICATION_ROOT.rstrip('/'), 'carrier_bills_url': CARRIER_BILLS_URL}

# ==============================================================================
# ROUTES
# ==============================================================================

@app.route('/')
def index():
    sid = str(uuid.uuid4())
    session['sid'] = sid  # fallback when cookies work
    return render_template('index.html', sid=sid)


@app.route('/upload-auto', methods=['POST'])
def upload_auto():
    """Single endpoint that accepts any file and auto-classifies it."""
    sid = _get_sid()
    if not sid:
        return jsonify(error='Session expired — refresh the page.'), 400

    f = request.files.get('file')
    if not f:
        return jsonify(error='No file received.'), 400

    filename    = f.filename or 'upload'
    filename_lc = filename.lower()
    is_csv      = filename_lc.endswith('.csv')
    # Preserve the real extension for legacy binary .xls (WWEX raw export) —
    # pandas picks its Excel engine (xlrd vs openpyxl) from the file suffix,
    # so saving old-format .xls content under a .xlsx name breaks reading it.
    ext = 'csv' if is_csv else ('xls' if filename_lc.endswith('.xls') else 'xlsx')
    upload_dir = UPLOAD_DIR / sid
    upload_dir.mkdir(parents=True, exist_ok=True)

    temp_path = upload_dir / f'_temp.{ext}'
    f.save(str(temp_path))

    # Compute MD5 to detect duplicate uploads
    with open(str(temp_path), 'rb') as _fh:
        file_hash = hashlib.md5(_fh.read()).hexdigest()

    stored_hashes: dict = _hashes.get(sid, {})

    file_type, err = classify_file(temp_path, filename)
    if file_type == 'unknown':
        temp_path.unlink(missing_ok=True)
        return jsonify(error=err or 'Unrecognized file type.'), 400

    def move(src, dst):
        shutil.move(str(src), str(dst))

    # --- Duplicate detection ---
    TYPE_LABELS = {'carrier': 'Carrier Import', 'priority1': 'Priority 1', 'pacejet': 'PaceJet'}
    label = TYPE_LABELS.get(file_type, file_type)
    if file_type in stored_hashes:
        if stored_hashes[file_type] == file_hash:
            temp_path.unlink(missing_ok=True)
            return jsonify(error=f'Duplicate file — this {label} file has already been uploaded this session.'), 400
        else:
            temp_path.unlink(missing_ok=True)
            return jsonify(error=f'{label} already uploaded — a different {label} file was already processed. Refresh to start a new session.'), 400

    if file_type == 'carrier':
        is_valid, carriers, validation_msg = validate_carrier_file(temp_path)
        if not is_valid:
            temp_path.unlink(missing_ok=True)
            return jsonify(error=validation_msg), 400
        carrier_path = upload_dir / 'carrier.xlsx'
        move(temp_path, carrier_path)
        stored_hashes['carrier'] = file_hash
        _hashes[sid] = stored_hashes
        # Audit snapshot (same as auto-load)
        try:
            stamp = date.today().strftime('%m.%d.%y')
            CARRIER_AUDIT_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy(str(carrier_path), str(CARRIER_AUDIT_DIR / f'Carrier Import File {stamp}.xlsx'))
        except Exception:
            pass
        analysis       = analyze_carrier_file(carrier_path)
        # Required cells left blank — stop here rather than after processing.
        # Drop the file and its hash so the corrected one can be re-uploaded.
        if analysis['missing_message']:
            carrier_path.unlink(missing_ok=True)
            stored_hashes.pop('carrier', None)
            return jsonify(error=analysis['missing_message'], blocking=True), 400
        needs_priority1 = analysis['needs_priority1']
        summary         = analysis['summary']
        _start_acumatica_fetch(sid)
        return jsonify(type='carrier', needs_priority1=needs_priority1,
                       warning=validation_msg, carriers=carriers, summary=summary)

    elif file_type == 'priority1':
        is_valid, validation_msg = validate_priority1_file(temp_path, is_csv)
        if not is_valid:
            temp_path.unlink(missing_ok=True)
            return jsonify(error=validation_msg), 400
        move(temp_path, upload_dir / f'priority1.{ext}')
        stored_hashes['priority1'] = file_hash
        _hashes[sid] = stored_hashes
        _start_acumatica_fetch(sid)
        return jsonify(type='priority1')

    elif file_type == 'pacejet':
        is_valid, validation_msg = validate_pacejet_file(temp_path, is_csv)
        if not is_valid:
            temp_path.unlink(missing_ok=True)
            return jsonify(error=validation_msg), 400
        move(temp_path, upload_dir / f'pacejet.{ext}')
        stored_hashes['pacejet'] = file_hash
        _hashes[sid] = stored_hashes
        _start_acumatica_fetch(sid)
        return jsonify(type='pacejet')

    elif file_type == 'wwex_raw':
        is_valid, validation_msg = validate_wwex_raw_file(temp_path)
        if not is_valid:
            temp_path.unlink(missing_ok=True)
            return jsonify(error=validation_msg), 400

        # Detection only — stage the file and return immediately so the pill
        # can show up + start its loading spinner. The actual (slow) append
        # to the shared master workbook happens in /wwex-append below.
        staged_path = upload_dir / f'wwex_raw.{ext}'
        move(temp_path, staged_path)
        return jsonify(type='wwex_raw')


@app.route('/wwex-append', methods=['POST'])
def wwex_append():
    """Performs the actual (slow) append of a staged WWEX raw file into the
    shared master workbook. Split from /upload-auto so the front-end can show
    the WWEX pill's loading spinner for the duration of this call."""
    if not WWEX_SMALL_PARCEL_ENABLED:
        abort(404)

    sid = _get_sid()
    if not sid:
        return jsonify(error='Session expired — refresh the page.'), 400

    upload_dir = UPLOAD_DIR / sid
    staged_path = next(iter(upload_dir.glob('wwex_raw.*')), None)
    if not staged_path:
        return jsonify(error='No WWEX file staged for this session.'), 400

    master_path = _find_master_carrier_file()
    if not master_path:
        staged_path.unlink(missing_ok=True)
        return jsonify(error='Carrier Import File not found in the Iris shared folder.'), 400

    try:
        with _master_lock:
            stats = append_wwex_raw_to_master(staged_path, master_path)
    except Exception as e:
        return jsonify(error=str(e)), 400
    finally:
        staged_path.unlink(missing_ok=True)

    return jsonify(**stats)


@app.route('/upload-carrier', methods=['POST'])
def upload_carrier():
    sid = _get_sid()
    if not sid:
        return jsonify(error='Session expired — refresh the page.'), 400

    f = request.files.get('file')
    if not f:
        return jsonify(error='No file received.'), 400

    upload_dir = UPLOAD_DIR / sid
    upload_dir.mkdir(parents=True, exist_ok=True)
    carrier_path = upload_dir / 'carrier.xlsx'
    f.save(str(carrier_path))

    is_valid, carriers, validation_msg = validate_carrier_file(carrier_path)
    if not is_valid:
        carrier_path.unlink(missing_ok=True)
        return jsonify(error=validation_msg), 400

    try:
        analysis = analyze_carrier_file(carrier_path)
    except Exception as e:
        return jsonify(error=f'Error reading file: {e}'), 400

    _start_acumatica_fetch(sid)

    return jsonify(needs_priority1=analysis['needs_priority1'], warning=validation_msg,
                   carriers=carriers)


@app.route('/upload-priority1', methods=['POST'])
def upload_priority1():
    sid = _get_sid()
    if not sid:
        return jsonify(error='Session expired — refresh the page.'), 400

    f = request.files.get('file')
    if not f:
        return jsonify(error='No file received.'), 400

    upload_dir = UPLOAD_DIR / sid
    upload_dir.mkdir(parents=True, exist_ok=True)

    filename = (f.filename or '').lower()
    is_csv   = filename.endswith('.csv')
    ext      = 'csv' if is_csv else 'xlsx'
    p1_path  = upload_dir / f'priority1.{ext}'
    f.save(str(p1_path))

    is_valid, validation_msg = validate_priority1_file(p1_path, is_csv)
    if not is_valid:
        p1_path.unlink(missing_ok=True)
        return jsonify(error=validation_msg), 400

    return jsonify(ok=True)


@app.route('/upload-pacejet', methods=['POST'])
def upload_pacejet():
    sid = _get_sid()
    if not sid:
        return jsonify(error='Session expired — refresh the page.'), 400

    f = request.files.get('file')
    if not f:
        return jsonify(error='No file received.'), 400

    upload_dir = UPLOAD_DIR / sid
    upload_dir.mkdir(parents=True, exist_ok=True)

    filename = (f.filename or '').lower()
    is_csv   = filename.endswith('.csv')
    ext      = 'csv' if is_csv else 'xlsx'
    pj_path  = upload_dir / f'pacejet.{ext}'
    f.save(str(pj_path))

    is_valid, validation_msg = validate_pacejet_file(pj_path, is_csv)
    if not is_valid:
        pj_path.unlink(missing_ok=True)
        return jsonify(error=validation_msg), 400

    return jsonify(ok=True)


@app.route('/carrier-auto', methods=['POST'])
def carrier_auto():
    """Auto-load carrier file from Iris shared folder, copy audit snapshot."""
    sid = _get_sid()
    if not sid:
        return jsonify(ok=False), 400

    source = _find_master_carrier_file()
    if not source:
        return jsonify(found=False)

    # Dated audit snapshot (overwrites same-day copy if re-opened)
    stamp      = date.today().strftime('%m.%d.%y')
    audit_name = f'Carrier Import File {stamp}.xlsx'
    try:
        CARRIER_AUDIT_DIR.mkdir(parents=True, exist_ok=True)
        shutil.copy(str(source), str(CARRIER_AUDIT_DIR / audit_name))
    except Exception:
        pass

    # Copy into session upload dir
    upload_dir   = UPLOAD_DIR / sid
    upload_dir.mkdir(parents=True, exist_ok=True)
    carrier_path = upload_dir / 'carrier.xlsx'
    shutil.copy2(str(source), str(carrier_path))

    # Hash (prevents duplicate-upload false-positives on manual re-upload)
    with open(str(carrier_path), 'rb') as fh:
        file_hash = hashlib.md5(fh.read()).hexdigest()
    stored_hashes = _hashes.get(sid, {})
    stored_hashes['carrier'] = file_hash
    _hashes[sid] = stored_hashes

    is_valid, carriers, validation_msg = validate_carrier_file(carrier_path)
    if not is_valid:
        carrier_path.unlink(missing_ok=True)
        return jsonify(found=True, error=validation_msg), 400

    analysis = analyze_carrier_file(carrier_path)
    if analysis['missing_message']:
        carrier_path.unlink(missing_ok=True)
        stored_hashes.pop('carrier', None)
        _hashes[sid] = stored_hashes
        return jsonify(found=True, error=analysis['missing_message'], blocking=True), 400

    _start_acumatica_fetch(sid)

    return jsonify(
        found=True,
        needs_priority1=analysis['needs_priority1'],
        warning=validation_msg,
        carriers=carriers,
        summary=analysis['summary'],
    )


@app.route('/init', methods=['POST'])
def init():
    sid = _get_sid()
    if not sid:
        return jsonify(ok=False), 400
    _start_acumatica_fetch(sid)
    return jsonify(ok=True)


@app.route('/acumatica-status')
def acumatica_status():
    sid = _get_sid()
    if not sid:
        return jsonify(status='error', error='No session')
    with _lock:
        job = _jobs.get(sid)
    if not job:
        return jsonify(status='not_started')
    return jsonify(status=job['status'], error=job.get('error'))


@app.route('/process', methods=['POST'])
def process():
    sid = _get_sid()
    if not sid:
        return jsonify(error='Session expired — refresh the page.'), 400

    # Wait for Acumatica thread (up to 2 minutes)
    with _lock:
        job = _jobs.get(sid)

    if job and job['status'] == 'running':
        job['thread'].join(timeout=120)
        with _lock:
            job = _jobs.get(sid)

    if not job or job['status'] == 'error':
        err = job.get('error', 'unknown error') if job else 'Acumatica job not found'
        return jsonify(error=f'Acumatica sync failed: {err}'), 500

    shipments_df = job.get('data', pd.DataFrame())
    upload_dir   = UPLOAD_DIR / sid

    # Load Priority 1 file if uploaded
    priority1_df = pd.DataFrame()
    for ext in ('xlsx', 'csv'):
        p1_path = upload_dir / f'priority1.{ext}'
        if p1_path.exists():
            priority1_df = (pd.read_excel(str(p1_path))
                            if ext == 'xlsx' else pd.read_csv(str(p1_path), low_memory=False))
            break

    # Load PaceJet file if uploaded
    pacejet_df = pd.DataFrame()
    for ext in ('csv', 'xlsx'):
        pj_path = upload_dir / f'pacejet.{ext}'
        if pj_path.exists():
            pacejet_df = (pd.read_csv(str(pj_path), low_memory=False)
                          if ext == 'csv' else pd.read_excel(str(pj_path)))
            break

    carrier_path = upload_dir / 'carrier.xlsx'
    output_dir   = OUTPUT_DIR / sid
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        result = process_freight_bills(
            carrier_path, priority1_df, pacejet_df, shipments_df, output_dir
        )
    except Exception as e:
        return jsonify(error=str(e)), 500

    # Copy to shared OneDrive folder (syncs to SharePoint automatically)
    # shutil.copy (not copy2) avoids metadata errors when OneDrive is actively syncing
    output_path  = OUTPUT_DIR / sid / result['filename']
    copy_error   = None
    try:
        CARRIER_BILLS_DIR.mkdir(parents=True, exist_ok=True)
        shutil.copy(str(output_path), str(CARRIER_BILLS_DIR / result['filename']))
    except Exception as e:
        copy_error = str(e)

    shutil.rmtree(str(upload_dir), ignore_errors=True)
    with _lock:
        _jobs.pop(sid, None)
    _hashes.pop(sid, None)

    if not copy_error:
        shutil.rmtree(str(output_dir), ignore_errors=True)

    return jsonify({**result, 'copy_error': copy_error, 'sid': sid})

@app.route('/download')
def download():
    sid      = request.args.get('sid') or session.get('output_sid')
    filename = request.args.get('file') or session.get('output_file')
    if not sid or not filename:
        return 'File not found.', 404

    file_path = OUTPUT_DIR / sid / filename
    if not file_path.exists():
        return 'File not found.', 404

    output_dir = OUTPUT_DIR / sid

    @after_this_request
    def _cleanup(response):
        shutil.rmtree(str(output_dir), ignore_errors=True)
        return response

    return send_file(
        str(file_path),
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


# ==============================================================================
# INTERNAL: background Acumatica fetch
# ==============================================================================

def _start_acumatica_fetch(sid: str):
    with _lock:
        if sid in _jobs:
            return
        job = {'status': 'running', 'thread': None, 'data': None, 'error': None}
        _jobs[sid] = job

    def _fetch():
        try:
            cfg = get_acumatica_config()
            df  = fetch_acumatica_shipments(cfg['username'], cfg['password'], cfg['url'])
            with _lock:
                _jobs[sid]['data']   = df
                _jobs[sid]['status'] = 'done'
        except Exception as e:
            with _lock:
                _jobs[sid]['error']  = str(e)
                _jobs[sid]['status'] = 'error'

    t = threading.Thread(target=_fetch, daemon=True, name=f'acumatica-{sid[:8]}')
    with _lock:
        _jobs[sid]['thread'] = t
    t.start()
