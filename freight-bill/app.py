import io
import uuid
import shutil
import hashlib
import zipfile
import threading
import contextlib
import mimetypes
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from flask import (Flask, abort, render_template, request, jsonify,
                   send_file, session, after_this_request)

from config import (APPLICATION_ROOT, SECRET_KEY,
                    BASE_DIR, UPLOAD_DIR, OUTPUT_DIR,
                    CARRIER_BILLS_DIR, CARRIER_BILLS_URL,
                    CARRIER_SOURCE_DIR, CARRIER_SP_SOURCE_DIR, CARRIER_AUDIT_DIR,
                    get_acumatica_config)
from processor import (analyze_carrier_file, classify_file,
                       process_freight_bills,
                       probe_acumatica,
                       fetch_acumatica_shipments_matching,
                       freight_lookup_values,
                       small_parcel_lookup_values,
                       validate_carrier_file,
                       validate_priority1_file,
                       validate_pacejet_file,
                       validate_wwex_raw_file,
                       validate_fedex_pdf,
                       validate_small_parcel_master,
                       append_wwex_raw_to_master,
                       append_fedex_pdf_to_master,
                       build_small_parcel_import,
                       set_carrier_tab_status,
                       SMALL_PARCEL_ENABLED,
                       BlockingImportError,
                       STATUS_PENDING, STATUS_BILL_GENERATED)

# Under the IIS pool identity the registry lookup can map .css to text/plain,
# and browsers refuse a stylesheet served that way (seen in Lowe's Invoice
# Reconciler with .js). Pin the types the templates load.
mimetypes.add_type('text/css', '.css')
mimetypes.add_type('text/javascript', '.js')
mimetypes.add_type('font/woff2', '.woff2')

app = Flask(__name__)
app.secret_key = SECRET_KEY
app.config['APPLICATION_ROOT'] = APPLICATION_ROOT
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB
app.config['TEMPLATES_AUTO_RELOAD'] = True

# ==============================================================================
# In-memory stores
# ==============================================================================
_jobs:   dict = {}  # {sid: {status, thread, error, asked, frames, fetch_lock}}
_hashes: dict = {}  # {sid: {file_type: md5_hash}}
_lock = threading.Lock()
_master_lock = threading.Lock()  # serializes writes to the shared master Carrier Import File
_NULL_LOCK = contextlib.nullcontext()  # session copies are per-user, nothing to serialize

# Small-parcel uploads: file type -> (staged file stem, appender, target tab).
# Both tabs live in the same workbook, which is why _master_lock is required
# rather than merely nice to have.
PARCEL_FILE_TYPES = {
    'wwex_raw':  ('wwex_raw',  append_wwex_raw_to_master,  'WWEX'),
    'fedex_pdf': ('fedex_pdf', append_fedex_pdf_to_master, 'Fedex'),
}
# Name of the hand-uploaded workbook inside the session directory, used when the
# shared folder is unavailable.
SP_MASTER_COPY = 'Carrier Import File - Small Parcels.xlsx'
# Shown whenever the shared workbook is open somewhere. Appending into an open
# workbook loses the appended rows — see _excel_lock_files.
SP_MASTER_OPEN_ERROR = 'Close the Small Parcels workbook in Excel, then try again.'
# How long an import waits for the Acumatica fetch before giving up. The table
# is ~150k rows and takes about a minute from an idle server.
ACUMATICA_WAIT_SECONDS = 300
# Everything the Small Parcels screen accepts, and nothing the main screen does.
# The PaceJet export is shared with the main screen: there it resolves LTL
# invoices, here it names the shipment behind a tracking number Acumatica does
# not know, so it is allowed from either scope.
PARCEL_SCOPE_TYPES = set(PARCEL_FILE_TYPES) | {'sp_master'}
DUAL_SCOPE_TYPES   = {'pacejet'}


def _get_sid() -> str | None:
    """Read sid from form data, query string, or cookie — works in iframes too."""
    return (request.form.get('sid') or
            request.args.get('sid') or
            session.get('sid'))


def _excel_lock_files(path: Path):
    """Excel owner files ('~$<name>.xlsx') sitting next to a workbook — present
    for as long as anyone has it open.

    Worth checking before every write to a shared workbook: a OneDrive-synced
    file open with AutoSave on is NOT locked on disk, so openpyxl saves over it
    happily and Excel later resolves its own merge conflict by writing the copy
    it still holds in memory — silently discarding every row the app appended.
    A PermissionError never happens, so the owner file is the only warning.

    Excel drops leading characters from long names, so the tail of the owner
    file name is matched rather than the whole name."""
    return [f for f in path.parent.glob('~$*')
            if path.name.endswith(f.name[2:])]


def _find_master_carrier_file():
    """Locate the live 'Carrier Import File*.xlsx' in the Iris shared folder."""
    if not CARRIER_SOURCE_DIR.exists():
        return None
    xlsx_files = [f for f in CARRIER_SOURCE_DIR.glob('*.xlsx')
                  if f.name.startswith('Carrier Import File') and not f.name.startswith('~$')]
    return xlsx_files[0] if xlsx_files else None


def _find_master_small_parcel_file():
    """Locate the live small-parcel 'Carrier Import File*.xlsx' — a different
    workbook, in a different shared folder, than the LTL one above. The folder
    can hold more than one candidate during a rename, so the most recently
    modified file wins instead of whatever the filesystem lists first."""
    if not CARRIER_SP_SOURCE_DIR.exists():
        return None
    xlsx_files = [f for f in CARRIER_SP_SOURCE_DIR.glob('*.xlsx')
                  if f.name.startswith('Carrier Import File') and not f.name.startswith('~$')]
    if not xlsx_files:
        return None
    return max(xlsx_files, key=lambda f: f.stat().st_mtime)

def _asset_version() -> int:
    """Newest mtime of the page's CSS/JS, so an edit reaches browsers without a hard refresh."""
    static = Path(app.static_folder)
    files = ('css/swcorp-tokens.css', 'css/app.css', 'js/app.js')
    return max((int((static / f).stat().st_mtime) for f in files if (static / f).exists()), default=0)

@app.context_processor
def inject_base():
    return {'base_url': APPLICATION_ROOT.rstrip('/'), 'carrier_bills_url': CARRIER_BILLS_URL,
            'asset_v': _asset_version()}

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
    # PDFs (FedEx invoices) keep their own suffix for the same reason.
    if is_csv:
        ext = 'csv'
    elif filename_lc.endswith('.pdf'):
        ext = 'pdf'
    elif filename_lc.endswith('.xls'):
        ext = 'xls'
    else:
        ext = 'xlsx'
    # Which screen the drop came from. Small-parcel files are only accepted from
    # the Small Parcels screen, and that screen accepts nothing else — enforced
    # here rather than client-side only.
    scope      = request.form.get('scope', 'main')
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
        # Logged so a refused upload can be explained afterwards — the user only
        # ever sees the toast, and toasts do not survive the trip back.
        app.logger.warning('upload refused (scope=%s): %s — %s',
                           scope, filename, err or 'Unrecognized file type.')
        return jsonify(error=err or 'Unrecognized file type.'), 400

    if file_type in PARCEL_SCOPE_TYPES and scope != 'parcels':
        temp_path.unlink(missing_ok=True)
        return jsonify(error='Small-parcel file — open the Small Parcels '
                             'screen to import it.'), 400
    if (file_type not in PARCEL_SCOPE_TYPES and file_type not in DUAL_SCOPE_TYPES
            and scope == 'parcels'):
        temp_path.unlink(missing_ok=True)
        return jsonify(error='Wrong file — the Small Parcels screen only accepts '
                             'the WWEX raw export, a FedEx invoice PDF, the '
                             'PaceJet export, or the Small Parcels workbook.'), 400

    def move(src, dst):
        shutil.move(str(src), str(dst))

    # --- Duplicate detection ---
    TYPE_LABELS = {'carrier': 'Carrier Import', 'priority1': 'Priority 1',
                   'pacejet': 'PaceJet', 'wwex_raw': 'WWEX raw export',
                   'fedex_pdf': 'FedEx invoice', 'sp_master': 'Small Parcels workbook'}
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
        _prefetch_freight(sid)
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
        _prefetch_freight(sid)
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
        _prefetch_freight(sid)
        return jsonify(type='pacejet')

    elif file_type in ('wwex_raw', 'fedex_pdf'):
        validate = (validate_wwex_raw_file if file_type == 'wwex_raw'
                    else validate_fedex_pdf)
        is_valid, validation_msg = validate(temp_path)
        if not is_valid:
            temp_path.unlink(missing_ok=True)
            return jsonify(error=validation_msg), 400

        # Detection only — stage the file and return immediately so the pill
        # can show up + start its loading spinner. The actual (slow) append
        # to the master workbook happens in /parcels-append below.
        move(temp_path, upload_dir / f'{PARCEL_FILE_TYPES[file_type][0]}.{ext}')
        return jsonify(type=file_type)

    elif file_type == 'sp_master':
        # Manual fallback: the shared folder was unavailable, so the user
        # supplies the workbook and gets the updated copy back to save.
        is_valid, validation_msg = validate_small_parcel_master(temp_path)
        if not is_valid:
            temp_path.unlink(missing_ok=True)
            return jsonify(error=validation_msg), 400
        move(temp_path, upload_dir / SP_MASTER_COPY)
        return jsonify(type='sp_master', filename=filename)

    temp_path.unlink(missing_ok=True)
    return jsonify(error='Unrecognized file type.'), 400


# ==============================================================================
# SMALL PARCELS
# ==============================================================================

def _session_master_copy(sid):
    """The hand-uploaded Small Parcels workbook for this session, if any."""
    path = UPLOAD_DIR / sid / SP_MASTER_COPY
    return path if path.exists() else None


def _session_pacejet_df(sid):
    """The PaceJet export uploaded this session, if any. Optional everywhere —
    it only rescues rows the carrier billed without an Acumatica reference, so a
    missing or unreadable file must never stop an import."""
    upload_dir = UPLOAD_DIR / sid
    for ext in ('csv', 'xlsx', 'xls'):
        path = upload_dir / f'pacejet.{ext}'
        if path.exists():
            try:
                return (pd.read_csv(str(path), low_memory=False) if ext == 'csv'
                        else pd.read_excel(str(path)))
            except Exception:
                return None
    return None


# A OneDrive file that is synced but not downloaded is a placeholder: it is
# listed, it reports its real size, and reading it makes Windows ask the sync
# engine for the bytes. That request is served only inside the session OneDrive
# runs in, and this app runs as a service in session 0 — so a read of a
# not-yet-downloaded file fails there with WinError 395, "Access to the cloud
# file is denied", even though the same read works from the desktop. Enumerating
# the folder needs no bytes, which is why the file is always found first and
# only the copy fails.
_CLOUD_PLACEHOLDER_ATTRS = (
    0x00400000 |  # FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS
    0x00040000)   # FILE_ATTRIBUTE_RECALL_ON_OPEN
# WinError 395 = ERROR_CLOUD_FILE_ACCESS_DENIED. 362..399 is the cloud-filter
# error block; the neighbours (provider not running, sync root gone, hydration
# request timed out) all mean the same thing to the user.
_CLOUD_ERRNOS = set(range(362, 400))
CLOUD_FILE_ERROR = (
    'The {name} in the shared folder has not been downloaded to this server — '
    'OneDrive is keeping it online-only, and a server app cannot download it. '
    'In File Explorer, right-click the folder and choose "Always keep on this '
    'device", or upload the file manually below.')


def _is_cloud_placeholder(path):
    """True when a file is present in the folder listing but its contents are
    not on this disk."""
    try:
        return bool(path.stat().st_file_attributes & _CLOUD_PLACEHOLDER_ATTRS)
    except (AttributeError, OSError):
        # st_file_attributes is Windows-only, and a file that cannot be stat'ed
        # at all is a different failure — let the read report it.
        return False


def _is_cloud_error(exc):
    """True for the OSError family Windows raises when a file's contents live
    only in the cloud."""
    return isinstance(exc, OSError) and getattr(exc, 'winerror', None) in _CLOUD_ERRNOS


def _copy_from_shared_folder(source, dest, name):
    """Copies a file out of a OneDrive-synced folder into local session storage.
    Returns (ok, error) — the error is the message the user acts on, never a
    traceback: an undownloaded file used to raise straight out of the request,
    and the browser, expecting JSON, dropped the spinner and said nothing at
    all.

    The copy is attempted before anything is concluded from the file's
    attributes. A placeholder is only unreadable from a service session; run
    from a desktop session the same read downloads the file and succeeds, and
    refusing up front would break the case that works."""
    try:
        shutil.copy2(str(source), str(dest))
    except Exception as e:
        if _is_cloud_error(e) or _is_cloud_placeholder(Path(source)):
            app.logger.warning('%s could not be read from the cloud (%s): %s',
                               name, source, e)
            return False, CLOUD_FILE_ERROR.format(name=name)
        app.logger.exception('%s could not be copied (%s): %s', name, source, e)
        return False, f'Could not read the {name} from the shared folder — {e}'
    return True, None


def _keep_blocked_upload(staged_path, tab):
    """Copies a file that refused to import into logs/blocked/ and returns the
    copy's name.

    The staged upload is deleted as soon as /parcels-append returns, so an
    invoice the parser could not reconcile used to leave nothing behind to look
    at — the FedEx invoice that failed on a renamed charge class had to be
    hunted down from the billing portal before it could be diagnosed. Best
    effort: a blocked import is already the answer to the user, and failing to
    keep the evidence must not change what they are told."""
    try:
        keep_dir = BASE_DIR / 'logs' / 'blocked'
        keep_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime('%Y-%m-%d_%H%M%S')
        target = keep_dir / f'{stamp}_{tab}{staged_path.suffix.lower()}'
        shutil.copy2(str(staged_path), str(target))
        return target.name
    except Exception as e:
        app.logger.warning('could not keep blocked upload (%s): %s', staged_path.name, e)
        return ''


def _wait_for_acumatica(sid):
    """Returns an error message, or None once Acumatica has answered. Blocks on
    the connection check started when the page opened.

    Every Subaccount on the import file comes from Acumatica, so a dead
    connection stops the import before anything is written — a bill with 400
    blank Subaccounts is worse than no bill at all. The rows themselves are
    asked afterwards, for the values of the file (_acumatica_shipments)."""
    with _lock:
        job = _jobs.get(sid)
    if job and job['status'] == 'running':
        job['thread'].join(timeout=ACUMATICA_WAIT_SECONDS)
        with _lock:
            job = _jobs.get(sid)

    if not job:
        return 'Acumatica sync never started — refresh the page and try again.'
    if job['status'] == 'running':
        return (f'Acumatica did not answer within '
                f'{ACUMATICA_WAIT_SECONDS // 60} minutes. Nothing was logged — try again.')
    if job['status'] != 'done':
        return f'Acumatica sync failed: {job.get("error") or "unknown error"}'
    return None


def _acumatica_shipments(sid, field_values):
    """Shipments rows for these values (DataFrame), asking Acumatica only for
    the values this session has not asked yet; everything asked before is kept
    in the job, so a later call (Process after the prefetch, a second invoice)
    costs only what is new. Raises when Acumatica fails; the values of a failed
    call stay unasked, so calling again retries them."""
    with _lock:
        job = _jobs.get(sid)
    if not job:
        raise ValueError('Acumatica sync never started — refresh the page and try again.')
    with job['fetch_lock']:
        asked = job['asked']
        todo = {f: {v for v in vals if v.upper() not in asked.setdefault(f, set())}
                for f, vals in field_values.items()}
        if any(todo.values()):
            cfg = get_acumatica_config()
            df = fetch_acumatica_shipments_matching(cfg['username'], cfg['password'],
                                                    cfg['url'], todo)
            job['frames'].append(df)
            for f, vals in todo.items():
                asked[f].update(v.upper() for v in vals)
            app.logger.info('Acumatica lookup (sid=%s): %s values -> %s rows',
                            sid[:8], sum(len(v) for v in todo.values()), len(df))
        frames = list(job['frames'])
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True).drop_duplicates(ignore_index=True)


def _session_freight_inputs(sid):
    """(carrier_path, priority1_df, pacejet_df) staged for this session."""
    upload_dir   = UPLOAD_DIR / sid
    priority1_df = pd.DataFrame()
    for ext in ('xlsx', 'csv'):
        p1_path = upload_dir / f'priority1.{ext}'
        if p1_path.exists():
            priority1_df = (pd.read_excel(str(p1_path))
                            if ext == 'xlsx' else pd.read_csv(str(p1_path), low_memory=False))
            break
    pacejet_df = pd.DataFrame()
    for ext in ('csv', 'xlsx'):
        pj_path = upload_dir / f'pacejet.{ext}'
        if pj_path.exists():
            pacejet_df = (pd.read_csv(str(pj_path), low_memory=False)
                          if ext == 'csv' else pd.read_excel(str(pj_path)))
            break
    return upload_dir / 'carrier.xlsx', priority1_df, pacejet_df


def _prefetch_freight(sid):
    """Asks Acumatica, in the background, for the rows of the files in hand
    (as LIR does when a file is dropped), so Process finds them already there.
    Best effort: a failure here is retried by Process, which asks again."""
    def _run():
        try:
            if _wait_for_acumatica(sid):
                return
            _acumatica_shipments(sid, freight_lookup_values(*_session_freight_inputs(sid)))
        except Exception as e:
            app.logger.warning('Acumatica prefetch failed (sid=%s): %s', sid[:8], e)
    threading.Thread(target=_run, daemon=True, name=f'acu-prefetch-{sid[:8]}').start()


def _snapshot_small_parcel_master(master_path):
    """Dated backup of the shared workbook before it is modified. Best-effort:
    a failed snapshot must not block the import."""
    try:
        stamp = date.today().strftime('%m.%d.%y')
        CARRIER_AUDIT_DIR.mkdir(parents=True, exist_ok=True)
        shutil.copy(str(master_path),
                    str(CARRIER_AUDIT_DIR / f'Carrier Import File - Small Parcels {stamp}.xlsx'))
    except Exception:
        pass


@app.route('/parcels-master', methods=['POST'])
def parcels_master():
    """Locates the Small Parcels workbook in the shared folder and confirms it is
    readable and correctly shaped, before any file is accepted for import.
    A workbook already uploaded by hand this session wins."""
    if not SMALL_PARCEL_ENABLED:
        abort(404)

    sid = _get_sid()
    if not sid:
        return jsonify(error='Session expired — refresh the page.'), 400

    # Subaccounts on the Acumatica import file come from matching each shipment
    # against Acumatica, so the fetch is kicked off here — it runs while the user
    # is still picking a file.
    _start_acumatica_fetch(sid)

    session_copy = _session_master_copy(sid)
    if session_copy:
        return jsonify(found=True, filename=session_copy.name, session_copy=True)

    source = _find_master_small_parcel_file()
    if not source:
        return jsonify(found=False,
                       error='Small Parcels workbook not found — upload it manually.')

    # A synced-but-not-yet-downloaded OneDrive placeholder, or a workbook open in
    # Excel, both look present on disk. Reading it here surfaces that now rather
    # than halfway through an import.
    is_valid, validation_msg = validate_small_parcel_master(source)
    if not is_valid:
        # A read that failed on a file whose contents are not on this disk gets
        # named for what it is. Reported raw it reads as a bare WinError, which
        # sends the user looking for a corrupt workbook instead of a sync
        # setting.
        if _is_cloud_placeholder(source):
            app.logger.warning('Small Parcels workbook is an undownloaded '
                               'OneDrive placeholder (%s): %s', source, validation_msg)
            validation_msg = CLOUD_FILE_ERROR.format(name='Small Parcels workbook')
        return jsonify(found=False, error=validation_msg)

    # Surfaced here as well as at append time so the user closes the workbook
    # before picking a file rather than after waiting through a parse.
    if _excel_lock_files(source):
        return jsonify(found=False, error=SP_MASTER_OPEN_ERROR)

    return jsonify(found=True, filename=source.name, session_copy=False)


@app.route('/parcels-append', methods=['POST'])
def parcels_append():
    """Performs the actual (slow) append of a staged small-parcel file into the
    Small Parcels workbook. Split from /upload-auto so the front-end can show
    the carrier pill's loading spinner for the duration of this call."""
    if not SMALL_PARCEL_ENABLED:
        abort(404)

    sid = _get_sid()
    if not sid:
        return jsonify(error='Session expired — refresh the page.'), 400

    file_type = request.form.get('type', '')
    if file_type not in PARCEL_FILE_TYPES:
        app.logger.warning('parcels-append refused: unknown type %r', file_type)
        return jsonify(error='Unknown small-parcel file type.'), 400
    stem, appender, tab = PARCEL_FILE_TYPES[file_type]

    upload_dir  = UPLOAD_DIR / sid
    staged_path = next(iter(upload_dir.glob(f'{stem}.*')), None)
    if not staged_path:
        app.logger.warning('parcels-append refused: no %s file staged (sid=%s)', tab, sid)
        return jsonify(error=f'No {tab} file staged for this session.'), 400

    # Re-resolved rather than trusted from the earlier /parcels-master call —
    # minutes can pass and the shared folder can stop syncing in between.
    session_copy = _session_master_copy(sid)
    master_path  = session_copy or _find_master_small_parcel_file()
    if not master_path:
        staged_path.unlink(missing_ok=True)
        app.logger.warning('parcels-append refused: master workbook not found (%s)', tab)
        return jsonify(error='Small Parcels workbook not found — upload it manually.'), 400

    # Re-checked immediately before the write: the workbook can be opened during
    # the minutes between /parcels-master and here, and appending into an open
    # workbook silently loses every appended row. The staged file is kept so the
    # user can retry with one click once the workbook is closed.
    if not session_copy and _excel_lock_files(master_path):
        app.logger.warning('parcels-append refused: master workbook is open (%s)', master_path)
        return jsonify(error=SP_MASTER_OPEN_ERROR, blocking=True), 400

    # Before the workbook is touched: without Acumatica the bill would come out
    # with every Subaccount blank, and the rows would already be logged, so the
    # invoice could not simply be re-imported. Failing here leaves the workbook
    # untouched and the file re-uploadable.
    acu_error = _wait_for_acumatica(sid)
    if acu_error:
        staged_path.unlink(missing_ok=True)
        app.logger.warning('parcels-append refused: %s', acu_error)
        return jsonify(error=acu_error, blocking=True), 400

    try:
        if session_copy:
            stats = appender(staged_path, master_path)
        else:
            with _master_lock:
                _snapshot_small_parcel_master(master_path)
                stats = appender(staged_path, master_path)
    # Blocking = the user has to resolve something before this invoice can be
    # imported at all, so the browser holds the message until it is dismissed.
    except BlockingImportError as e:
        kept = _keep_blocked_upload(staged_path, tab)
        app.logger.warning('parcels-append blocked for %s (%s): %s%s',
                           tab, staged_path.name, e,
                           f' [kept as logs/blocked/{kept}]' if kept else '')
        return jsonify(error=str(e), blocking=True), 400
    except Exception as e:
        kept = _keep_blocked_upload(staged_path, tab)
        app.logger.exception('parcels-append failed for %s (%s): %s%s',
                             tab, staged_path.name, e,
                             f' [kept as logs/blocked/{kept}]' if kept else '')
        return jsonify(error=str(e)), 400
    finally:
        staged_path.unlink(missing_ok=True)

    written = stats.pop('_written', [])
    extra   = _build_parcel_import_file(sid, tab, stats.get('invoice_number', ''),
                                        written)

    # The rows stay 'Pending' unless their Acumatica import file was produced.
    if extra.get('import_file'):
        try:
            lock = _master_lock if not session_copy else _NULL_LOCK
            with lock:
                set_carrier_tab_status(master_path, tab,
                                       [w.get('_row') for w in written],
                                       STATUS_BILL_GENERATED)
            extra['status'] = STATUS_BILL_GENERATED
        except Exception as e:
            extra['status'] = STATUS_PENDING
            extra['status_error'] = (f'The Acumatica file was created but the log '
                                     f'rows are still {STATUS_PENDING}: {e}')
    elif written:
        extra['status'] = STATUS_PENDING

    return jsonify(type=file_type, tab=tab,
                   session_copy=bool(session_copy), **stats, **extra)


def _build_parcel_import_file(sid, tab, invoice_number, written):
    """Builds the Acumatica import file for the rows just appended. Returns the
    fields to merge into the /parcels-append response. A re-upload that appended
    nothing produces no file — there is nothing new to import.

    The caller has already checked the Acumatica connection; the shipments rows
    are asked here, only for the references of the rows just written."""
    if not written:
        return {}

    pacejet_df = _session_pacejet_df(sid)
    try:
        shipments_df = _acumatica_shipments(sid, small_parcel_lookup_values(written, pacejet_df))
    except Exception as e:
        return {'import_error': f'Rows were logged, but Acumatica did not answer the '
                                f'Subaccount lookup: {e}'}

    output_dir = OUTPUT_DIR / sid
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = date.today().strftime('%m-%d-%Y')
    name  = f'Acumatica Import - {tab}{f" {invoice_number}" if invoice_number else ""} {stamp}.xlsx'

    try:
        info = build_small_parcel_import(written, shipments_df, output_dir / name,
                                         pacejet_df=pacejet_df)
    except Exception as e:
        return {'import_error': f'Rows were logged, but the Acumatica file failed: {e}'}

    return {
        'import_file':        info['filename'],
        'import_rows':        info['rows'],
        'import_matched':     info['matched'],
        'import_unmatched':   info['unmatched'],
        'import_via_pacejet': info['via_pacejet'],
        'import_via_sender':  info['via_sender'],
        'import_total':       info['total'],
    }


@app.route('/parcels-import-download')
def parcels_import_download():
    """Hands back a generated Acumatica import file. Keeps the directory so both
    carriers can be imported in one session and re-downloaded."""
    if not SMALL_PARCEL_ENABLED:
        abort(404)

    sid  = _get_sid()
    name = request.args.get('file', '')
    # Guards against a crafted name escaping the session directory.
    if not sid or not name or Path(name).name != name:
        return 'File not found.', 404

    path = OUTPUT_DIR / sid / name
    if not path.exists():
        return 'File not found.', 404

    return send_file(
        str(path),
        as_attachment=True,
        download_name=name,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@app.route('/parcels-import-zip')
def parcels_import_zip():
    """Every Acumatica bill built this session, in one archive. Browsers block
    the second and later of several automatic downloads from the same page, so
    one archive is the only way to hand over more than one file in one click."""
    if not SMALL_PARCEL_ENABLED:
        abort(404)

    sid = _get_sid()
    if not sid:
        return 'File not found.', 404

    out_dir = OUTPUT_DIR / sid
    files   = sorted(out_dir.glob('*.xlsx')) if out_dir.exists() else []
    if not files:
        return 'File not found.', 404

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(str(f), arcname=f.name)
    buf.seek(0)

    stamp = date.today().strftime('%m-%d-%Y')
    return send_file(buf, as_attachment=True,
                     download_name=f'Acumatica Bills {stamp}.zip',
                     mimetype='application/zip')


@app.route('/parcels-download')
def parcels_download():
    """Hands back the hand-uploaded workbook after rows were appended to it.
    Unlike /download this keeps the session directory, so more invoices can be
    added to the same copy and downloaded again."""
    if not SMALL_PARCEL_ENABLED:
        abort(404)

    sid = _get_sid()
    master_path = _session_master_copy(sid) if sid else None
    if not master_path:
        return 'File not found.', 404

    return send_file(
        str(master_path),
        as_attachment=True,
        download_name=SP_MASTER_COPY,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


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
    ok, copy_error = _copy_from_shared_folder(source, carrier_path, 'Carrier Import File')
    if not ok:
        return jsonify(found=True, error=copy_error, blocking=True), 400

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
    _prefetch_freight(sid)

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


# Files the × on a pill can take back: Freight Bill inputs still waiting to be
# used. WWEX/FedEx invoices are not here — they are appended to the Small
# Parcels workbook the moment they land, and that cannot be undone.
REMOVABLE_FILE_TYPES = ('carrier', 'priority1', 'pacejet')


@app.route('/remove-file', methods=['POST'])
def remove_file():
    """Drops a file loaded this session so another can take its place: the
    staged copy and its duplicate-upload hash both go, or the replacement would
    be refused as "already uploaded"."""
    sid = _get_sid()
    try:
        sid = str(uuid.UUID(sid or ''))   # it names a folder: never trust it raw
    except ValueError:
        return jsonify(error='Session expired — refresh the page.'), 400
    kind = request.form.get('type', '')
    if kind not in REMOVABLE_FILE_TYPES:
        return jsonify(error='This file cannot be removed.'), 400

    upload_dir = UPLOAD_DIR / sid
    for ext in ('xlsx', 'xls', 'csv'):
        (upload_dir / f'{kind}.{ext}').unlink(missing_ok=True)
    stored_hashes = _hashes.get(sid, {})
    stored_hashes.pop(kind, None)
    _hashes[sid] = stored_hashes
    app.logger.info('file removed by user (sid=%s): %s', sid, kind)
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

    acu_error = _wait_for_acumatica(sid)
    if acu_error:
        return jsonify(error=acu_error), 500

    upload_dir = UPLOAD_DIR / sid
    carrier_path, priority1_df, pacejet_df = _session_freight_inputs(sid)

    # The prefetch asked most of this already; only what is new goes out.
    try:
        shipments_df = _acumatica_shipments(
            sid, freight_lookup_values(carrier_path, priority1_df, pacejet_df))
    except Exception as e:
        return jsonify(error=f'Acumatica sync failed: {e}'), 500

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
        # status is the connection: 'done' once Acumatica answered one row
        # (that is the pill's check). The rows are asked per value later.
        job = {'status': 'running', 'thread': None, 'error': None,
               'asked': {}, 'frames': [], 'fetch_lock': threading.Lock()}
        _jobs[sid] = job

    def _fetch():
        try:
            cfg = get_acumatica_config()
            probe_acumatica(cfg['username'], cfg['password'], cfg['url'])
            with _lock:
                _jobs[sid]['status'] = 'done'
        except Exception as e:
            with _lock:
                _jobs[sid]['error']  = str(e)
                _jobs[sid]['status'] = 'error'

    t = threading.Thread(target=_fetch, daemon=True, name=f'acumatica-{sid[:8]}')
    with _lock:
        _jobs[sid]['thread'] = t
    t.start()
