import io
import re
import pandas as pd
import requests
from copy import copy
from requests.auth import HTTPBasicAuth
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def _to_bytes(file_path):
    """Read file into a BytesIO buffer, releasing the OS file handle immediately."""
    with open(str(file_path), 'rb') as f:
        return io.BytesIO(f.read())

# ==============================================================================
# BUSINESS CONSTANTS
# ==============================================================================

ACCOUNT = '410150'

CARRIER_VENDOR_IDS = {
    'AAA Cooper': 'V100087', 'ArcBest': 'V100532', 'Pilot': 'V100533',
    'Priority 1': 'V100559', 'R+L':     'V100468', 'RoadRunner': 'V100365',
    'Saia':       'V100471', 'Seko':    'V100472', 'WWEX': 'V100989',
}

# Master switch for the WWEX (Small Parcel) raw-file ingestion feature.
# Not needed yet — flip to True to re-enable without touching anything else.
WWEX_SMALL_PARCEL_ENABLED = False

NEGATIVE_AMOUNT_CONFIG = {
    'WWEX': {
        'account': '320180',
        'subaccount': 'DIS-000000-000000-0000000000',
        'doc_description': 'Freight Claims Revenue',
    }
}

# Small-parcel fee lines auto-inserted from the raw WWEX file (see
# append_wwex_raw_to_master) always post to this GL/Subaccount instead of the
# normal customer lookup — these are carrier-generated fees, not shipments.
WWEX_FEE_CONFIG = {
    'account':     '410160',
    'subaccount':  'DIS-000000-000000-0000000000',
}
WWEX_FEE_DESCRIPTIONS = {'WEEKLY SERVICE CHARGE', 'INVOICE PROCESSING FEE'}
WWEX_FEE_NOTE_PREFIX = 'AUTO-FEE:'

SUBACCOUNTS = {
    'Default':                       'SAL-DEALER-000000-0000000000',
    'C101408':                       'SAL-ACEHW -000000-0000000000',
    'C100026':                       'CAT-AMZCAN-000000-0000000000',
    'C100004':                       'CAT-AMZFBA-000000-0000000000',
    'C100007':                       'CAT-AMZFBM-000000-0000000000',
    'C100171':                       'SAL-ANZDIR-000000-0000000000',
    'C100125':                       'SAL-ANZZI -000000-0000000000',
    'Big Box Customers':             'SAL-DEALER-000000-0000000000',
    'C100003':                       'ACM-BUILD -000000-0000000000',
    'C100106':                       'ACM-COSTCO-000000-0000000000',
    'Direct to Consumer':            'SAL-DEALER-000000-0000000000',
    'Dealers (Dot Com)':             'SAL-DEALER-000000-0000000000',
    'Dealers (Brick and mortar)':    'SAL-DEALER-000000-0000000000',
    'Direct Sales':                  'SAL-DEALER-000000-0000000000',
    'C100120':                       'ACM-EBAY  -000000-0000000000',
    'C100126':                       'SAL-ENDTUB-000000-0000000000',
    'C100020':                       'SAL-FERGUS-000000-0000000000',
    'C100013':                       'ACM-HD SOS-000000-0000000000',
    'C100001':                       'ACM-HD.COM-000000-0000000000',
    'C100107':                       'ACM-HDCA  -000000-0000000000',
    'C101662':                       'ACM-HOMDFC-000000-0000000000',
    'C100124':                       'SAL-HOMMRT-000000-0000000000',
    'C100119':                       'ACM-HOUZZ -000000-0000000000',
    'C100002':                       'ACM-LOWES -000000-0000000000',
    'Main Customers':                'SAL-DEALER-000000-0000000000',
    'C100006':                       'ACM-MENARD-000000-0000000000',
    'C100009':                       'ACM-OVSTCK-000000-0000000000',
    'C100122':                       'SAL-SSDCOM-000000-0000000000',
    'C100172':                       'SAL-SSDDIR-000000-0000000000',
    'C100008':                       'SAL-SWP   -000000-0000000000',
    'C101413':                       'SAL-SWS   -000000-0000000000',
    'C100123':                       'SAL-UNITUB-000000-0000000000',
    'C100014':                       'ACM-WALMRT-000000-0000000000',
    'C100118':                       'SAL-WARRAN-000000-000000',
    'C100005':                       'ACM-WAYFAI-000000-0000000000',
    'C100015':                       'ACM-ZORO  -STGNTR-0000000000',
}
DEFAULT_SUBACCOUNT = 'SAL-DEALER-000000-0000000000'

# ==============================================================================
# HELPERS
# ==============================================================================

def safe_str(val):
    try:
        if pd.isna(val):
            return ''
    except (TypeError, ValueError):
        pass
    s = str(val).strip()
    if s.upper() in ('-', 'N/A', 'NAN', 'NONE', '0'):
        return ''
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return s

def safe_date(val):
    if isinstance(val, pd.Timestamp) and pd.notna(val):
        return val
    try:
        return pd.Timestamp(val)
    except Exception:
        return pd.Timestamp.now()

def is_missing_date(val):
    """True when a date cell is blank or unparseable. pd.Timestamp(None/NaN)
    returns NaT rather than raising, so a plain try/except never catches it."""
    try:
        return pd.isna(pd.Timestamp(val))
    except Exception:
        return True

def safe_amount(val):
    try:
        if pd.isna(val):
            return 0.0
    except (TypeError, ValueError):
        pass
    try:
        return round(float(str(val).replace('$', '').replace(',', '').strip()), 2)
    except Exception:
        return 0.0

def safe_date_only(val):
    if val is None:
        return None
    try:
        if isinstance(val, float) and pd.isna(val):
            return None
    except Exception:
        pass
    try:
        return pd.Timestamp(val).date()
    except Exception:
        return None

def find_col(df, keywords):
    for col in df.columns:
        if any(kw.lower() in str(col).lower() for kw in keywords):
            return col
    return None

def normalize_carrier_name(carrier_raw):
    """Strip suffix like '(ABFS)' for fuzzy carrier matching."""
    return re.sub(r'\s*\([^)]+\)\s*$', '', str(carrier_raw)).strip().lower()

# ==============================================================================
# LOOKUP BUILDERS
# ==============================================================================

def build_shipments_lookups(shipments_df):
    """Returns (ordernbr_lookup, shipmentnbr_lookup, tracking_lookup).
    Each maps key -> {Customer, CustomerName, OrderNbr}."""
    ordernbr_lookup    = {}
    shipmentnbr_lookup = {}
    tracking_lookup    = {}

    if shipments_df.empty:
        return ordernbr_lookup, shipmentnbr_lookup, tracking_lookup

    col_order    = find_col(shipments_df, ['ordernbr', 'order nbr'])
    col_shipment = find_col(shipments_df, ['shipmentnbr', 'shipment nbr'])
    col_tracking = find_col(shipments_df, ['trackingnumber', 'tracking number'])
    col_custname = find_col(shipments_df, ['customername', 'customer name'])
    col_customer = next(
        (col for col in shipments_df.columns if str(col).strip().lower() == 'customer'),
        None
    )

    # Use safe_str for key columns — critical: handles float→int (12345.0 → '12345')
    # so keys match what safe_str produces when extracting values from other files.
    def _key_col(col):
        if not col: return [''] * len(shipments_df)
        return shipments_df[col].apply(safe_str).tolist()

    def _str_col(col):
        if not col: return [''] * len(shipments_df)
        return shipments_df[col].fillna('').astype(str).str.strip().tolist()

    orders    = _key_col(col_order)
    shipments = _key_col(col_shipment)
    trackings = _key_col(col_tracking)
    customers = _str_col(col_customer)
    custnames = _str_col(col_custname)

    for o, s, t, c, n in zip(orders, shipments, trackings, customers, custnames):
        cust_data = {'Customer': c, 'CustomerName': n, 'OrderNbr': o}
        if o and o not in ordernbr_lookup:    ordernbr_lookup[o]    = cust_data
        if s and s not in shipmentnbr_lookup: shipmentnbr_lookup[s] = cust_data
        if t and t not in tracking_lookup:    tracking_lookup[t]    = cust_data

    return ordernbr_lookup, shipmentnbr_lookup, tracking_lookup


def build_p1_lookup(p1_df):
    """Maps Invoice Number -> {PRO, BOL, SO, Carrier, CarrierMode, RemainingBalance, ActualShip}."""
    if p1_df.empty:
        return {}

    col_inv     = find_col(p1_df, ['invoice number', 'invoice #', 'invoice#'])
    col_pro     = find_col(p1_df, ['pro number', 'pro #', 'pro'])
    col_bol     = find_col(p1_df, ['bol'])
    col_carrier = find_col(p1_df, ['carrier'])
    col_mode    = find_col(p1_df, ['carrier mode'])
    col_rembal  = find_col(p1_df, ['remaining balance'])
    col_actship = find_col(p1_df, ['actual ship'])
    # Exact match for "SO" — avoids false hits like "SW Corp BOL"
    col_so = next((col for col in p1_df.columns if str(col).strip().upper() == 'SO'), None)

    if not col_inv:
        return {}

    # Key columns must use safe_str to handle float→int (12345.0 → '12345')
    def _key(col):
        if not col: return [''] * len(p1_df)
        return p1_df[col].apply(safe_str).tolist()

    def _s(col):
        if not col: return [''] * len(p1_df)
        return p1_df[col].fillna('').astype(str).str.strip().tolist()

    invs      = _key(col_inv)
    pros      = _key(col_pro)
    bols      = _key(col_bol)
    sos       = _key(col_so)
    carriers  = _s(col_carrier)
    modes     = _s(col_mode)
    rembals   = (pd.to_numeric(
                    p1_df[col_rembal].astype(str)
                        .str.replace('$', '', regex=False)
                        .str.replace(',', '', regex=False)
                        .str.strip(),
                    errors='coerce').fillna(0.0).round(2).tolist()
                 if col_rembal else [0.0] * len(p1_df))
    actships  = (p1_df[col_actship].apply(safe_date_only).tolist()
                 if col_actship else [None] * len(p1_df))

    lookup = {}
    for inv, pro, bol, so, car, mode, rembal, actship in zip(
            invs, pros, bols, sos, carriers, modes, rembals, actships):
        if not inv:  # safe_str already blanks noise values
            continue
        lookup[inv] = {
            'PRO': pro, 'BOL': bol, 'SO': so,
            'Carrier': car, 'CarrierMode': mode,
            'RemainingBalance': rembal, 'ActualShip': actship,
        }
    return lookup


def build_pacejet_records(pacejet_df):
    """Returns (records_list, amount_index).
    amount_index maps each amount -> [list of record indices] for O(1) first-pass filter."""
    if pacejet_df.empty:
        return [], {}

    col_carr    = find_col(pacejet_df, ['carrierclassofservice'])
    col_carrnum = find_col(pacejet_df, ['carriernumber'])
    col_shipdt  = find_col(pacejet_df, ['shipshipdatetime', 'shipdatetime', 'shipmentshipdatetime'])
    col_uf3     = find_col(pacejet_df, ['shipmentuserfield3'])
    col_cf      = find_col(pacejet_df, ['consigneefreight'])
    col_nf      = find_col(pacejet_df, ['consignorfreight'])
    col_lf      = find_col(pacejet_df, ['listfreight'])
    col_fs      = find_col(pacejet_df, ['fuelsurcharge'])

    def _amounts_series(col):
        if not col: return pd.Series([0.0] * len(pacejet_df))
        return pd.to_numeric(
            pacejet_df[col].astype(str).str.replace('$', '', regex=False)
                           .str.replace(',', '', regex=False).str.strip(),
            errors='coerce').fillna(0.0).round(2)

    def _str_col(col):
        if not col: return [''] * len(pacejet_df)
        return pacejet_df[col].fillna('').astype(str).str.strip().tolist()

    cf_vals = _amounts_series(col_cf).tolist()
    nf_vals = _amounts_series(col_nf).tolist()
    lf_vals = _amounts_series(col_lf).tolist()
    fs_vals = _amounts_series(col_fs).tolist()
    carr_vals   = _str_col(col_carr)
    carrnum_vals = _str_col(col_carrnum)
    shipdt_vals = (pacejet_df[col_shipdt].apply(safe_date_only).tolist()
                   if col_shipdt else [None] * len(pacejet_df))
    uf3_vals    = _str_col(col_uf3)

    records     = []
    amount_index = {}  # amount -> [idx, ...]

    for i, (cf, nf, lf, fs, carr, carrnum, shipdt, uf3) in enumerate(zip(
            cf_vals, nf_vals, lf_vals, fs_vals,
            carr_vals, carrnum_vals, shipdt_vals, uf3_vals)):
        amounts = {v for v in (cf, nf, lf, fs) if v > 0}
        rec = {
            'amounts':        amounts,
            'carrier_code':   carr.lower(),
            'carrier_number': carrnum.lower(),
            'ship_date':      shipdt,
            'user_field3':    uf3,
        }
        records.append(rec)
        for amt in amounts:
            amount_index.setdefault(amt, []).append(i)

    return records, amount_index


def find_pacejet_match(amount, carrier_raw, carrier_mode, actual_ship_date,
                       pacejet_records, amount_index):
    """Returns user_field3 using the amount_index for O(1) candidate filtering."""
    candidates = amount_index.get(amount)
    if not candidates:
        return None
    carrier_name = normalize_carrier_name(carrier_raw)
    mode = carrier_mode.strip().lower() if carrier_mode else ''
    for idx in candidates:
        rec = pacejet_records[idx]
        if carrier_name and carrier_name not in rec['carrier_code']:
            continue
        if mode and mode not in rec['carrier_number']:
            continue
        if actual_ship_date and rec['ship_date'] != actual_ship_date:
            continue
        return rec['user_field3']
    return None

# ==============================================================================
# EXCEL OUTPUT HELPERS
# ==============================================================================

def _build_debit_adj_row(carrier_name, counter, doc_date, vendor_id,
                          pronumber, line_desc, amount, now):
    cfg = NEGATIVE_AMOUNT_CONFIG[carrier_name]
    return {
        'Type': 'Debit Adj.',
        'Reference Nbr.': counter,
        'Document Date': doc_date.strftime('%m-%d-%Y'),
        'Post Period': doc_date.strftime('%m-%Y'),
        'Document Description': cfg['doc_description'],
        'Vendor ID': vendor_id,
        'Carrier': carrier_name,
        'Location': '',
        'Vendor Ref.': pronumber,
        'Branch': 'MAIN',
        'Line Branch': 'MAIN',
        'Line Description': line_desc,
        'Quantity': 1,
        'Unit Cost': abs(amount),
        'Amount': abs(amount),
        'Account': cfg['account'],
        'Sub': '',
        'Subaccount': cfg['subaccount'],
        'Ap received Bills date ': now,
        'Invoice Number': pronumber,
        'Case': '',
        '_anomalous': False,
    }


def _write_formatted_excel(output_path, export_df, anomalous_flags, debit_positions):
    SP_HEADERS = {'Sub', 'Invoice Number', 'Carrier', 'Case'}
    hdr_fill    = PatternFill(fill_type='solid', fgColor='00375C')
    hdr_fill_sp = PatternFill(fill_type='solid', fgColor='0062A4')
    anom_fill   = PatternFill(fill_type='solid', fgColor='FD9091')
    debit_fill  = PatternFill(fill_type='solid', fgColor='C9E6FF')
    hdr_font    = Font(name='Aptos Narrow', size=10, bold=True, color='FFFFFF')
    dat_font    = Font(name='Aptos Narrow', size=10)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Import by Scenario'

    headers = list(export_df.columns)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = hdr_fill_sp if h in SP_HEADERS else hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal='center')

    debit_set = set(debit_positions)
    for r, (_, drow) in enumerate(export_df.iterrows()):
        is_anom  = anomalous_flags[r]
        is_debit = r in debit_set
        row_fill = debit_fill if is_debit else (anom_fill if is_anom else None)
        for c, val in enumerate(drow, 1):
            cell = ws.cell(row=r + 2, column=c, value=val)
            cell.font = dat_font
            if row_fill:
                cell.fill = row_fill

    for c, h in enumerate(headers, 1):
        cl = get_column_letter(c)
        mx = len(str(h))
        for r in range(2, min(22, ws.max_row + 1)):
            v = ws.cell(row=r, column=c).value
            if v:
                mx = max(mx, len(str(v)))
        ws.column_dimensions[cl].width = min(mx + 3, 40)

    ws.freeze_panes = 'A2'
    wb.save(str(output_path))

# ==============================================================================
# ACUMATICA FETCH
# ==============================================================================

def fetch_acumatica_shipments(username, password, base_url):
    table   = 'ShipmentsSubAc - JJ'
    encoded = table.replace(' ', '%20')
    url     = f'{base_url}/{encoded}'
    all_rows = []
    while url:
        resp = requests.get(
            url,
            auth=HTTPBasicAuth(username, password),
            headers={'Accept': 'application/json'},
            timeout=60,
        )
        if resp.status_code == 401:
            raise ValueError('Acumatica 401: credenciales incorrectas.')
        resp.raise_for_status()
        data = resp.json()
        all_rows.extend(data.get('value', []))
        url = data.get('@odata.nextLink')
    return pd.DataFrame(all_rows) if all_rows else pd.DataFrame()

# ==============================================================================
# VALIDATION & PRE-CHECKS
# ==============================================================================

def validate_carrier_file(carrier_path):
    """Returns (is_valid, carriers_found, warning)."""
    try:
        with pd.ExcelFile(_to_bytes(carrier_path)) as xl:
            carriers_found = []
            bad_cols       = []

            for sheet in xl.sheet_names:
                carrier = next((c for c in CARRIER_VENDOR_IDS if c.lower() in sheet.lower()), None)
                if not carrier:
                    continue
                try:
                    df_head = xl.parse(sheet, nrows=3)
                except Exception:
                    continue

                col_date   = find_col(df_head, ['doc date'])
                col_pro    = find_col(df_head, ['pronumber', 'pro number', 'vendor ref'])
                col_amount = find_col(df_head, ['import amount'])

                if col_date and col_pro and col_amount:
                    carriers_found.append(carrier)
                else:
                    bad_cols.append(carrier)
    except Exception as e:
        return False, [], f'Unable to read file — {e}'

    if not carriers_found and not bad_cols:
        return False, [], 'Wrong file — no carrier sheets found in this file.'

    if bad_cols and not carriers_found:
        return False, [], (f'Wrong format — required columns missing in: {", ".join(bad_cols)}. '
                           'Expected: Doc Date, Pro Number, Import Amount.')

    warning = None
    if bad_cols:
        warning = f'Missing columns in {", ".join(bad_cols)} — those tabs will be skipped.'

    return True, carriers_found, warning


def validate_priority1_file(p1_path, is_csv):
    """Returns (is_valid, error)."""
    try:
        df_head = (pd.read_csv(_to_bytes(p1_path), nrows=3, low_memory=False)
                   if is_csv else pd.read_excel(_to_bytes(p1_path), nrows=3))
    except Exception as e:
        return False, f'Unable to read file — {e}'

    col_inv = find_col(df_head, ['invoice number', 'invoice #', 'invoice'])
    if not col_inv:
        return False, 'Wrong file — "Invoice Number" column not found. Upload the Priority 1 Pending Invoices file.'
    return True, None


def validate_pacejet_file(pj_path, is_csv):
    """Returns (is_valid, error)."""
    try:
        df_head = (pd.read_csv(_to_bytes(pj_path), nrows=3, low_memory=False)
                   if is_csv else pd.read_excel(_to_bytes(pj_path), nrows=3))
    except Exception as e:
        return False, f'Unable to read file — {e}'

    col_uf3 = find_col(df_head, ['shipmentuserfield3', 'userfield3'])
    if col_uf3:
        return True, None

    # PaceJet BulkExport has many columns; if it's wide enough, accept it
    if len(df_head.columns) > 60:
        return True, None

    return False, 'Wrong file — "ShipmentUserField3" column not found. Upload the PaceJet BulkExport file.'


# ==============================================================================
# WWEX RAW FILE INGESTION — append mapped Pending rows to the master workbook
# ==============================================================================

_WWEX_RAW_REQUIRED_KEYWORDS = {
    'Invoice Date':      ['invoice date'],
    'Ship date':         ['ship date'],
    'Vendor Reference 2': ['vendor reference 2'],
    'Charge Total':      ['charge total'],
    'Invoice #':         ['invoice #'],
}

def _norm_header(v):
    return ' '.join(str(v).split()).strip().lower() if v is not None else ''

def validate_wwex_raw_file(raw_path):
    """Returns (is_valid, error)."""
    try:
        df_head = pd.read_excel(str(raw_path), nrows=3)
    except Exception as e:
        return False, f'Unable to read file — {e}'

    missing = [label for label, kws in _WWEX_RAW_REQUIRED_KEYWORDS.items()
               if not find_col(df_head, kws)]
    if missing:
        return False, f'Wrong format — required columns missing: {", ".join(missing)}.'
    return True, None


def _wwex_raw_key(pro_ref, po_sos, doc_date, amount):
    return (pro_ref, po_sos, safe_date_only(doc_date), amount)


def append_wwex_raw_to_master(raw_path, master_path, sheet_name='WWEX'):
    """Maps a raw WWEX carrier export into the master Carrier Import File's
    WWEX tab and appends new Pending rows after the last existing row.
    Returns {'added', 'skipped_duplicates', 'fee_rows_added'}."""
    if not WWEX_SMALL_PARCEL_ENABLED:
        raise ValueError('WWEX (Small Parcel) feature is currently disabled.')

    df = pd.read_excel(str(raw_path))

    col_invdate    = find_col(df, ['invoice date'])
    col_shipdate   = find_col(df, ['ship date'])
    col_vendorref  = find_col(df, ['vendor reference 2'])
    col_amount     = find_col(df, ['charge total'])
    col_invoicenum = find_col(df, ['invoice #'])
    col_billref1   = find_col(df, ['billing reference 1'])
    charge_type_cols = [c for c in df.columns if _norm_header(c).startswith('charge type')]

    if not all([col_invdate, col_shipdate, col_vendorref, col_amount, col_invoicenum]):
        raise ValueError('WWEX raw file missing required columns.')

    upload_stamp = f'Upload Freight Bill Processor Tool {datetime.now().strftime("%m/%d/%Y - %H:%M")}'

    mapped_rows = []
    for _, row in df.iterrows():
        fee_desc = None
        for ct_col in charge_type_cols:
            val = _norm_header(row.get(ct_col, '')).upper()
            if val in WWEX_FEE_DESCRIPTIONS:
                fee_desc = val
                break
        is_fee = fee_desc is not None

        invoicenum = safe_str(row.get(col_invoicenum, ''))
        tracking   = safe_str(row.get(col_vendorref, ''))
        # PO/SOS and BOL both carry the tracking number (Vendor Reference 2) —
        # it's what the reconciliation cascade in process_freight_bills matches
        # against Acumatica's shipment tracking numbers. Falls back to the
        # invoice # on fee lines, which have no tracking number of their own.
        po_sos_bol = tracking or invoicenum
        pro_number = safe_str(row.get(col_billref1, '')) if col_billref1 else ''

        mapped_rows.append({
            'carrier_date_of_file': safe_date(row.get(col_invdate)),
            'doc_date':             safe_date(row.get(col_shipdate)),
            'pro_vendor_ref':       invoicenum,
            'po_sos':               po_sos_bol,
            'bol':                  po_sos_bol,
            'pro_number':           pro_number,
            'amount':               safe_amount(row.get(col_amount, 0)),
            'notes':                f'{WWEX_FEE_NOTE_PREFIX}{fee_desc}' if is_fee else '',
            'is_fee':               is_fee,
            'addtl_notes':          upload_stamp,
        })

    try:
        wb = load_workbook(str(master_path))
    except PermissionError:
        raise ValueError('Carrier Import File está abierto en Excel — cerralo e intentá de nuevo.')

    ws = next((wb[s] for s in wb.sheetnames if s.strip().lower() == sheet_name.lower()), None)
    if ws is None:
        raise ValueError(f'Sheet "{sheet_name}" not found in master workbook.')

    header_idx = {_norm_header(c.value): c.column for c in ws[1] if c.value}
    required_headers = ['carrier date of file', 'doc date', 'pro number / vendor ref',
                         'open amount', 'import amount', 'po / sos', 'pro number', 'bol', 'notes',
                         'status (pending or imported)', 'quote amount', 'variance']
    missing_headers = [h for h in required_headers if h not in header_idx]
    if missing_headers:
        raise ValueError(f'Master WWEX tab missing expected columns: {", ".join(missing_headers)}.')

    col_carrier_date = header_idx['carrier date of file']
    col_doc_date     = header_idx['doc date']
    col_pro_ref      = header_idx['pro number / vendor ref']
    col_open_amt     = header_idx['open amount']
    col_import_amt   = header_idx['import amount']
    col_po_sos       = header_idx['po / sos']
    col_pro_number   = header_idx['pro number']
    col_bol          = header_idx['bol']
    col_notes        = header_idx['notes']
    col_status       = header_idx['status (pending or imported)']
    col_quote_amt    = header_idx['quote amount']
    col_variance     = header_idx['variance']
    # Optional — only written when the master tab actually has this column
    # (distinct from "Addtil Notes for Edilson", which is never touched).
    col_addtl_notes  = header_idx.get('additional notes')

    # Template row to clone formatting from (number format, font, alignment,
    # fill, border) — appended rows must look identical to existing ones, not
    # fall back to openpyxl's blank "General" default styling. Read the most
    # recent existing Pending row rather than assuming row 2: Pending rows
    # carry their own highlighting (e.g. Doc Date / Pro number-Vendor Ref in
    # yellow) that an old Imported row doesn't have.
    template_row = None
    for r in range(ws.max_row, 1, -1):
        if safe_str(ws.cell(row=r, column=col_status).value).lower() == 'pending':
            template_row = r
            break
    if template_row is None and ws.max_row >= 2:
        template_row = ws.max_row
    max_col = ws.max_column

    existing_keys = set()
    for r in range(2, ws.max_row + 1):
        existing_keys.add(_wwex_raw_key(
            safe_str(ws.cell(row=r, column=col_pro_ref).value),
            safe_str(ws.cell(row=r, column=col_po_sos).value),
            ws.cell(row=r, column=col_doc_date).value,
            safe_amount(ws.cell(row=r, column=col_import_amt).value),
        ))

    added = skipped = fee_rows_added = 0
    next_row = ws.max_row + 1

    for mr in mapped_rows:
        key = _wwex_raw_key(mr['pro_vendor_ref'], mr['po_sos'], mr['doc_date'], mr['amount'])
        if key in existing_keys:
            skipped += 1
            continue
        existing_keys.add(key)

        if template_row:
            for c in range(1, max_col + 1):
                src = ws.cell(row=template_row, column=c)
                dst = ws.cell(row=next_row, column=c)
                dst.number_format = src.number_format
                dst.font          = copy(src.font)
                dst.alignment     = copy(src.alignment)
                dst.fill          = copy(src.fill)
                dst.border        = copy(src.border)

        # Doc Date, Pro number/Vendor Ref, Import Amount, PO/SOS always get
        # yellow fill + black font — fixed regardless of the template row.
        for c in (col_doc_date, col_pro_ref, col_import_amt, col_po_sos):
            cell = ws.cell(row=next_row, column=c)
            base = cell.font
            cell.font = Font(name=base.name, size=base.size, bold=base.bold,
                              italic=base.italic, color='FF000000')
            cell.fill = PatternFill(fill_type='solid', fgColor='FFFFFF00')

        ws.cell(row=next_row, column=col_carrier_date, value=mr['carrier_date_of_file'])
        ws.cell(row=next_row, column=col_doc_date,     value=mr['doc_date'])
        ws.cell(row=next_row, column=col_pro_ref,      value=mr['pro_vendor_ref'])
        ws.cell(row=next_row, column=col_open_amt,     value=mr['amount'])
        ws.cell(row=next_row, column=col_import_amt,   value=mr['amount'])
        ws.cell(row=next_row, column=col_po_sos,       value=mr['po_sos'])
        ws.cell(row=next_row, column=col_pro_number,   value=mr['pro_number'])
        ws.cell(row=next_row, column=col_bol,          value=mr['bol'])
        ws.cell(row=next_row, column=col_quote_amt,    value=0)
        ws.cell(row=next_row, column=col_variance,     value=0)
        ws.cell(row=next_row, column=col_notes,        value=mr['notes'])
        ws.cell(row=next_row, column=col_status,       value='Pending')

        if col_addtl_notes:
            existing_note = safe_str(ws.cell(row=next_row, column=col_addtl_notes).value)
            new_note = f'{existing_note} {mr["addtl_notes"]}'.strip() if existing_note else mr['addtl_notes']
            ws.cell(row=next_row, column=col_addtl_notes, value=new_note)

        next_row += 1
        added += 1
        if mr['is_fee']:
            fee_rows_added += 1

    if added:
        try:
            wb.save(str(master_path))
        except PermissionError:
            raise ValueError('Carrier Import File está abierto en Excel — cerralo e intentá de nuevo.')

    return {'added': added, 'skipped_duplicates': skipped, 'fee_rows_added': fee_rows_added}


def _vectorized_amount(df, col):
    """Parse an amount column to float Series without Python-level loops."""
    return pd.to_numeric(
        df[col].astype(str).str.replace('$', '', regex=False)
                           .str.replace(',', '', regex=False).str.strip(),
        errors='coerce').fillna(0.0)


def analyze_carrier_file(carrier_path):
    """Single read of the carrier file. Returns:
      {summary: {total, carriers:[{name,count}]}, needs_priority1: bool,
       missing_cells: [(sheet, column, excel_row)], missing_message: str|None}
    Replaces the separate get_carrier_summary + check_needs_priority1 calls."""
    sheets = pd.read_excel(_to_bytes(carrier_path), sheet_name=None)
    result        = []
    total         = 0
    needs_p1      = False
    missing_cells = []

    for sheet_name, df in sheets.items():
        carrier = next((c for c in CARRIER_VENDOR_IDS if c.lower() in sheet_name.lower()), None)
        if not carrier:
            continue

        col_status = find_col(df, ['pending', 'imported', 'status'])
        col_amount = find_col(df, ['import amount'])
        col_date   = find_col(df, ['doc date'])

        if col_status and col_amount:
            is_pending = df[col_status].astype(str).str.strip().str.lower() == 'pending'
            supports_neg = carrier in NEGATIVE_AMOUNT_CONFIG
            parsed = _vectorized_amount(df, col_amount)
            has_amount = (parsed != 0) if supports_neg else (parsed > 0)
            mask  = is_pending & has_amount
            count = int(mask.sum())
        elif col_status:
            mask  = df[col_status].astype(str).str.strip().str.lower() == 'pending'
            count = int(mask.sum())
        else:
            mask  = df.notna().any(axis=1)
            count = int(mask.sum())

        # Same blank-Doc-Date check process_freight_bills enforces, run here so
        # the user is told at upload time instead of after the whole pipeline.
        if col_date is not None:
            missing_cells.extend(
                (sheet_name, str(col_date), int(idx) + 2)
                for idx in df.index[mask]
                if is_missing_date(df.at[idx, col_date])
            )

        if count > 0:
            result.append({'name': carrier, 'count': count})
            total += count
            if carrier == 'Priority 1':
                needs_p1 = True

    return {
        'summary':         {'total': total, 'carriers': result},
        'needs_priority1': needs_p1,
        'missing_cells':   missing_cells,
        'missing_message': _missing_cells_message(missing_cells) if missing_cells else None,
    }


def classify_file(file_path, filename):
    """Auto-detect file type from content structure.
    Returns ('carrier' | 'priority1' | 'pacejet' | 'unknown', error_or_None).
    """
    is_csv = (filename or '').lower().endswith('.csv')
    try:
        if not is_csv:
            try:
                with pd.ExcelFile(_to_bytes(file_path)) as xl:
                    sheet_names = xl.sheet_names
                    # Carrier: has sheets matching known carrier names
                    if any(any(c.lower() in s.lower() for c in CARRIER_VENDOR_IDS)
                           for s in sheet_names):
                        return 'carrier', None
                    if not sheet_names:
                        return 'unknown', 'Excel file has no sheets.'
                    df_head = xl.parse(sheet_names[0], nrows=3)
            except Exception as e:
                return 'unknown', f'Cannot read Excel file: {e}'
        else:
            try:
                df_head = pd.read_csv(str(file_path), nrows=3, low_memory=False)
            except Exception as e:
                return 'unknown', f'Cannot read CSV file: {e}'

        cols = [str(c).strip() for c in df_head.columns]
        cols_lower = [c.lower() for c in cols]

        # WWEX (Small Parcel) raw export — checked before PaceJet/Priority 1
        # since its own "Invoice #" / "Airbill #" columns would otherwise
        # misfire those checks. Disabled via WWEX_SMALL_PARCEL_ENABLED.
        if WWEX_SMALL_PARCEL_ENABLED and {'airbill #', 'scac', 'charge type 1'} <= set(cols_lower):
            return 'wwex_raw', None

        # PaceJet: many Shipment* columns
        if (sum(1 for c in cols if c.startswith('Shipment')) > 5
                or any('shipmentuserfield3' in c for c in cols_lower)):
            return 'pacejet', None

        # Priority 1: Invoice Number column
        if any('invoice number' in c or 'invoice #' in c for c in cols_lower):
            return 'priority1', None

        return 'unknown', ('Unrecognized file — expected a Carrier Import, '
                           'Priority 1 Pending Invoices, or PaceJet Export file.')
    except Exception as e:
        return 'unknown', f'Error analyzing file: {e}'



# ==============================================================================
# MAIN PROCESSING
# ==============================================================================

def _missing_cells_message(cells, max_rows_listed=6):
    """Builds the error-toast text for required cells left blank in the master
    workbook. Rows are grouped per sheet+column so a run of blanks reads as
    'rows 417, 418, 419' instead of three separate clauses. Format is
    'Title — detail' so the front-end toast splits it into heading and body."""
    groups = {}
    for sheet, col, row in cells:
        groups.setdefault((sheet, col), []).append(row)

    def _rows(rows):
        shown = ', '.join(str(r) for r in rows[:max_rows_listed])
        if len(rows) > max_rows_listed:
            shown += f' and {len(rows) - max_rows_listed} more'
        return f'{"row" if len(rows) == 1 else "rows"} {shown}'

    total   = len(cells)
    closing = 'Fill it in and try again.' if total == 1 else 'Fill them in and try again.'

    # Single sheet+column reads as a plain sentence; several get a count and a
    # semicolon-separated list, same shape as the other upload errors.
    if len(groups) == 1:
        (sheet, col), rows = next(iter(groups.items()))
        return f'Missing data — "{col}" is empty in the {sheet} sheet, {_rows(rows)}. {closing}'

    parts = [f'"{col}" in {sheet} {_rows(rows)}' for (sheet, col), rows in groups.items()]
    return (f'Missing data — {total} required cells are empty: '
            f'{"; ".join(parts)}. {closing}')


def process_freight_bills(carrier_path, priority1_df, pacejet_df, shipments_df, output_dir):
    all_sheets = pd.read_excel(str(carrier_path), sheet_name=None)

    # Build Acumatica shipment lookup tables
    ordernbr_lookup, shipmentnbr_lookup, tracking_lookup = build_shipments_lookups(shipments_df)

    # Build Priority 1 invoice → PRO / BOL / SO lookup
    p1_lookup = build_p1_lookup(priority1_df)

    # Build PaceJet records + amount index for step-4 fuzzy matching
    pacejet_records, pacejet_amount_index = build_pacejet_records(pacejet_df)

    all_bills     = []
    all_debit_adj = []
    counter       = 1
    counter_debit = 1
    now           = pd.Timestamp.now().strftime('%m-%d-%Y')
    wwex_stats    = {'bills': 0, 'debits': 0}
    # (sheet, column, excel_row) for required cells left blank in the master
    # workbook. Never defaulted — the run stops and the user fills them in.
    missing_cells = []

    for sheet_name, df in all_sheets.items():
        matched_carrier = next(
            (c for c in CARRIER_VENDOR_IDS if c.lower() in sheet_name.lower()), None
        )
        if not matched_carrier:
            continue

        vendor_id         = CARRIER_VENDOR_IDS[matched_carrier]
        supports_negative = matched_carrier in NEGATIVE_AMOUNT_CONFIG
        if supports_negative:
            counter_debit = 1

        col_date   = find_col(df, ['doc date'])
        col_pro    = find_col(df, ['pronumber', 'pro number', 'vendor ref'])
        col_amount = find_col(df, ['import amount'])
        col_po     = find_col(df, ['po / sos', 'po/', 'sos'])
        col_status = find_col(df, ['pending', 'imported', 'status'])
        col_notes  = find_col(df, ['notes']) if matched_carrier == 'WWEX' else None

        if not all([col_date, col_pro, col_amount]):
            continue

        if col_status:
            is_pending = df[col_status].astype(str).str.strip().str.lower() == 'pending'
        else:
            is_pending = pd.Series([True] * len(df), index=df.index)

        parsed_amounts = _vectorized_amount(df, col_amount)
        has_amount     = (parsed_amounts != 0) if supports_negative else (parsed_amounts > 0)
        pending        = df[is_pending & has_amount].copy()

        # Doc Date drives Document Date and Post Period — there is no sane
        # default, so blanks are collected and reported instead of guessed.
        # +2 converts the 0-based DataFrame index to the Excel row number
        # (row 1 is the header).
        blank_dates = {idx for idx in pending.index
                       if is_missing_date(pending.at[idx, col_date])}
        missing_cells.extend((sheet_name, str(col_date), int(idx) + 2)
                             for idx in sorted(blank_dates))

        for idx, row in pending.iterrows():
            if idx in blank_dates:
                continue
            pronumber  = safe_str(row.get(col_pro,    ''))
            po_sos     = safe_str(row.get(col_po,     '')) if col_po else ''
            doc_date   = safe_date(row.get(col_date,  now))
            amount     = safe_amount(row.get(col_amount, 0))

            customer      = ''
            customer_name = ''
            is_anomalous  = False
            case          = ''

            wwex_fee_notes = (safe_str(row.get(col_notes, ''))
                              if matched_carrier == 'WWEX' and col_notes else '')
            is_wwex_fee = wwex_fee_notes.startswith(WWEX_FEE_NOTE_PREFIX)
            line_desc = po_sos

            # ------------------------------------------------------------------
            # PRIORITY 1 — 4-step cascading lookup (mirrors base.py exactly)
            # ------------------------------------------------------------------
            if matched_carrier == 'Priority 1':
                p1d        = p1_lookup.get(pronumber, {})
                p1_pro     = p1d.get('PRO', '')
                p1_bol     = p1d.get('BOL', '')
                p1_so      = p1d.get('SO', '')
                p1_carrier = p1d.get('Carrier', '')
                p1_mode    = p1d.get('CarrierMode', '')
                p1_rembal  = p1d.get('RemainingBalance', 0.0)
                p1_actship = p1d.get('ActualShip')

                match = None

                # Step 1: PRO → ShipmentsSubAc TrackingNumber
                if p1_pro and p1_pro in tracking_lookup:
                    match, case = tracking_lookup[p1_pro], 'Direct'

                # Step 2: BOL → ShipmentsSubAc TrackingNumber
                elif p1_bol and p1_bol in tracking_lookup:
                    match, case = tracking_lookup[p1_bol], 'Direct'

                # Step 3: SO present → OrderNbr lookup (return shipment)
                elif p1_so:
                    case  = 'Return'
                    match = ordernbr_lookup.get(p1_so)
                    if not match:
                        is_anomalous = True

                # Step 4: PaceJet fuzzy match (amount + carrier + mode + date)
                elif pacejet_records:
                    uf3 = find_pacejet_match(p1_rembal, p1_carrier, p1_mode, p1_actship,
                                             pacejet_records, pacejet_amount_index)
                    if uf3 and uf3 in ordernbr_lookup:
                        match, case = ordernbr_lookup[uf3], 'Via PaceJet'
                    else:
                        case, is_anomalous = 'Review', True

                # No path found
                else:
                    case, is_anomalous = 'Review', True

                if match:
                    customer      = match['Customer']
                    customer_name = match['CustomerName']

            # ------------------------------------------------------------------
            # WWEX small-parcel fee lines (auto-inserted by append_wwex_raw_to_master)
            # — carrier-generated fees, not shipments, so they skip the customer
            # lookup entirely and post to their own fixed GL/Subaccount.
            # ------------------------------------------------------------------
            elif is_wwex_fee:
                case = 'Fee'
                line_desc = wwex_fee_notes[len(WWEX_FEE_NOTE_PREFIX):].title()

            # ------------------------------------------------------------------
            # ALL OTHER CARRIERS — generic cascade
            # ------------------------------------------------------------------
            else:
                match = None
                for key in [po_sos, pronumber]:
                    if not key:
                        continue
                    match = (ordernbr_lookup.get(key)
                             or shipmentnbr_lookup.get(key)
                             or tracking_lookup.get(key))
                    if match:
                        break

                if match:
                    customer      = match['Customer']
                    customer_name = match['CustomerName']
                    case          = 'Direct'
                else:
                    is_anomalous, case = True, 'Review'

            account    = WWEX_FEE_CONFIG['account']    if is_wwex_fee else ACCOUNT
            subaccount = (WWEX_FEE_CONFIG['subaccount'] if is_wwex_fee
                          else (SUBACCOUNTS.get(customer, DEFAULT_SUBACCOUNT) if customer else ''))

            if supports_negative and amount < 0:
                all_debit_adj.append(
                    _build_debit_adj_row(matched_carrier, counter_debit, doc_date,
                                         vendor_id, pronumber, po_sos, amount, now)
                )
                counter_debit += 1
                if matched_carrier == 'WWEX':
                    wwex_stats['debits'] += 1
            else:
                all_bills.append({
                    'Type': 'Bill',
                    'Reference Nbr.': counter,
                    'Document Date': doc_date.strftime('%m-%d-%Y'),
                    'Post Period': doc_date.strftime('%m-%Y'),
                    'Document Description': 'Bill uploaded from distribution file',
                    'Vendor ID': vendor_id,
                    'Carrier': matched_carrier,
                    'Location': '',
                    'Vendor Ref.': pronumber,
                    'Branch': 'MAIN',
                    'Line Branch': 'MAIN',
                    'Line Description': line_desc,
                    'Quantity': 1,
                    'Unit Cost': amount,
                    'Amount': amount,
                    'Account': account,
                    'Sub': customer_name,
                    'Subaccount': subaccount,
                    'Ap received Bills date ': now,
                    'Invoice Number': pronumber,
                    'Case': case,
                    '_anomalous': is_anomalous,
                })
                counter += 1
                if matched_carrier == 'WWEX':
                    wwex_stats['bills'] += 1

    if missing_cells:
        raise ValueError(_missing_cells_message(missing_cells))

    if not all_bills and not all_debit_adj:
        raise ValueError('No pending invoices found in the carrier file.')

    bills_df     = pd.DataFrame(all_bills)
    debit_adj_df = pd.DataFrame(all_debit_adj)
    import_df    = pd.concat([bills_df, debit_adj_df], ignore_index=True)

    debit_positions  = import_df.index[import_df['Type'] == 'Debit Adj.'].tolist()
    anomalous_flags  = (import_df['_anomalous'].tolist()
                        if '_anomalous' in import_df.columns
                        else [False] * len(import_df))

    export_df = import_df.drop(columns=['_anomalous'], errors='ignore')

    timestamp   = datetime.now().strftime('%m%d%y_%H%M')
    filename    = f'Carrier_Bills_{timestamp}.xlsx'
    output_path = Path(output_dir) / filename
    _write_formatted_excel(output_path, export_df, anomalous_flags, debit_positions)

    total_bills  = int((import_df['Type'] == 'Bill').sum())
    total_debits = int((import_df['Type'] == 'Debit Adj.').sum())
    bill_amount  = float(import_df.loc[import_df['Type'] == 'Bill',       'Amount'].sum()) if total_bills  else 0.0
    debit_amount = float(import_df.loc[import_df['Type'] == 'Debit Adj.', 'Amount'].sum()) if total_debits else 0.0

    # Priority 1 case breakdown
    p1_cases = {}
    if not bills_df.empty and 'Case' in bills_df.columns and 'Carrier' in bills_df.columns:
        p1_rows = bills_df[bills_df['Carrier'] == 'Priority 1']
        for case_name in ['Direct', 'Return', 'Via PaceJet', 'Review']:
            n = int((p1_rows['Case'] == case_name).sum())
            if n > 0:
                p1_cases[case_name] = n

    return {
        'filename':      filename,
        'total_bills':   total_bills,
        'total_debits':  total_debits,
        'total_records': total_bills + total_debits,
        'bill_amount':   bill_amount,
        'debit_amount':  debit_amount,
        'wwex_bills':    wwex_stats['bills'],
        'wwex_debits':   wwex_stats['debits'],
        'p1_cases':      p1_cases,
    }
