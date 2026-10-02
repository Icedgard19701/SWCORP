import io
import re
import pandas as pd
import requests
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from copy import copy
from requests.auth import HTTPBasicAuth
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus
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

# Master switch for the Small Parcels ingestion feature (WWEX raw export and
# FedEx invoice PDF). Flip to False to disable both without touching anything
# else — the Small Parcels routes then return 404.
SMALL_PARCEL_ENABLED = True

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

# Notes written into the flat carrier tabs carry a machine-readable head (the
# fee prefix, the FedEx invoice number) followed by the upload stamp. Readers
# split on this separator instead of consuming the whole cell.
NOTE_SEPARATOR = ' | '

# FedEx invoice PDFs print a form code right after the "P.O.#:" label when the
# shipment has no purchase order. Real POs on these invoices run 6-17 chars, so
# anything outside that range is treated as noise rather than a PO.
FEDEX_PO_NOISE_TOKENS = {'177'}
FEDEX_PO_MIN_LENGTH = 6
FEDEX_PO_MAX_LENGTH = 17

class BlockingImportError(ValueError):
    """An import that stopped for a reason the user has to resolve — a charge
    the parser cannot account for, a workbook it must not write to. Separate
    from a plain ValueError so the web layer can mark it blocking and the
    browser can hold the message on screen until it is dismissed by hand,
    rather than flashing it for a few seconds and leaving no trace."""


# A parsed FedEx invoice must add up to the total printed on the invoice itself
# before anything is written to the master workbook.
FEDEX_TOTAL_TOLERANCE = 0.01

# FedEx charges that belong to the account rather than to any one shipment
# (a scheduled-pickup fee, a late fee on an earlier invoice) skip the customer
# lookup and post to their own fixed GL/Subaccount. Deliberately its own config
# rather than a reuse of WWEX_FEE_CONFIG: the two carriers bill different kinds
# of fee and must stay free to be routed differently.
FEDEX_FEE_CONFIG = {
    'account':    '410160',
    'subaccount': 'DIS-000000-000000-0000000000',
}

_MASTER_OPEN_ERROR = ('Carrier Import File is open in Excel — '
                      'close it and try again.')

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
    """Returns (ordernbr_lookup, shipmentnbr_lookup, tracking_lookup,
    custorder_lookup). Each maps key -> {Customer, CustomerName, OrderNbr}.

    custorder_lookup is keyed on Acumatica's CustomerOrderNbr — the customer's
    own PO, which is what the carriers print in their reference fields. A PO is
    only unique per customer, so keys claimed by more than one customer are
    dropped rather than resolved to whichever shipment came first."""
    ordernbr_lookup    = {}
    shipmentnbr_lookup = {}
    tracking_lookup    = {}
    custorder_lookup   = {}

    if shipments_df.empty:
        return ordernbr_lookup, shipmentnbr_lookup, tracking_lookup, custorder_lookup

    col_order    = find_col(shipments_df, ['ordernbr', 'order nbr'])
    col_shipment = find_col(shipments_df, ['shipmentnbr', 'shipment nbr'])
    col_tracking = find_col(shipments_df, ['trackingnumber', 'tracking number'])
    col_custorder = find_col(shipments_df, ['customerordernbr', 'customer order nbr'])
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

    orders     = _key_col(col_order)
    shipments  = _key_col(col_shipment)
    trackings  = _key_col(col_tracking)
    custorders = _key_col(col_custorder)
    customers  = _str_col(col_customer)
    custnames  = _str_col(col_custname)

    # A customer PO repeats across that customer's own shipments, so the first
    # one wins as everywhere else; only a PO claimed by two different customers
    # is ambiguous, and those keys are removed below.
    custorder_owners: dict = {}

    for o, s, t, p, c, n in zip(orders, shipments, trackings, custorders,
                                customers, custnames):
        cust_data = {'Customer': c, 'CustomerName': n, 'OrderNbr': o}
        if o and o not in ordernbr_lookup:    ordernbr_lookup[o]    = cust_data
        if s and s not in shipmentnbr_lookup: shipmentnbr_lookup[s] = cust_data
        if t and t not in tracking_lookup:    tracking_lookup[t]    = cust_data
        if p:
            custorder_owners.setdefault(p, set()).add(c)
            if p not in custorder_lookup:     custorder_lookup[p]   = cust_data

    for p, owners in custorder_owners.items():
        if len(owners) > 1:
            custorder_lookup.pop(p, None)

    return ordernbr_lookup, shipmentnbr_lookup, tracking_lookup, custorder_lookup


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


def build_pacejet_parcel_bridge(pacejet_df):
    """Maps parcel tracking number -> PaceJet shipment control number.

    Small-parcel invoices carry a tracking number and, at best, the customer's
    PO. When neither is in Acumatica, PaceJet still knows the shipment it
    rated: its control number is what Acumatica stores as ShipmentNbr, so the
    tracking number resolves to a customer through PaceJet even though the
    carrier never printed an Acumatica reference.

    Both tracking columns are indexed — they hold the same value on
    single-package shipments and differ on multi-package ones, where the carrier
    bills each package under its own number."""
    if pacejet_df is None or pacejet_df.empty:
        return {}

    col_pkg_track = find_col(pacejet_df, ['shipmentpackagetrackingid', 'packagetrackingid'])
    col_track     = find_col(pacejet_df, ['shipmenttrackingid', 'trackingid'])
    # ShipmentUserField1 mirrors ShipmentControlNumber in the BulkExport; either
    # one is the value Acumatica knows.
    col_control   = (find_col(pacejet_df, ['shipmentcontrolnumber', 'controlnumber'])
                     or find_col(pacejet_df, ['shipmentuserfield1']))
    if not col_control or not (col_pkg_track or col_track):
        return {}

    def _col(col):
        if not col: return [''] * len(pacejet_df)
        return pacejet_df[col].apply(safe_str).tolist()

    bridge = {}
    for pkg, trk, ctl in zip(_col(col_pkg_track), _col(col_track), _col(col_control)):
        if not ctl:
            continue
        for key in (pkg, trk):
            if key and key not in bridge:
                bridge[key] = ctl
    return bridge


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


def _write_formatted_excel(output_path, export_df, anomalous_flags, debit_positions,
                           notice_flags=None):
    """notice_flags marks rows that are complete but were filled in by a weaker
    rule than the rest — worth a second look without being an error, so they get
    a neutral grey rather than the red of a row that is actually missing data."""
    SP_HEADERS = {'Sub', 'Invoice Number', 'Carrier', 'Case'}
    hdr_fill    = PatternFill(fill_type='solid', fgColor='00375C')
    hdr_fill_sp = PatternFill(fill_type='solid', fgColor='0062A4')
    anom_fill   = PatternFill(fill_type='solid', fgColor='FD9091')
    debit_fill  = PatternFill(fill_type='solid', fgColor='C9E6FF')
    notice_fill = PatternFill(fill_type='solid', fgColor='F2F2F2')
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
        is_anom   = anomalous_flags[r]
        is_debit  = r in debit_set
        is_notice = bool(notice_flags[r]) if notice_flags else False
        row_fill = (debit_fill if is_debit
                    else anom_fill if is_anom
                    else notice_fill if is_notice
                    else None)
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

# A slow or dropped request is retried before the lookup gives up.
ACUMATICA_PAGE_ATTEMPTS = 3


# ------------------------------------------------------------------------------
# Filtered lookup (same pattern as Lowe's Invoice Reconciler, recon/odata.py).
#
# The full inquiry is ~150k rows and took minutes; a run only ever looks rows up
# by a handful of keys, so the inquiry is asked with $filter for the values of
# the files in hand: `Field eq 'v1' or Field eq 'v2' ...`, one field per
# request, in batches that keep the URL under Acumatica's limit, a few batches
# side by side. The rows that come back have the same columns as the full
# download, so build_shipments_lookups and every cascade read them unchanged.
# ------------------------------------------------------------------------------

ACUMATICA_SHIPMENTS_GI   = 'ShipmentsSubAc - JJ'
# Columns the lookups read (build_shipments_lookups). Asking only these keeps
# each answer small.
ACUMATICA_SHIPMENT_COLS  = ('OrderNbr', 'ShipmentNbr', 'TrackingNumber',
                            'CustomerOrderNbr', 'Customer', 'CustomerName')
ACUMATICA_LOOKUP_FIELDS  = ('TrackingNumber', 'OrderNbr', 'ShipmentNbr', 'CustomerOrderNbr')
# Acumatica's IIS answers a bare 404 past ~2048 URL characters (measured in
# LIR on 2026-09-29: 1802 passes, 2195 fails); keep a margin.
ACUMATICA_MAX_URL        = 1900
ACUMATICA_LOOKUP_THREADS = 4
ACUMATICA_LOOKUP_TIMEOUT = 60


def _shipments_gi_url(base_url):
    return f"{base_url.rstrip('/')}/{ACUMATICA_SHIPMENTS_GI.replace(' ', '%20')}"


def _acumatica_get(url, params, username, password, timeout):
    """GET with the page retry of the full fetch; 401 reads as bad credentials."""
    for attempt in range(1, ACUMATICA_PAGE_ATTEMPTS + 1):
        try:
            resp = requests.get(url, params=params,
                                auth=HTTPBasicAuth(username, password),
                                headers={'Accept': 'application/json'},
                                timeout=timeout)
            break
        except requests.RequestException as e:
            if attempt == ACUMATICA_PAGE_ATTEMPTS:
                raise ValueError(f'Acumatica did not answer after '
                                 f'{ACUMATICA_PAGE_ATTEMPTS} attempts: {e}') from e
    if resp.status_code == 401:
        raise ValueError('Acumatica 401: wrong credentials.')
    resp.raise_for_status()
    return resp.json()


def probe_acumatica(username, password, base_url):
    """Raises unless Acumatica answers: one row of the shipments inquiry. The
    page asks this on arrival, so the Acumatica pill shows the connection
    before any file is in; the rows themselves are asked with the files."""
    _acumatica_get(_shipments_gi_url(base_url),
                   {'$top': 1, '$select': 'OrderNbr', '$format': 'json'},
                   username, password, ACUMATICA_LOOKUP_TIMEOUT)


def _filter_batches(field_values, base_len):
    """`$filter` expressions, one field each, as many values as fit the URL."""
    joiner = len(quote_plus(' or '))
    for field in ACUMATICA_LOOKUP_FIELDS:
        batch, used = [], base_len
        for value in sorted(field_values.get(field, ())):
            literal = str(value).replace("'", "''")
            term = f"{field} eq '{literal}'"
            cost = len(quote_plus(term)) + joiner
            if batch and used + cost > ACUMATICA_MAX_URL:
                yield ' or '.join(batch)
                batch, used = [], base_len
            batch.append(term)
            used += cost
        if batch:
            yield ' or '.join(batch)


def fetch_acumatica_shipments_matching(username, password, base_url, field_values):
    """Shipments rows whose field equals any of its values.

    field_values: {GI field: iterable of values}. Acumatica compares
    case-insensitively. Returns a DataFrame with ACUMATICA_SHIPMENT_COLS; rows
    matched through more than one field come back once, in the order the
    batches were asked."""
    wanted = {f: {safe_str(v) for v in field_values.get(f, ())} - {''}
              for f in ACUMATICA_LOOKUP_FIELDS}
    url    = _shipments_gi_url(base_url)
    params = {'$select': ','.join(ACUMATICA_SHIPMENT_COLS), '$format': 'json'}
    base_len = (len(requests.Request('GET', url, params=params).prepare().url)
                + len('&%24filter='))
    filters = list(_filter_batches(wanted, base_len))
    if not filters:
        return pd.DataFrame(columns=list(ACUMATICA_SHIPMENT_COLS))

    def ask(expr):
        rows, next_url, p = [], url, dict(params, **{'$filter': expr})
        while next_url:
            data = _acumatica_get(next_url, p, username, password, ACUMATICA_LOOKUP_TIMEOUT)
            rows.extend(data.get('value', []))
            # A server link already carries the paging state.
            next_url, p = data.get('@odata.nextLink') or data.get('odata.nextLink'), None
        return rows

    with ThreadPoolExecutor(max_workers=ACUMATICA_LOOKUP_THREADS,
                            thread_name_prefix='acu-lookup') as pool:
        found = [r for batch in pool.map(ask, filters) for r in batch]
    df = pd.DataFrame(found, columns=list(ACUMATICA_SHIPMENT_COLS))
    return df.drop_duplicates(ignore_index=True)


def freight_lookup_values(carrier_path, priority1_df, pacejet_df):
    """Every value process_freight_bills can look up in Acumatica, by GI field.

    Mirrors its cascade (a superset: every value a branch might ask, whichever
    branch ends up deciding), so the filtered rows answer exactly what the full
    inquiry would have."""
    want = {f: set() for f in ACUMATICA_LOOKUP_FIELDS}
    if not carrier_path or not Path(carrier_path).exists():
        return want
    p1_lookup = (build_p1_lookup(priority1_df)
                 if priority1_df is not None and not priority1_df.empty else {})
    pacejet_records, pacejet_amount_index = (
        build_pacejet_records(pacejet_df)
        if pacejet_df is not None and not pacejet_df.empty else ([], {}))
    all_sheets = pd.read_excel(str(carrier_path), sheet_name=None)
    for sheet_name, df in all_sheets.items():
        if is_detail_sheet(sheet_name):
            continue
        carrier = next((c for c in CARRIER_VENDOR_IDS if c.lower() in sheet_name.lower()), None)
        if not carrier:
            continue
        col_pro    = find_col(df, ['pronumber', 'pro number', 'vendor ref'])
        col_po     = find_col(df, ['po / sos', 'po/', 'sos'])
        col_status = find_col(df, ['pending', 'imported', 'status'])
        if not col_pro:
            continue
        if col_status:
            df = df[df[col_status].astype(str).str.strip().str.lower() == 'pending']
        for _, row in df.iterrows():
            pro = safe_str(row.get(col_pro, ''))
            po  = safe_str(row.get(col_po, '')) if col_po else ''
            if carrier == 'Priority 1':
                p1d = p1_lookup.get(pro, {})
                want['TrackingNumber'].update((p1d.get('PRO', ''), p1d.get('BOL', '')))
                want['OrderNbr'].add(p1d.get('SO', ''))
                if pacejet_records and p1d:
                    uf3 = find_pacejet_match(p1d.get('RemainingBalance', 0.0), p1d.get('Carrier', ''),
                                             p1d.get('CarrierMode', ''), p1d.get('ActualShip'),
                                             pacejet_records, pacejet_amount_index)
                    want['OrderNbr'].add(uf3 or '')
                continue
            want['TrackingNumber'].add(pro)
            for key in (po, pro):
                want['OrderNbr'].add(key)
                want['ShipmentNbr'].add(key)
                want['CustomerOrderNbr'].add(key)
    return {f: v - {''} for f, v in want.items()}


def small_parcel_lookup_values(rows, pacejet_df=None):
    """Every value _sp_import_subaccount can look up in Acumatica, by GI field."""
    want = {f: set() for f in ACUMATICA_LOOKUP_FIELDS}
    bridge = build_pacejet_parcel_bridge(pacejet_df)
    for row in rows:
        if row.get('is_fee'):
            continue
        tracking, po = row.get('pro_vendor_ref') or '', row.get('po_sos') or ''
        want['TrackingNumber'].add(tracking)
        for key in (po, tracking):
            want['OrderNbr'].add(key)
            want['ShipmentNbr'].add(key)
            want['CustomerOrderNbr'].add(key)
        control = bridge.get(tracking) if tracking else None
        if control:
            want['ShipmentNbr'].add(control)
            want['OrderNbr'].add(control)
    return {f: {safe_str(x) for x in v} - {''} for f, v in want.items()}


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
# SMALL PARCELS INGESTION — append mapped Pending rows to the master workbook
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


def _upload_stamp():
    """Trace stamp written into the Notes column of every appended row, so the
    AP team can tell which rows this tool added and when."""
    return f'Upload Freight Bill Processor Tool {datetime.now().strftime("%m/%d/%Y - %H:%M")}'

def is_detail_sheet(sheet_name):
    """True for the "<Carrier> - SP" tabs of the Small Parcels workbook. Those
    are the AP team's reconciliation detail; the flat "<Carrier>" tab is the
    import feed. Both carry the carrier name, so reading both would bill every
    small-parcel row twice."""
    return _norm_header(sheet_name).endswith('- sp')


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


def _carrier_tab_row_key(pro_ref, po_sos, doc_date, amount):
    """Dedupe identity of a row in a flat carrier tab. Re-uploading the same
    carrier file must not duplicate rows that are already there."""
    return (pro_ref, po_sos, safe_date_only(doc_date), amount)


def _extend_table_ranges(ws, last_row):
    """Grow every Excel Table defined on the sheet down to last_row. The flat
    carrier tabs are real Tables, so appended rows land outside the table (no
    banding, excluded from structured references) unless the ref is widened."""
    for tbl in (getattr(ws, 'tables', None) or {}).values():
        start, end = tbl.ref.split(':')
        m = re.match(r'([A-Z]+)(\d+)$', end)
        if not m or last_row <= int(m.group(2)):
            continue
        tbl.ref = f'{start}:{m.group(1)}{last_row}'
        if tbl.autoFilter is not None:
            tbl.autoFilter.ref = tbl.ref


# Raised when the rows are gone from the workbook the moment it is read back.
# Phrased for the AP user, not the developer: the fix is always at their end.
_MASTER_LOST_ROWS_ERROR = (
    'The workbook did not keep the appended rows — {lost} of {total} were '
    'already gone when "{name}" was read back straight after saving. Excel or '
    'OneDrive is holding another copy of it and overwrote the write. Close the '
    'workbook everywhere (other machines and Excel Online included), wait for '
    'OneDrive to finish syncing, then import the file again. No bill was '
    'generated for these rows.'
)


def _verify_carrier_rows_persisted(master_path, sheet_name, written):
    """Re-reads the workbook straight after the save and confirms the appended
    rows are really in it.

    A OneDrive-synced workbook open in Excel with AutoSave on is not locked on
    disk: openpyxl saves over it without raising, and Excel later writes back
    the copy it still holds in memory, silently discarding every appended row.
    A save that returns cleanly is therefore not evidence the rows landed — the
    only honest check is reading the file back. Raises ValueError when rows are
    missing so the caller fails loudly instead of reporting a clean import.

    Catches the loss only when it has already happened by the time this runs.
    Excel can still overwrite the file minutes later, which nothing inside this
    process can see; the guard for that is _excel_lock_files upstream, and it is
    not complete either."""
    if not written:
        return

    try:
        wb = load_workbook(str(master_path))
    except Exception as e:
        raise ValueError(f'The rows were written but "{Path(master_path).name}" '
                         f'could not be read back to confirm they landed — {e}')

    try:
        ws = next((wb[s] for s in wb.sheetnames
                   if s.strip().lower() == sheet_name.lower()), None)
        if ws is None:
            raise ValueError(_MASTER_LOST_ROWS_ERROR.format(
                lost=len(written), total=len(written), name=Path(master_path).name))

        header_idx = {_norm_header(c.value): c.column for c in ws[1] if c.value}
        col_pro_ref    = header_idx.get('pro number / vendor ref')
        col_import_amt = header_idx.get('import amount')
        # Shape was validated before the append, so a column missing here means
        # the sheet came back as something else entirely — treated as total loss
        # rather than quietly passing.
        if not (col_pro_ref and col_import_amt):
            raise ValueError(_MASTER_LOST_ROWS_ERROR.format(
                lost=len(written), total=len(written), name=Path(master_path).name))

        lost = 0
        for mr in written:
            row = mr.get('_row')
            if not row or row > ws.max_row:
                lost += 1
                continue
            same_ref = (safe_str(ws.cell(row=row, column=col_pro_ref).value)
                        == safe_str(mr['pro_vendor_ref']))
            same_amt = (safe_amount(ws.cell(row=row, column=col_import_amt).value)
                        == safe_amount(mr['amount']))
            if not (same_ref and same_amt):
                lost += 1
    finally:
        wb.close()

    if lost:
        raise ValueError(_MASTER_LOST_ROWS_ERROR.format(
            lost=lost, total=len(written), name=Path(master_path).name))


# Status lifecycle of a row in the Small Parcels log:
#   Pending         — the carrier file's data has been logged, nothing billed yet
#   Bill Generated  — the Acumatica import file for this row has been produced
#   Imported        — set by the AP team once the file is actually imported
# 'Pending' is deliberately kept as the first state: it is the value the bill
# pipeline filters on (analyze_carrier_file, process_freight_bills) and the value
# the team already uses, so nothing existing changes meaning.
STATUS_PENDING = 'Pending'
STATUS_BILL_GENERATED = 'Bill Generated'
STATUS_IMPORTED = 'Imported'

SMALL_PARCEL_TABS = ('Fedex', 'WWEX')
_SP_TAB_REQUIRED_HEADERS = ('doc date', 'pro number / vendor ref', 'import amount',
                            'po / sos', 'notes')


def validate_small_parcel_master(master_path):
    """Returns (is_valid, error) for a Small Parcels master workbook — used both
    for the shared copy and for one uploaded by hand."""
    try:
        wb = load_workbook(str(master_path), read_only=True)
    except PermissionError:
        return False, _MASTER_OPEN_ERROR
    except Exception as e:
        return False, f'Unable to read workbook — {e}'

    try:
        by_name = {s.strip().lower(): s for s in wb.sheetnames}
        for tab in SMALL_PARCEL_TABS:
            actual = by_name.get(tab.lower())
            if actual is None:
                return False, (f'Wrong workbook — no "{tab}" tab found. '
                               'Upload the Carrier Import File - Small Parcels.')
            headers = {_norm_header(c.value) for c in next(wb[actual].iter_rows(max_row=1))}
            missing = [h for h in _SP_TAB_REQUIRED_HEADERS if h not in headers]
            if missing:
                return False, (f'The "{tab}" tab is missing expected columns: '
                               f'{", ".join(missing)}.')
    finally:
        wb.close()
    return True, None


def append_rows_to_carrier_tab(master_path, sheet_name, mapped_rows):
    """Appends mapped rows to a flat carrier tab of the Small Parcels master
    workbook (Doc Date | Pro number / Vendor Ref | Import Amount | PO / SOS |
    Notes, plus an optional Status column), skipping rows already present.

    Each mapped row is {'doc_date', 'pro_vendor_ref', 'amount', 'po_sos',
    'notes'}; extra keys are ignored, so callers can carry their own flags.
    Returns {'added', 'skipped_duplicates', 'written'} where 'written' holds the
    mapped rows actually appended — callers count their own categories from it
    instead of re-running the dedupe.

    Shared by the WWEX raw export and the FedEx invoice PDF ingests; both tabs
    have identical shape, only the mapping upstream differs."""
    try:
        wb = load_workbook(str(master_path))
    except PermissionError:
        raise ValueError(_MASTER_OPEN_ERROR)

    ws = next((wb[s] for s in wb.sheetnames if s.strip().lower() == sheet_name.lower()), None)
    if ws is None:
        raise ValueError(f'Sheet "{sheet_name}" not found in master workbook.')

    header_idx = {_norm_header(c.value): c.column for c in ws[1] if c.value}
    required_headers = ['doc date', 'pro number / vendor ref', 'import amount',
                        'po / sos', 'notes']
    missing_headers = [h for h in required_headers if h not in header_idx]
    if missing_headers:
        raise ValueError(f'Master "{sheet_name}" tab missing expected columns: '
                         f'{", ".join(missing_headers)}.')

    col_doc_date   = header_idx['doc date']
    col_pro_ref    = header_idx['pro number / vendor ref']
    col_import_amt = header_idx['import amount']
    col_po_sos     = header_idx['po / sos']
    col_notes      = header_idx['notes']
    # Optional on the flat tab — same keyword match the bill pipeline uses, so
    # rows are written as Pending whenever the column exists and are simply
    # left unmarked (and therefore all-pending by default) when it doesn't.
    col_status = next((idx for h, idx in header_idx.items()
                       if any(kw in h for kw in ('status', 'pending', 'imported'))), None)

    # Template row to clone formatting from (number format, font, alignment,
    # fill, border) — appended rows must look identical to existing ones, not
    # fall back to openpyxl's blank "General" default styling. Read the most
    # recent existing Pending row rather than assuming row 2, so appended rows
    # inherit whatever styling the tab currently gives a Pending row. No fill
    # is applied on top of that.
    template_row = None
    if col_status:
        for r in range(ws.max_row, 1, -1):
            if safe_str(ws.cell(row=r, column=col_status).value).lower() == 'pending':
                template_row = r
                break
    if template_row is None and ws.max_row >= 2:
        template_row = ws.max_row
    max_col = ws.max_column

    # Counted, not a plain set: a single invoice can legitimately bill the same
    # tracking number twice for the same amount (a re-delivery, a duplicated
    # surcharge). Comparing counts lets the second copy through while still
    # skipping everything when the whole file is re-uploaded.
    existing_counts = Counter(
        _carrier_tab_row_key(
            safe_str(ws.cell(row=r, column=col_pro_ref).value),
            safe_str(ws.cell(row=r, column=col_po_sos).value),
            ws.cell(row=r, column=col_doc_date).value,
            safe_amount(ws.cell(row=r, column=col_import_amt).value),
        )
        for r in range(2, ws.max_row + 1)
    )

    seen     = Counter()
    written  = []
    skipped  = 0
    next_row = ws.max_row + 1

    for mr in mapped_rows:
        key = _carrier_tab_row_key(mr['pro_vendor_ref'], mr['po_sos'],
                                   mr['doc_date'], mr['amount'])
        seen[key] += 1
        if seen[key] <= existing_counts[key]:
            skipped += 1
            continue

        if template_row:
            for c in range(1, max_col + 1):
                src = ws.cell(row=template_row, column=c)
                dst = ws.cell(row=next_row, column=c)
                dst.number_format = src.number_format
                dst.font          = copy(src.font)
                dst.alignment     = copy(src.alignment)
                dst.fill          = copy(src.fill)
                dst.border        = copy(src.border)

        ws.cell(row=next_row, column=col_doc_date,   value=mr['doc_date'])
        ws.cell(row=next_row, column=col_pro_ref,    value=mr['pro_vendor_ref'])
        ws.cell(row=next_row, column=col_import_amt, value=mr['amount'])
        ws.cell(row=next_row, column=col_po_sos,     value=mr['po_sos'])
        ws.cell(row=next_row, column=col_notes,      value=mr['notes'])
        if col_status:
            ws.cell(row=next_row, column=col_status, value=STATUS_PENDING)

        # Kept so the caller can flip these rows to 'Bill Generated' once the
        # Acumatica import file for them exists.
        mr['_row'] = next_row
        next_row += 1
        written.append(mr)

    if written:
        _extend_table_ranges(ws, next_row - 1)
        try:
            wb.save(str(master_path))
        except PermissionError:
            raise ValueError(_MASTER_OPEN_ERROR)
        wb.close()
        # Raises rather than returning, so a workbook that threw the rows away
        # cannot be reported as a successful import.
        _verify_carrier_rows_persisted(master_path, sheet_name, written)

    return {'added': len(written), 'skipped_duplicates': skipped, 'written': written}


def set_carrier_tab_status(master_path, sheet_name, excel_rows, status):
    """Sets Status on specific rows of a flat carrier tab. Used to move rows from
    'Pending' to 'Bill Generated' once their Acumatica import file exists.
    Returns the number of rows updated."""
    rows = [r for r in excel_rows if r]
    if not rows:
        return 0

    try:
        wb = load_workbook(str(master_path))
    except PermissionError:
        raise ValueError(_MASTER_OPEN_ERROR)

    ws = next((wb[s] for s in wb.sheetnames if s.strip().lower() == sheet_name.lower()), None)
    if ws is None:
        raise ValueError(f'Sheet "{sheet_name}" not found in master workbook.')

    header_idx = {_norm_header(c.value): c.column for c in ws[1] if c.value}
    col_status = next((idx for h, idx in header_idx.items()
                       if any(kw in h for kw in ('status', 'pending', 'imported'))), None)
    if not col_status:
        wb.close()
        return 0

    for r in rows:
        if 2 <= r <= ws.max_row:
            ws.cell(row=r, column=col_status, value=status)

    try:
        wb.save(str(master_path))
    except PermissionError:
        raise ValueError(_MASTER_OPEN_ERROR)
    return len(rows)


def append_wwex_raw_to_master(raw_path, master_path, sheet_name='WWEX'):
    """Maps a raw WWEX carrier export into the Small Parcels master workbook's
    flat WWEX tab and appends the rows that aren't there yet.
    Returns {'added', 'skipped_duplicates', 'fee_rows_added'}."""
    if not SMALL_PARCEL_ENABLED:
        raise ValueError('Small Parcels ingestion is currently disabled.')

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

    stamp = _upload_stamp()

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
        # Tracking number (Vendor Reference 2, same value as Airbill #) — it is
        # what the reconciliation cascade in process_freight_bills matches
        # against Acumatica's shipment tracking numbers. Fee lines carry no
        # tracking number of their own, so they fall back to the invoice #.
        tracking = safe_str(row.get(col_vendorref, '')) or invoicenum
        # Customer PO, blank on plenty of rows.
        po_sos = safe_str(row.get(col_billref1, '')) if col_billref1 else ''

        # The fee prefix must stay at the head of the note so the bill pipeline
        # still recognises the line; the upload stamp goes after the separator.
        note = f'{WWEX_FEE_NOTE_PREFIX}{fee_desc}{NOTE_SEPARATOR}' if is_fee else ''

        mapped_rows.append({
            # Every row of an invoice shares the invoice date, exactly like the
            # FedEx ingest — the per-shipment ship date is not what AP posts on.
            'doc_date':       safe_date(row.get(col_invdate)),
            'pro_vendor_ref': tracking,
            'po_sos':         po_sos,
            'amount':         safe_amount(row.get(col_amount, 0)),
            'notes':          f'{note}{stamp}',
            'is_fee':         is_fee,
            # Stated on the row rather than inferred downstream, so each
            # carrier's fees keep their own routing.
            'fee_subaccount': WWEX_FEE_CONFIG['subaccount'] if is_fee else '',
        })

    result = append_rows_to_carrier_tab(master_path, sheet_name, mapped_rows)
    return {
        'added':              result['added'],
        'skipped_duplicates': result['skipped_duplicates'],
        'fee_rows_added':     sum(1 for mr in result['written'] if mr['is_fee']),
        'invoice_number':     safe_str(df[col_invoicenum].iloc[0]) if len(df) else '',
        # Consumed by the caller to build the Acumatica import file; stripped
        # before the stats are sent to the browser.
        '_written':           result['written'],
    }


# ==============================================================================
# SMALL PARCELS — ACUMATICA IMPORT FILE
# ==============================================================================
# The master workbook is the AP team's log. This is the companion file they
# actually import into Acumatica: one per uploaded carrier file, holding only the
# rows that were newly appended to the log.

SP_IMPORT_COLUMNS = [
    'Branch', 'Inventory ID', 'Transaction Descr.', 'Quantity', 'UOM',
    'Unit Cost', 'Ext. Cost', 'Discount Amount', 'Amount', 'Account',
    'Description', 'Subaccount', 'Tax Category', 'PO Number', 'PO Receipt Nbr.',
]
SP_IMPORT_BRANCH = 'MAIN'
SP_IMPORT_QUANTITY = 1
SP_IMPORT_ACCOUNT = '410160'
SP_IMPORT_DESCRIPTION = 'COGS - PACKAGING & FREIGHT: Outbound Shipping Small Parcel'
SP_IMPORT_DESCR_PREFIX = 'Bill uploaded from distribution file - '


def _sp_import_subaccount(row, lookups, pacejet_bridge=None):
    """Subaccount for one small-parcel row. Carrier fee lines carry the
    subaccount they post to, since each carrier routes its own fees; everything
    else resolves the customer against Acumatica shipments and is left blank
    when nothing matches.

    The tracking number is tried first because it is the carrier's own key and
    is unique; the PO comes second, since a PO only identifies a shipment
    together with its customer. Rows the carrier billed without any reference
    Acumatica knows fall through to PaceJet, which rated the shipment and can
    name the Acumatica shipment behind the tracking number, and finally to the
    sender company printed on the invoice."""
    if row.get('is_fee'):
        return row.get('fee_subaccount') or WWEX_FEE_CONFIG['subaccount'], 'Fee'

    ordernbr_lookup, shipmentnbr_lookup, tracking_lookup, custorder_lookup = lookups
    tracking = row.get('pro_vendor_ref')
    po_sos   = row.get('po_sos')

    match = tracking_lookup.get(tracking) if tracking else None
    if not match:
        for key in (po_sos, tracking):
            if not key:
                continue
            match = (ordernbr_lookup.get(key)
                     or shipmentnbr_lookup.get(key)
                     or custorder_lookup.get(key))
            if match:
                break
    if match:
        return SUBACCOUNTS.get(match['Customer'], DEFAULT_SUBACCOUNT), 'Direct'

    if pacejet_bridge and tracking:
        control = pacejet_bridge.get(tracking)
        if control:
            match = shipmentnbr_lookup.get(control) or ordernbr_lookup.get(control)
            if match:
                return SUBACCOUNTS.get(match['Customer'], DEFAULT_SUBACCOUNT), 'Via PaceJet'

    # Nothing referenced the shipment, so the sender company printed on the
    # invoice is the only identification left. Deliberately last: it names the
    # customer, never the individual shipment, so any real reference wins over
    # it. Only senders listed in FEDEX_SENDER_SUBACCOUNTS qualify.
    subaccount = FEDEX_SENDER_SUBACCOUNTS.get(row.get('sender'))
    if subaccount:
        return subaccount, 'Via Sender'

    return '', 'Review'


def build_small_parcel_import(rows, shipments_df, output_path, pacejet_df=None):
    """Writes the Acumatica import file for the rows just appended to the log.
    Returns {'filename', 'rows', 'matched', 'unmatched', 'via_pacejet',
    'via_sender', 'total'}."""
    lookups        = build_shipments_lookups(shipments_df)
    pacejet_bridge = build_pacejet_parcel_bridge(pacejet_df)

    records     = []
    cases       = []
    matched     = 0
    via_pacejet = 0
    via_sender  = 0
    for row in rows:
        subaccount, case = _sp_import_subaccount(row, lookups, pacejet_bridge)
        cases.append(case)
        if case != 'Review':
            matched += 1
        if case == 'Via PaceJet':
            via_pacejet += 1
        if case == 'Via Sender':
            via_sender += 1
        amount = row['amount']
        records.append({
            'Branch':             SP_IMPORT_BRANCH,
            'Inventory ID':       '',
            'Transaction Descr.': f'{SP_IMPORT_DESCR_PREFIX}{row["pro_vendor_ref"]}',
            'Quantity':           SP_IMPORT_QUANTITY,
            'UOM':                '',
            'Unit Cost':          amount,
            'Ext. Cost':          amount,
            'Discount Amount':    '',
            'Amount':             amount,
            'Account':            SP_IMPORT_ACCOUNT,
            'Description':        SP_IMPORT_DESCRIPTION,
            'Subaccount':         subaccount,
            'Tax Category':       '',
            'PO Number':          '',
            'PO Receipt Nbr.':    '',
        })

    export_df = pd.DataFrame(records, columns=SP_IMPORT_COLUMNS)
    # Rows with no Subaccount are highlighted the same way process_freight_bills
    # flags its Review rows, so they are obvious before the import. Rows whose
    # Subaccount came from the sender company are greyed instead: they are
    # complete, but the customer was identified by who shipped the package
    # rather than by a reference, so the AP team can still spot-check them.
    _write_formatted_excel(output_path, export_df,
                           [not r['Subaccount'] for r in records], [],
                           notice_flags=[c == 'Via Sender' for c in cases])
    return {
        'filename':    Path(output_path).name,
        'rows':        len(records),
        'matched':     matched,
        'unmatched':   len(records) - matched,
        'via_pacejet': via_pacejet,
        'via_sender':  via_sender,
        'total':       round(sum(r['amount'] for r in rows), 2),
    }


# ==============================================================================
# FEDEX INVOICE PDF INGESTION
# ==============================================================================
# The invoice PDF is parsed directly — no intermediate spreadsheet. Text is
# extracted per page and matched with label-anchored regexes rather than by line
# position: the extractor glues adjacent tokens together, so labels ("Ship
# Date:", "Total Charge", "P.O.#:") are the only stable anchors.

_FEDEX_AMOUNT = r'-?\$?[\d,]+\.\d\d'
_FEDEX_INVOICE_NUM_RE = re.compile(r'Invoice Number\s*(\S+?)\s*Account Number')
_FEDEX_INVOICE_DATE_RE = re.compile(r'Invoice Date\s*([A-Z][a-z]{2} \d{1,2}, \d{4})')
_FEDEX_INVOICE_TOTAL_RE = re.compile(r'TOTAL THIS INVOICE\s*USD\s*(' + _FEDEX_AMOUNT + ')')
_FEDEX_TOTAL_CHARGE_RE = re.compile(r'Total Charge\s*USD\s*(' + _FEDEX_AMOUNT + ')')
# Ground tracking numbers are 12 digits and are the first value of the row that
# follows the "Rated Weight" column header. Both ends of that number can be glued
# to neighbouring values: to the service type on the left-hand side
# ("872388092075Direct Sign") and, on shipments printed without a service type,
# to the zone and weights on the right ("7924523186498123.6 lbs"). So the match
# is anchored on the header and must not require a boundary after the 12th digit.
_FEDEX_TRACKING_HEADER = 'Rated Weight'
_FEDEX_TRACKING_RE = re.compile(r'(?<!\d)(\d{12})')
# The PO sits between the "P.O.#:" label and the "Tracking ID" column header,
# followed by whichever boilerplate footnotes apply to the shipment.
_FEDEX_PO_SEGMENT_RE = re.compile(r'P\.O\.#:(.*?)Tracking ID', re.DOTALL)
# Footnote sentences FedEx prints in the same area. The PO, when present, always
# precedes them, so the segment is cut at the earliest marker found.
_FEDEX_NOTE_MARKERS = (
    'The Earned Discount', 'We calculated', 'This shipment', 'Additional Handling',
    'Minimum Billable', 'Dimensions -', 'Residential', 'Address Correction',
)
# Not every invoice prints a "P.O.#:" label. Some print the same references in
# the FedEx reference fields instead, glued into one run of text
# ("Ref.#3: SOS140604Ref.#2: SOS129956"), and the shipment that gave the AP team
# a blank Subaccount had its sales order printed there all along. Read in this
# order: the sales order Acumatica knows comes first, the customer PO second,
# the free-text customer reference last.
_FEDEX_REF_FIELDS = ('Ref.#3', 'Ref.#2', 'Cust. Ref.')
_FEDEX_REF_SEGMENT_RES = {
    label: re.compile(
        re.escape(label) + r'\s*:\s*(.*?)(?=Cust\. Ref\.\s*:|Ref\.#[123]\s*:|'
        r'Dept\.#\s*:|P\.O\.#:|Tracking ID|$)', re.DOTALL)
    for label in _FEDEX_REF_FIELDS
}
# A reference field runs straight into whatever boilerplate the shipment carries,
# and that boilerplate holds numbers of its own (a revenue threshold, a zone), so
# the segment is cut at the first marker before anything is read out of it.
_FEDEX_REF_CUT_MARKERS = _FEDEX_NOTE_MARKERS + (
    'Payor:', 'NO REFERENCE INFORMATION', 'Distance Based', 'The delivery commitment',
    'The Earned Discount', 'Fuel Surcharge', 'The required information',
    'FedEx has audited', 'Package Delivered', 'Automation', 'Svc Area',
    'Your package', 'Please be sure', 'FedEx Use', 'Sender ', 'Recipient ',
)
# Only the leading tokens are considered: a reference sits at the start of its
# field, so a number found further in is boilerplate that survived the cut.
_FEDEX_REF_MAX_TOKENS = 3
# A reference token is one unbroken run — no spaces, since the multi-word values
# in these fields are project names and contact names, never a reference
# Acumatica can match. Upper case and digits only, at least two digits: that is
# how every reference on these invoices is printed (SOS140604, 410467754,
# KIS12034AB6856F05), and it is what keeps a contact name out of the field —
# "JS-Joann Montalvo2" carries a digit too.
_FEDEX_REF_TOKEN_RE = re.compile(r'^(?=(?:\D*\d){2})[A-Z0-9][A-Z0-9\-/.]{4,}$')

# A PO carries a digit, a separator, or is all caps. Requiring that rejects the
# stray plain words left over when unknown boilerplate is not cut, so an
# unrecognised footnote leaves PO / SOS blank instead of writing garbage.
_FEDEX_PO_SHAPE_RE = re.compile(r'^(?=.*(?:\d|[-/]|[A-Z]{2}))[A-Za-z0-9][A-Za-z0-9\-/. ]*$')
# An aggregated (non-itemised) charge class is reported as a small table of its
# own instead of as shipments. It is recognised by that table's column header
# rather than by the class name: an aggregated table counts "Packages", an
# itemised one counts "Shipments". Keying on the header is what survives FedEx
# renaming the class — "FedEx SmartPost" became "FedEx Ground Economy", and the
# first invoice printed with the new name reconciled short and imported nothing.
_FEDEX_AGGREGATE_SECTION_RE = re.compile(
    r'FedEx [A-Za-z ]{1,30}?Shipments \((?:Original|Rebill)\)\s*DatePackages(.*?)'
    r'(?=FedEx [A-Za-z ]{1,30}?Shipments \(|Total FedEx|TOTAL THIS INVOICE|$)', re.DOTALL)
# One section can hold more than one class (1-70lbs and over-70lbs), each closed
# by its own "<class name> Subtotal$211.72".
_FEDEX_AGGREGATE_CLASS_RE = re.compile(r'Subtotal\s*\$(' + _FEDEX_AMOUNT + ')')
# Every row of a class reads "MM/DD<packages> <weight>". The count is summed
# across all of them: a class routinely spans several ship dates, and reading
# only the first row is what made the row label undercount the packages.
_FEDEX_AGGREGATE_PACKAGES_RE = re.compile(r'\d{2}/\d{2}(\d+)\s')
# Every charge class closes its shipment table with "<class name> Subtotal$1.23",
# and those subtotals add up to the printed invoice total. They are not used to
# build rows — they are the invoice's own breakdown, read only to name the class
# behind a residual. Without it a mismatch reports a bare dollar amount and the
# class has to be found by hand in the PDF.
_FEDEX_CLASS_SUBTOTAL_RE = re.compile(
    r'([A-Za-z][A-Za-z0-9 \-.>/]{2,40}?)\s*Subtotal\s*\$(' + _FEDEX_AMOUNT + ')')
# The first class of a "(Rebill)" table is printed glued to its own data row
# ("aid05/051 39.00 32.26 32.26Ground-Prepaid"), so the label is cut after the
# last amount that ran into it.
_FEDEX_SUBTOTAL_LABEL_NOISE_RE = re.compile(r'^.*\d\.\d\d')
# A breakdown longer than this is listed truncated rather than filling the
# browser with the whole invoice.
_FEDEX_MAX_LISTED_CLASSES = 12

FEDEX_AGGREGATE_LABEL = '(not itemized - {n} package{s} aggregated)'
# The page-1 Invoice Summary declares every charge class on the invoice, only
# one of which ("FedEx Express Services") is itemised as shipments further down.
# The rest are account-level charges that belong to no shipment at all, and
# leaving them unparsed is what made a whole invoice fail to reconcile and
# import nothing. Matched by their printed label, one pattern each: the summary
# prints some classes with a "Total Charges" column and some without.
_FEDEX_SUMMARY_RE = re.compile(r'Invoice Summary(.*?)TOTAL THIS INVOICE', re.DOTALL)
FEDEX_SUMMARY_FEE_PATTERNS = (
    # Customer-level fees — scheduled pickup and the like.
    ('FedEx Other Charges',
     re.compile(r'FedEx Other Charges\s*Total Charges\s*USD\s*(' + _FEDEX_AMOUNT + ')')),
    # Finance charges — late fees on previously issued invoices. The lookbehind
    # keeps this pattern off the "FedEx Other Charges" label above.
    ('Other Charges',
     re.compile(r'(?<!FedEx )Other Charges\s*USD\s*(' + _FEDEX_AMOUNT + ')')),
)
FEDEX_FEE_LABEL = '(not itemized - {label})'
# The sender block of a shipment, between the two address labels. FedEx prints
# the contact name first and the company on the next line, both glued into one
# run of text by the extractor, so the whole segment is kept and searched by
# keyword rather than split into fields.
_FEDEX_SENDER_RE = re.compile(r'Sender(.*?)Recipient', re.DOTALL)
# Last-resort customer identification, used only once the tracking number, the
# PO and the PaceJet bridge have all failed to name an Acumatica shipment.
# A return shipped from a retailer's own store carries no reference Acumatica
# knows, but the retailer's name is printed on the sender address — enough to
# post the charge to that customer instead of leaving it for manual review.
# Keyed by a company name as printed on the invoice. Names that belong to an
# Acumatica customer read their subaccount out of SUBACCOUNTS rather than
# repeating the string; a sender that is a vendor rather than a customer (an
# inbound purchase, not a return) has no customer id and names its subaccount
# directly. Matching is a substring test on the sender address, so a more
# specific name must come before a shorter one it contains.
FEDEX_SENDER_SUBACCOUNTS = {
    'LOWES':                SUBACCOUNTS['C100002'],
    'NUWHIRL SYSTEMS CORP': 'PUR-000000-000000-000000',
}


def _pdf_lines(pdf_path):
    """Full text of a PDF as one normalised string per page, in document order.
    Sole point of contact with the PDF library — swapping libraries only means
    rewriting this function."""
    from pypdf import PdfReader

    pages = []
    reader = PdfReader(str(pdf_path))
    for page in reader.pages:
        text = page.extract_text() or ''
        # \x7f is FedEx's footnote marker glyph; nbsp shows up around amounts.
        text = text.replace('\x7f', ' ').replace('\xa0', ' ')
        pages.append(' '.join(text.split()))
    return pages


def _fedex_amount(raw):
    """FedEx prints credits as -$12.34 and occasionally as ($12.34)."""
    raw = (raw or '').strip()
    if raw.startswith('(') and raw.endswith(')'):
        return -safe_amount(raw[1:-1])
    return safe_amount(raw)


def _fedex_tracking(head):
    """Tracking number of one shipment block. Preferred match is the first
    12-digit run after the column header; some shipments repeat the header for
    their dimension footnotes, so a whole-block search is the fallback."""
    anchor = head.find(_FEDEX_TRACKING_HEADER)
    if anchor >= 0:
        m = _FEDEX_TRACKING_RE.search(head[anchor:])
        if m:
            return m.group(1)
    m = _FEDEX_TRACKING_RE.search(head)
    return m.group(1) if m else ''


def _fedex_sender(block):
    """Company on the sender address of one shipment block, as one of the keys
    of FEDEX_SENDER_SUBACCOUNTS, or '' when none of them is printed there.

    Only the sender segment is searched: the same name on the recipient side
    means the opposite direction and must not identify the customer."""
    m = _FEDEX_SENDER_RE.search(block)
    if not m:
        return ''
    segment = m.group(1).upper()
    return next((name for name in FEDEX_SENDER_SUBACCOUNTS if name in segment), '')


def _fedex_ref_token(seg):
    """First reference-shaped token of a FedEx reference field, or '' when the
    field holds nothing that could be a reference."""
    cuts = [seg.find(marker) for marker in _FEDEX_REF_CUT_MARKERS]
    cuts = [c for c in cuts if c >= 0]
    if cuts:
        seg = seg[:min(cuts)]

    for token in seg.split()[:_FEDEX_REF_MAX_TOKENS]:
        if (token not in FEDEX_PO_NOISE_TOKENS
                and len(token) <= FEDEX_PO_MAX_LENGTH
                and _FEDEX_REF_TOKEN_RE.match(token)):
            return token
    return ''


def _fedex_po(block):
    """PO for one shipment block, or '' when the invoice printed no reference
    this tool can use.

    The "P.O.#:" label wins whenever the invoice prints it, whole segment and
    all: a PO there can be several words ("Bldg R: Door H/W") and is the field
    FedEx reserves for it. The reference fields are read only when that label is
    absent or holds nothing usable, and only one token deep, since they are
    shared with free text."""
    m = _FEDEX_PO_SEGMENT_RE.search(block)
    if m:
        seg = m.group(1)
        cuts = [seg.find(marker) for marker in _FEDEX_NOTE_MARKERS]
        cuts = [c for c in cuts if c >= 0]
        if cuts:
            seg = seg[:min(cuts)]

        po = ' '.join(seg.split())
        if (po not in FEDEX_PO_NOISE_TOKENS
                and FEDEX_PO_MIN_LENGTH <= len(po) <= FEDEX_PO_MAX_LENGTH
                and _FEDEX_PO_SHAPE_RE.match(po)):
            return po
        # The label was printed with the PO buried in text the markers above do
        # not cut ("RETURN PO 406617968", "408245429 Unauthorized"), so the same
        # token scan the reference fields use is the second chance.
        po = _fedex_ref_token(seg)
        if po:
            return po

    for label in _FEDEX_REF_FIELDS:
        rm = _FEDEX_REF_SEGMENT_RES[label].search(block)
        if not rm:
            continue
        po = _fedex_ref_token(rm.group(1))
        if po:
            return po
    return ''


def _fedex_class_subtotals(text):
    """[(class name, amount)] for every charge class the invoice closes with its
    own subtotal, in printed order. This is the invoice's own breakdown of the
    total, so it names the classes a parse can be missing."""
    classes = []
    for label, amount in _FEDEX_CLASS_SUBTOTAL_RE.findall(text):
        label = _FEDEX_SUBTOTAL_LABEL_NOISE_RE.sub('', label).strip()
        if label:
            classes.append((label, _fedex_amount(amount)))
    return classes


def _fedex_residual_diagnosis(text, residual):
    """One sentence naming the charges behind a non-zero residual, for the
    message the AP team reads. A class whose own subtotal equals the residual is
    the class that went unparsed — the usual cause, and the one FedEx renaming
    "FedEx SmartPost" to "FedEx Ground Economy" produced. When no single class
    accounts for it, the whole breakdown is listed so the invoice does not have
    to be opened by hand."""
    classes = _fedex_class_subtotals(text)
    if not classes:
        return ''

    target = round(abs(residual), 2)
    named = [f'"{label}" (${amount:,.2f})' for label, amount in classes
             if abs(amount - target) <= FEDEX_TOTAL_TOLERANCE]
    if named:
        return (f'Unparsed charge class: {", ".join(named)} — the invoice prints '
                f'it under a name this parser does not recognise.')

    listed = '; '.join(f'{label} ${amount:,.2f}'
                       for label, amount in classes[:_FEDEX_MAX_LISTED_CLASSES])
    if len(classes) > _FEDEX_MAX_LISTED_CLASSES:
        listed += f'; and {len(classes) - _FEDEX_MAX_LISTED_CLASSES} more'
    return f'Charge classes printed on this invoice: {listed}.'


def parse_fedex_pdf(pdf_path):
    """Parses a FedEx invoice PDF into rows for the flat Fedex tab.

    Returns {'invoice_number', 'invoice_date', 'rows', 'detail_count',
    'aggregate_count', 'detail_total', 'invoice_total', 'residual'} where
    'residual' is invoice_total - (detail + aggregate) — non-zero means the
    invoice holds charges this parser did not account for, which the caller
    surfaces instead of silently under-importing."""
    pages = _pdf_lines(pdf_path)
    text = ' '.join(pages)

    m = _FEDEX_INVOICE_NUM_RE.search(text)
    invoice_number = m.group(1) if m else ''
    m = _FEDEX_INVOICE_DATE_RE.search(text)
    invoice_date = safe_date(m.group(1)) if m else None
    m = _FEDEX_INVOICE_TOTAL_RE.search(text)
    invoice_total = _fedex_amount(m.group(1)) if m else 0.0

    note_head = f'FedEx Inv {invoice_number}' if invoice_number else 'FedEx Inv'
    note = f'{note_head}{NOTE_SEPARATOR}{_upload_stamp()}'

    rows = []
    # Each itemised shipment is a block that starts at a "Ship Date:" label and
    # runs until the next one.
    blocks = text.split('Ship Date:')[1:]
    for block in blocks:
        m = _FEDEX_TOTAL_CHARGE_RE.search(block)
        if not m:
            continue
        rows.append({
            'doc_date':       invoice_date,
            'pro_vendor_ref': _fedex_tracking(block[:m.start()]),
            'po_sos':         _fedex_po(block),
            'amount':         _fedex_amount(m.group(1)),
            'notes':          note,
            'is_aggregate':   False,
            # Not written to the log — carried through to the Acumatica import
            # file, where it is the last chance to name the customer.
            'sender':         _fedex_sender(block),
        })
    detail_count = len(rows)
    detail_total = round(sum(r['amount'] for r in rows), 2)

    # Aggregated classes (FedEx Ground Economy, formerly SmartPost) carry no
    # per-shipment detail — the invoice only reports a package count and a class
    # subtotal.
    aggregate_total = 0.0
    for section in _FEDEX_AGGREGATE_SECTION_RE.findall(text):
        cursor = 0
        for m in _FEDEX_AGGREGATE_CLASS_RE.finditer(section):
            amount = _fedex_amount(m.group(1))
            # This class owns every row since the previous class in the same
            # section closed.
            packages = sum(int(n) for n in _FEDEX_AGGREGATE_PACKAGES_RE.findall(
                section[cursor:m.start()]))
            cursor = m.end()
            aggregate_total += amount
            rows.append({
                'doc_date':       invoice_date,
                'pro_vendor_ref': FEDEX_AGGREGATE_LABEL.format(
                    n=packages, s='' if packages == 1 else 's'),
                'po_sos':         '',
                'amount':         amount,
                'notes':          note,
                'is_aggregate':   True,
                # An aggregated class prints no addresses at all.
                'sender':         '',
            })

    aggregate_count = len(rows) - detail_count

    # Account-level charges declared in the page-1 Invoice Summary. Read from
    # the summary rather than from their own detail pages: the summary is the
    # invoice's own statement of what it is charging for, and it is the block
    # TOTAL THIS INVOICE itself comes from, so the two always agree.
    fee_total = 0.0
    m = _FEDEX_SUMMARY_RE.search(pages[0] if pages else '')
    summary = m.group(1) if m else ''
    for label, pattern in FEDEX_SUMMARY_FEE_PATTERNS:
        fm = pattern.search(summary)
        if not fm:
            continue
        amount = _fedex_amount(fm.group(1))
        if not amount:
            continue
        fee_total += amount
        rows.append({
            'doc_date':       invoice_date,
            'pro_vendor_ref': FEDEX_FEE_LABEL.format(label=label),
            'po_sos':         '',
            'amount':         amount,
            'notes':          note,
            'is_aggregate':   False,
            # Skips the customer cascade in _sp_import_subaccount entirely and
            # carries the subaccount it posts to, so FedEx fees stay routed by
            # FEDEX_FEE_CONFIG and not by whatever WWEX uses.
            'is_fee':         True,
            'fee_subaccount': FEDEX_FEE_CONFIG['subaccount'],
            'sender':         '',
        })

    # Still non-zero when a charge class this parser does not know about is on
    # the invoice — the caller refuses to import rather than under-bill.
    residual = round(invoice_total - detail_total - aggregate_total - fee_total, 2)

    return {
        'invoice_number':  invoice_number,
        'invoice_date':    invoice_date,
        'rows':            rows,
        'detail_count':    detail_count,
        'aggregate_count': aggregate_count,
        'fee_count':       len(rows) - detail_count - aggregate_count,
        'detail_total':    detail_total,
        'fee_total':       round(fee_total, 2),
        'invoice_total':   invoice_total,
        'residual':        residual,
        # Empty unless the invoice failed to reconcile: the sentence that names
        # what the residual is made of, read straight into the blocking message.
        'diagnosis':       (_fedex_residual_diagnosis(text, residual)
                            if abs(residual) > FEDEX_TOTAL_TOLERANCE else ''),
    }


def validate_fedex_pdf(pdf_path):
    """Returns (is_valid, error)."""
    # Checked before handing the file to the PDF library, so a mislabelled file
    # gets a clear message instead of a parser error.
    try:
        with open(str(pdf_path), 'rb') as fh:
            if fh.read(5) != b'%PDF-':
                return False, 'Wrong file — this is not a PDF.'
    except Exception as e:
        return False, f'Unable to read file — {e}'

    try:
        pages = _pdf_lines(pdf_path)
    except Exception as e:
        return False, f'Unable to read PDF — {e}'

    text = ' '.join(pages)
    if not text.strip():
        return False, ('No text found in this PDF — it looks scanned. '
                       'Download the invoice PDF from the FedEx billing portal.')
    if not _FEDEX_INVOICE_NUM_RE.search(text):
        return False, 'Wrong file — no FedEx invoice number found in this PDF.'
    if 'Ship Date:' not in text:
        return False, 'Wrong file — no shipment detail found in this FedEx invoice.'
    return True, None


def append_fedex_pdf_to_master(pdf_path, master_path, sheet_name='Fedex'):
    """Parses a FedEx invoice PDF and appends its shipments to the Small Parcels
    master workbook's flat Fedex tab.
    Returns {'added', 'skipped_duplicates', 'aggregate_rows_added',
    'fee_rows_added', 'invoice_number', 'detail_total', 'invoice_total',
    'residual'}."""
    if not SMALL_PARCEL_ENABLED:
        raise ValueError('Small Parcels ingestion is currently disabled.')

    parsed = parse_fedex_pdf(pdf_path)
    if not parsed['rows']:
        raise ValueError('No shipments found in this FedEx invoice PDF.')

    # Nothing is written unless the parsed rows add up to the invoice total the
    # PDF itself declares. A short import would silently under-pay the carrier,
    # so a mismatch is reported with both figures instead, followed by the
    # invoice's own breakdown of what the missing amount is.
    if abs(parsed['residual']) > FEDEX_TOTAL_TOLERANCE:
        raise BlockingImportError(
            f'FedEx invoice {parsed["invoice_number"]} requires manual review — '
            f'${abs(parsed["residual"]):,.2f} of the '
            f'${parsed["invoice_total"]:,.2f} total is unrecognised. '
            f'Nothing was imported.'
            + (f' {parsed["diagnosis"]}' if parsed['diagnosis'] else ''))

    result = append_rows_to_carrier_tab(master_path, sheet_name, parsed['rows'])
    return {
        'added':                result['added'],
        'skipped_duplicates':   result['skipped_duplicates'],
        'aggregate_rows_added': sum(1 for r in result['written'] if r['is_aggregate']),
        'fee_rows_added':       sum(1 for r in result['written'] if r.get('is_fee')),
        'invoice_number':       parsed['invoice_number'],
        'detail_total':         parsed['detail_total'],
        'invoice_total':        parsed['invoice_total'],
        'residual':             parsed['residual'],
        # Consumed by the caller to build the Acumatica import file; stripped
        # before the stats are sent to the browser.
        '_written':             result['written'],
    }


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
        if is_detail_sheet(sheet_name):
            continue
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


def is_small_parcel_master(sheet_names):
    """True for the Small Parcels master workbook. It has to be told apart from
    the LTL Carrier Import File before the carrier check below: its WWEX tab
    matches CARRIER_VENDOR_IDS, so the LTL branch would otherwise claim it."""
    normalized = {_norm_header(s) for s in sheet_names}
    return ({'fedex', 'wwex'} <= normalized
            and not any(n.startswith('master') for n in normalized))


def classify_file(file_path, filename):
    """Auto-detect file type from content structure.
    Returns ('carrier' | 'priority1' | 'pacejet' | 'wwex_raw' | 'fedex_pdf' |
    'sp_master' | 'unknown', error_or_None).
    """
    name_lc = (filename or '').lower()

    # FedEx invoice PDF — handled first, pandas cannot open a PDF at all.
    if name_lc.endswith('.pdf'):
        if not SMALL_PARCEL_ENABLED:
            return 'unknown', 'PDF files are not accepted.'
        ok, err = validate_fedex_pdf(file_path)
        return ('fedex_pdf', None) if ok else ('unknown', err)

    is_csv = name_lc.endswith('.csv')
    try:
        if not is_csv:
            try:
                with pd.ExcelFile(_to_bytes(file_path)) as xl:
                    sheet_names = xl.sheet_names
                    # Small Parcels master, uploaded by hand when the shared
                    # folder is unavailable.
                    if SMALL_PARCEL_ENABLED and is_small_parcel_master(sheet_names):
                        return 'sp_master', None
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
        # misfire those checks. Disabled via SMALL_PARCEL_ENABLED.
        if SMALL_PARCEL_ENABLED and {'airbill #', 'scac', 'charge type 1'} <= set(cols_lower):
            return 'wwex_raw', None

        # PaceJet: many Shipment* columns
        if (sum(1 for c in cols if c.startswith('Shipment')) > 5
                or any('shipmentuserfield3' in c for c in cols_lower)):
            return 'pacejet', None

        # Priority 1: Invoice Number column
        if any('invoice number' in c or 'invoice #' in c for c in cols_lower):
            return 'priority1', None

        return 'unknown', ('Unrecognized file — expected a Carrier Import, '
                           'Priority 1 Pending Invoices, PaceJet Export, '
                           'WWEX raw export, or FedEx invoice PDF.')
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
    (ordernbr_lookup, shipmentnbr_lookup,
     tracking_lookup, custorder_lookup) = build_shipments_lookups(shipments_df)

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
        if is_detail_sheet(sheet_name):
            continue
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
                # Notes hold "AUTO-FEE:<description> | <upload stamp>" — the
                # stamp must not leak into the bill's line description.
                line_desc = (wwex_fee_notes[len(WWEX_FEE_NOTE_PREFIX):]
                             .split(NOTE_SEPARATOR)[0].strip().title())

            # ------------------------------------------------------------------
            # ALL OTHER CARRIERS — generic cascade
            # ------------------------------------------------------------------
            else:
                # The PRO / tracking number is the carrier's own unique key, so
                # it decides before the PO, which is only unique per customer.
                match = tracking_lookup.get(pronumber) if pronumber else None
                if not match:
                    for key in [po_sos, pronumber]:
                        if not key:
                            continue
                        match = (ordernbr_lookup.get(key)
                                 or shipmentnbr_lookup.get(key)
                                 or custorder_lookup.get(key))
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
