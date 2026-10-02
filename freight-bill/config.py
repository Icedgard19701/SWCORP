import json
import os
from pathlib import Path

_CREDS_PATH = Path(os.environ.get(
    'SWCORP_CREDS_PATH',
    str(Path.home() / '.swcorp' / 'credentials.json')
))

def _load_creds():
    with open(_CREDS_PATH, encoding='utf-8') as f:
        return json.load(f)

def get_acumatica_config():
    return _load_creds()['acumatica']

APPLICATION_ROOT = '/freight-bill-processor'
SECRET_KEY       = _load_creds().get('secret_key', 'dev-fallback-key')

BASE_DIR   = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / 'uploads'
OUTPUT_DIR = BASE_DIR / 'output'

# Machine-specific paths — set as environment variables in IIS Application Pool.
# Fallback values are for local development only.
CARRIER_BILLS_DIR = Path(os.environ.get(
    'CARRIER_BILLS_DIR',
    r'C:\Users\MANNY.DEV1\SWCorp\AP Team - Documents\AP\Carrier Bills Import\Import Files'
))
CARRIER_SOURCE_DIR = Path(os.environ.get(
    'CARRIER_SOURCE_DIR',
    r'C:\Users\MANNY.DEV1\SWCorp\AP Team - Documents\AP\Carrier Bills Import\Carrier Import File - Iris'
))
# Small-parcel master workbook (Fedex / WWEX flat tabs + "- SP" detail tabs).
# Separate workbook and folder from the LTL one in CARRIER_SOURCE_DIR.
CARRIER_SP_SOURCE_DIR = Path(os.environ.get(
    'CARRIER_SP_SOURCE_DIR',
    r'C:\Users\MANNY.DEV1\SWCorp\AP Team - Documents\AP\Carrier Bills Import\Carrier Import File - Small Parcels'
))
CARRIER_AUDIT_DIR = Path(os.environ.get(
    'CARRIER_AUDIT_DIR',
    r'C:\Users\MANNY.DEV1\SWCorp\AP Team - Documents\AP\Carrier Bills Import\Carrier Import File - Audit'
))
CARRIER_BILLS_URL = 'https://swcorp.sharepoint.com/:f:/s/APTeam/IgBXo3fwVSwCRbUbYOGaV8lYARbA-L8OS3bbmYCcBhF7ZUc?e=foq1Gt'

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
