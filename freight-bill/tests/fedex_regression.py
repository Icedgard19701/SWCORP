"""Regression check for the FedEx invoice PDF parser.

Run it after any change to the FedEx side of processor.py:

    .venv\\Scripts\\python.exe tests\\fedex_regression.py

Two things are checked against every PDF in the corpus:

  1. The recorded parse — invoice number, totals, and the exact row counts and
     aggregate labels the parser produced when the expectations were taken. A
     changed number here is either a fix (re-record it) or a regression.
  2. Reconciliation under simulated FedEx relabelling. The parser reads the
     invoice by its printed labels, and FedEx does rename them — "FedEx
     SmartPost" became "FedEx Ground Economy" and the invoice that carried the
     new name reconciled $211.72 short and imported nothing. Each scenario
     rewrites one label and asserts the parse still refuses to under-import.

The corpus is real invoice PDFs and is not in the repo (see .gitignore). Point
FEDEX_CORPUS_DIR at them, or drop them in FEDEX_PDF/. Missing corpus is
reported as skipped, not as a pass.

    --update   re-record fedex_expected.json from the current parser
"""

import json
import os
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))

import processor as P  # noqa: E402  (needs the path above)

CORPUS_DIR    = Path(os.environ.get('FEDEX_CORPUS_DIR', BASE_DIR / 'FEDEX_PDF'))
EXPECTED_PATH = Path(__file__).resolve().parent / 'fedex_expected.json'

# One printed label rewritten per scenario, and whether the parse must refuse to
# import afterwards. A rename that moves money must always block; a rename that
# only costs attribution (the PO label) is expected to get through, and is here
# to record that gap rather than to claim it is covered.
RELABEL_SCENARIOS = [
    ('aggregate table header', 'DatePackages',       'DatePkgs',            True),
    ('class subtotal',         'Subtotal$',          'Class Total$',        True),
    ('invoice total',          'TOTAL THIS INVOICE', 'INVOICE GRAND TOTAL', True),
    ('per-shipment charge',    'Total ChargeUSD',    'Charge TotalUSD',     True),
    ('purchase order label',   'P.O.#:',             'Purch Ord#:',         False),
]


def corpus_files():
    if not CORPUS_DIR.is_dir():
        return []
    # Windows is case-insensitive, so one glob per case would return duplicates.
    return sorted({p.resolve() for p in CORPUS_DIR.iterdir()
                   if p.suffix.lower() == '.pdf'})


def fingerprint(pdf_path):
    """The parse facts worth pinning — everything a regression would move."""
    parsed = P.parse_fedex_pdf(pdf_path)
    return {
        'invoice_number':  parsed['invoice_number'],
        'invoice_date':    str(parsed['invoice_date'] or ''),
        'invoice_total':   parsed['invoice_total'],
        'detail_total':    parsed['detail_total'],
        'fee_total':       parsed['fee_total'],
        'residual':        parsed['residual'],
        'detail_count':    parsed['detail_count'],
        'aggregate_count': parsed['aggregate_count'],
        'fee_count':       parsed['fee_count'],
        'aggregate_rows':  [[r['pro_vendor_ref'], r['amount']]
                            for r in parsed['rows'] if r['is_aggregate']],
        'fee_rows':        [[r['pro_vendor_ref'], r['amount']]
                            for r in parsed['rows'] if r.get('is_fee')],
    }


def relabelled(original, old, new):
    """A _pdf_lines replacement that rewrites one printed label, so a future
    FedEx relabelling can be tested without a PDF that has it. Built from the
    real reader every time — chaining these would stack renames and stop
    testing one label at a time."""
    def read(path):
        return [page.replace(old, new) for page in original(path)]
    return read


def check_relabelling(pdf_path):
    """[failure] for every scenario whose outcome differs from the recorded
    expectation."""
    failures = []
    original = P._pdf_lines
    text     = ' '.join(original(pdf_path))
    try:
        for name, old, new, must_block in RELABEL_SCENARIOS:
            # A label this invoice never prints cannot be renamed on it. Not
            # every invoice carries every charge class.
            if old not in text:
                continue
            P._pdf_lines = relabelled(original, old, new)
            try:
                parsed  = P.parse_fedex_pdf(pdf_path)
                blocked = abs(parsed['residual']) > P.FEDEX_TOTAL_TOLERANCE
                excuse  = ''
            except Exception as e:
                # A parse that cannot run at all also refuses to under-import.
                blocked, excuse = True, f' ({type(e).__name__})'
            if blocked != must_block:
                failures.append(
                    f'{name} renamed: expected '
                    f'{"a blocked import" if must_block else "the import to proceed"}, '
                    f'got {"blocked" if blocked else "imported"}{excuse}')
    finally:
        P._pdf_lines = original
    return failures


def main():
    update = '--update' in sys.argv
    files  = corpus_files()
    if not files:
        print(f'SKIPPED — no PDFs in {CORPUS_DIR}')
        print('Set FEDEX_CORPUS_DIR to the invoice folder, or drop PDFs in FEDEX_PDF/.')
        return 0

    expected = {}
    if EXPECTED_PATH.exists():
        expected = json.loads(EXPECTED_PATH.read_text(encoding='utf-8'))

    recorded, failures, checked = {}, [], 0
    for pdf in files:
        actual = fingerprint(pdf)
        key    = actual['invoice_number'] or pdf.name
        recorded[key] = actual

        if update:
            print(f'recorded {key}')
            continue

        checked += 1
        # Reconciliation is the invariant that holds regardless of what was
        # recorded: parsed rows have to add up to the invoice's own total.
        if abs(actual['residual']) > P.FEDEX_TOTAL_TOLERANCE:
            failures.append(f'{key}: residual ${actual["residual"]:,.2f} — '
                            f'does not reconcile')

        want = expected.get(key)
        if want is None:
            failures.append(f'{key}: no recorded expectation — run with --update')
        else:
            for field, value in want.items():
                if actual.get(field) != value:
                    failures.append(f'{key}: {field} was {value!r}, '
                                    f'now {actual.get(field)!r}')

        failures += [f'{key}: {f}' for f in check_relabelling(pdf)]

    if update:
        EXPECTED_PATH.write_text(
            json.dumps(recorded, indent=2, sort_keys=True) + '\n', encoding='utf-8')
        print(f'\nWrote {EXPECTED_PATH.name} — {len(recorded)} invoices.')
        return 0

    print(f'{checked} invoices checked, {len(RELABEL_SCENARIOS)} relabelling '
          f'scenarios each.')
    if failures:
        print(f'\nFAILED — {len(failures)} problem(s):')
        for f in failures:
            print(f'  {f}')
        return 1
    print('PASSED')
    return 0


if __name__ == '__main__':
    sys.exit(main())
