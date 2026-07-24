"""
Core reconciliation logic.

Turns an "ibft-transaction_*.xlsx" export into a "need_to_reversal_*.xlsx"
workbook containing:

  - failed      : FAILED transactions, excluding "Insufficient funds"
  - coop        : manual-reversal transactions for every aggregator other
                  than IME REMIT / CITY REMIT, grouped by Member Name
  - imeremit    : manual-reversal transactions for the IME REMIT aggregator
  - cityremit   : manual-reversal transactions for the CITY REMIT aggregator
  - timeout     : TIMEOUT transactions, unchanged

This logic was reverse-engineered from a real ibft-transaction file and its
matching need_to_reversal output and validated row-for-row (104/104 reversal
rows, 29/29 failed rows, 1/1 timeout row) before being written here.
"""
from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# The house/clearing account that sits on the debit side of a reversal
# entry when the transaction's Debtor Bank is (Global IME) Bank — this is
# the normal/current case.
GLOBAL_BANK_DEBIT_ACCOUNT = "0002335524115"
GLOBAL_BANK_KEYWORD = "GLOBAL"
# kept as an alias since other code/tests refer to this name
CLEARING_DEBIT_ACCOUNT = GLOBAL_BANK_DEBIT_ACCOUNT

# Prabhu Bank is a rarer case (it was the Debtor Bank before Global IME
# Bank took over that role) — when it shows up, the reversal's Debit
# Account Number must be this dedicated account instead.
PRABHU_BANK_DEBIT_ACCOUNT = "99901170130555"
PRABHU_BANK_KEYWORD = "PRABHU"

# Network types for which the charge amount is always forced to 0 on the
# reversal, regardless of what the original file shows.
ZERO_CHARGE_NETWORK_KEYWORDS = ("NCHL",)

SOURCE_SHEET_NAME_CANDIDATES = ["Transactions", "transactions"]

# Column names expected in the source workbook (must match, case sensitive,
# the header row of the "Transactions" sheet).
REQUIRED_COLUMNS = [
    "S NO",
    "Transaction Date",
    "Member Name",
    "Aggregator",
    "Member Transaction Id",
    "Network Reference Id",
    "Session Id",
    "Payment Processor",
    "Transaction Amount",
    "Charge Amount",
    "Debtor Bank",
    "Debit Status",
    "Debit Response Code",
    "Debit Account Number",
    "Debitor Account Name",
    "Creditor Bank",
    "Credit Status",
    "Credit Response Code",
    "Credit Account Number",
    "Creditor Account Name",
    "Source Message",
    "Destination Message",
    "Overall Status",
]

# Columns pulled, in order, from the source row for the reversal-style sheets
# (coop / imeremit / cityremit).
REVERSAL_SOURCE_COLUMNS = [
    "S NO",
    "Transaction Date",
    "Member Name",
    "Aggregator",
    "Member Transaction Id",
    "Network Reference Id",
    "Session Id",
    "Payment Processor",
    "Transaction Amount",
    "Charge Amount",
    "Debtor Bank",
    "Creditor Bank",
    "Debit Account Number",
    "Credit Account Number",
]

# Header row actually written for the reversal-style sheets. The last three
# columns (Debit Account Number, Credit Account Number, Narration) are
# highlighted and are not a straight passthrough of the source row.
REVERSAL_SHEET_HEADERS = [
    "S NO",
    "Transaction Date",
    "Member Name",
    "Aggregator",
    "Member Transaction Id",
    "Network Reference Id",
    "Session Id",
    "Payment Processor",
    "Transaction Amount",
    "Charge Amount",
    "Debtor Bank",
    "Creditor Bank",
    "Debit Account Number",
    "Credit Account Number",
    "Narration",
]

# How many trailing columns get the yellow header highlight.
HIGHLIGHTED_TRAILING_COLUMNS = 3

IME_AGGREGATOR = "IME REMIT"
CITY_AGGREGATOR = "CITY REMIT"


class ProcessingError(Exception):
    """Raised when the uploaded workbook doesn't look like a valid export."""


@dataclass
class ProcessingStats:
    total_rows: int = 0
    # The latest "Transaction Date" value found anywhere in this upload's
    # source file — i.e. how far into the day this export actually
    # reaches (e.g. "downloaded at 10am, so data through 09:58am"), shown
    # on the dashboard's day-by-day breakdown. Not the same as wall-clock
    # processing time.
    data_through_at: Any = None
    success_count: int = 0
    success_amount: float = 0.0
    failed_total: int = 0
    failed_insufficient_funds: int = 0
    failed_kept: int = 0
    failed_amount: float = 0.0
    failed_charge: float = 0.0
    # Amount/charge summed only across the rows actually kept in the
    # "failed" sheet (i.e. failed_total minus insufficient-funds) — needed
    # so the bank-statement check can subtract out "failed but credited"
    # rows precisely.
    failed_kept_amount: float = 0.0
    failed_kept_charge: float = 0.0
    reversal_total: int = 0
    reversal_manual_kept: int = 0
    reversal_manual_amount: float = 0.0
    reversal_manual_charge: float = 0.0
    reversal_system_count: int = 0
    reversal_system_amount: float = 0.0
    reversal_system_charge: float = 0.0
    coop_count: int = 0
    imeremit_count: int = 0
    cityremit_count: int = 0
    timeout_count: int = 0
    timeout_amount: float = 0.0
    coop_member_count: int = 0
    prabhu_rerouted: int = 0
    duplicate_skipped: int = 0
    # Rows skipped because their Network Reference Id was already seen in
    # an *earlier upload's source file* (any status) — this catches the
    # "downloaded an overlapping time window" case, as opposed to
    # duplicate_skipped above, which only catches manual-reversal rows
    # repeated across generated reversal files.
    duplicate_source_skipped: int = 0
    unrecognized_debtor_bank_rows: int = 0
    unrecognized_debtor_banks: list[str] | None = None
    failed_reason_breakdown: dict[str, int] | None = None
    failed_onus_count: int = 0
    failed_onus_amount: float = 0.0
    failed_offus_count: int = 0
    failed_offus_amount: float = 0.0
    failed_reason_breakdown_onus: dict[str, int] | None = None
    failed_reason_breakdown_offus: dict[str, int] | None = None
    # Not persisted to the DB directly (see views.py) — every normalized
    # Network Reference Id that was actually kept/counted this run, so the
    # caller can remember them for next time's overlapping-window dedup.
    kept_reference_ids: set[str] = field(default_factory=set)

    def as_dict(self) -> dict[str, Any]:
        return {
            "total_rows": self.total_rows,
            "data_through_at": self.data_through_at,
            "success_count": self.success_count,
            "success_amount": self.success_amount,
            "failed_total": self.failed_total,
            "failed_insufficient_funds": self.failed_insufficient_funds,
            "failed_kept": self.failed_kept,
            "failed_amount": self.failed_amount,
            "failed_charge": self.failed_charge,
            "failed_kept_amount": self.failed_kept_amount,
            "failed_kept_charge": self.failed_kept_charge,
            "reversal_total": self.reversal_total,
            "reversal_manual_kept": self.reversal_manual_kept,
            "reversal_manual_amount": self.reversal_manual_amount,
            "reversal_manual_charge": self.reversal_manual_charge,
            "reversal_system_count": self.reversal_system_count,
            "reversal_system_amount": self.reversal_system_amount,
            "reversal_system_charge": self.reversal_system_charge,
            "coop_count": self.coop_count,
            "imeremit_count": self.imeremit_count,
            "cityremit_count": self.cityremit_count,
            "timeout_count": self.timeout_count,
            "timeout_amount": self.timeout_amount,
            "coop_member_count": self.coop_member_count,
            "prabhu_rerouted": self.prabhu_rerouted,
            "duplicate_skipped": self.duplicate_skipped,
            "duplicate_source_skipped": self.duplicate_source_skipped,
            "unrecognized_debtor_bank_rows": self.unrecognized_debtor_bank_rows,
            "unrecognized_debtor_banks": self.unrecognized_debtor_banks or [],
            "failed_reason_breakdown": self.failed_reason_breakdown or {},
            "failed_onus_count": self.failed_onus_count,
            "failed_onus_amount": self.failed_onus_amount,
            "failed_offus_count": self.failed_offus_count,
            "failed_offus_amount": self.failed_offus_amount,
            "failed_reason_breakdown_onus": self.failed_reason_breakdown_onus or {},
            "failed_reason_breakdown_offus": self.failed_reason_breakdown_offus or {},
        }


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def transform_member_id(raw_id: Any) -> str:
    """Strip the leading "noise" off a member transaction id.

    Rule (derived from real data, validated on 104 real rows):
      1. Strip any leading zeros.
      2. Strip everything from the front up to (but not including) the
         first digit that remains.

    Examples:
        'REQ_1783249894525'                 -> '1783249894525'
        'IME-00113275637-572'               -> '00113275637-572'
        'TXN_2421483084574022'              -> '2421483084574022'
        'ACQ_SETTL-2607050005472234LFE'     -> '2607050005472234LFE'
        '0000000000CEMTP402131874/7028457327' -> '402131874/7028457327'
        '26180083314147965' (plain numeric) -> '26180083314147965'
    """
    s = "" if raw_id is None else str(raw_id).strip()

    i = 0
    while i < len(s) and s[i] == "0":
        i += 1
    s = s[i:]

    j = 0
    while j < len(s) and not s[j].isdigit():
        j += 1
    return s[j:]


def build_narration(member_transaction_id: Any, session_id: Any) -> str:
    return f"REV{transform_member_id(member_transaction_id)}-{session_id}"


def is_insufficient_funds(source_message: Any) -> bool:
    return "insufficient funds" in str(source_message or "").lower()


def is_manual_reversal(source_message: Any) -> bool:
    return "manual reversal" in str(source_message or "").lower()


def is_zero_charge_network(payment_processor: Any) -> bool:
    """NCHL-routed transactions always carry a 0 charge on the reversal."""
    value = str(payment_processor or "").upper()
    return any(keyword in value for keyword in ZERO_CHARGE_NETWORK_KEYWORDS)


def resolve_charge_amount(row, col) -> Any:
    charge = row[col["Charge Amount"]]
    if is_zero_charge_network(row[col["Payment Processor"]]):
        return 0
    return charge


def to_float(value: Any) -> float:
    """Safely coerce a cell value to float for summing (amount/charge
    columns). Blank/None/non-numeric values count as 0 rather than
    raising, so a stray text cell can't blow up the whole run."""
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def parse_transaction_datetime(value: Any):
    """Parse a "Transaction Date" cell (e.g. "2026-07-17T13:59:05.461")
    into a datetime, or None if it isn't parseable. Used to figure out
    how far into the day an uploaded ibft export actually reaches — i.e.
    "processed up to" — rather than reporting wall-clock processing time."""
    import datetime as _dt

    if value is None or value == "":
        return None
    if isinstance(value, _dt.datetime):
        return value
    text = str(value).strip()
    if not text:
        return None
    try:
        return _dt.datetime.fromisoformat(text)
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%d-%m-%Y %H:%M:%S"):
        try:
            return _dt.datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def is_prabhu_bank(debtor_bank: Any) -> bool:
    return PRABHU_BANK_KEYWORD in str(debtor_bank or "").upper()


def is_global_bank(debtor_bank: Any) -> bool:
    return GLOBAL_BANK_KEYWORD in str(debtor_bank or "").upper()


def resolve_debit_account(debtor_bank: Any) -> tuple[str, bool]:
    """Returns (debit_account_number, recognized).

    Debtor Bank is Global IME Bank in the normal/current case, and was
    Prabhu Bank before that — both are explicitly recognized and routed to
    their own dedicated account. Anything else is *not* silently assumed to
    be fine: it still falls back to the standard account (so processing
    isn't blocked), but `recognized=False` is bubbled up so it gets counted
    and surfaced on the audit log / dashboard for manual review, in case a
    third bank shows up in the future that also needs its own account.
    """
    if is_prabhu_bank(debtor_bank):
        return PRABHU_BANK_DEBIT_ACCOUNT, True
    if is_global_bank(debtor_bank):
        return GLOBAL_BANK_DEBIT_ACCOUNT, True
    return GLOBAL_BANK_DEBIT_ACCOUNT, False


def normalize_reference_id(value: Any) -> str:
    """Normalize a Network Reference Id (or similar) for reliable
    comparison — same value read as ' ABC123 ', 'abc123', or a numeric type
    should all compare equal."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip().upper()


def normalize_failure_reason(source_message: Any) -> str:
    """Bucket a raw Source Message into one of a small, fixed set of
    reason labels for analytics, matched case-insensitively:

        Timeout, Insufficient fund, Card issuer timeout, Response timeout,
        Transaction amount exceeded, Other

    Everything that doesn't match a known phrase (regardless of case)
    falls into "Other" — so the dashboard breakdown always sums cleanly
    to the total failed count instead of fragmenting into many one-off
    raw-message buckets.
    """
    msg = str(source_message or "").strip()
    if not msg:
        return "Other"

    lowered = msg.lower()

    # Order matters: check the more specific phrases before the generic
    # "timeout" one so e.g. "card issuer timeout" doesn't just land in
    # the generic "Timeout" bucket.
    known_buckets = [
        ("insufficient fund", "Insufficient fund"),
        ("card issuer timeout", "Card issuer timeout"),
        ("issuer timeout", "Card issuer timeout"),
        ("response timeout", "Response timeout"),
        ("transaction amount exceed", "Transaction amount exceeded"),
        ("amount exceed", "Transaction amount exceeded"),
        ("time out", "Timeout"),
        ("timeout", "Timeout"),
    ]
    for needle, label in known_buckets:
        if needle in lowered:
            return label

    return "Other"


def is_on_us(debtor_bank: Any, creditor_bank: Any) -> bool:
    """On-Us = both the Debtor Bank AND the Creditor Bank are our own bank
    (Global IME Bank) — i.e. Global-to-Global. Off-Us = Global-to-other
    (or any other combination). Used to split the failed-transaction
    breakdown by where the failure actually originated."""
    return is_global_bank(debtor_bank) and is_global_bank(creditor_bank)


def _find_header_row(ws) -> tuple[int, dict[str, int]]:
    """Locate the header row (some exports have a merged 'Period: ...'
    banner row above the real header) and return (row_index, col_index_map).
    col_index_map maps column name -> 0-based index within that row.
    """
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=min(10, ws.max_row), values_only=True)
    ):
        names = [str(c).strip() if c is not None else "" for c in row]
        if "S NO" in names and "Overall Status" in names:
            col_map = {name: i for i, name in enumerate(names) if name}
            missing = [c for c in REQUIRED_COLUMNS if c not in col_map]
            if missing:
                raise ProcessingError(
                    "The uploaded file is missing expected column(s): "
                    + ", ".join(missing)
                )
            return row_idx, col_map
    raise ProcessingError(
        "Could not find the header row (expected a row containing "
        "'S NO' and 'Overall Status'). Please upload the original "
        "ibft-transaction export."
    )


def _load_transactions(input_path: str | Path):
    wb = openpyxl.load_workbook(input_path, data_only=True)

    ws = None
    for name in SOURCE_SHEET_NAME_CANDIDATES:
        if name in wb.sheetnames:
            ws = wb[name]
            break
    if ws is None:
        # fall back to the first sheet
        ws = wb[wb.sheetnames[0]]

    header_row_idx, col_map = _find_header_row(ws)

    rows = []
    for row in ws.iter_rows(min_row=header_row_idx + 2, values_only=True):
        if row[col_map["S NO"]] is None and all(v is None for v in row):
            continue
        rows.append(row)

    return rows, col_map


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def process_ibft_file(
    input_path: str | Path,
    output_path: str | Path,
    previously_reversed_ids: set[str] | None = None,
    seen_reference_ids: set[str] | None = None,
) -> ProcessingStats:
    """Read the ibft-transaction export at `input_path` and write the
    reconciled need_to_reversal workbook to `output_path`.

    `previously_reversed_ids` is the set of (normalized) Network Reference
    Ids that already appeared in the *previous* generated reversal file.
    Any manual-reversal row whose Network Reference Id is in that set is
    skipped here, so the same transaction can never be reversed twice just
    because it was still showing up as REVERSAL in a later export.

    `seen_reference_ids` is the set of (normalized) Network Reference Ids
    already processed in *any earlier upload* (any status — success,
    failed, reversal, timeout). Daily exports are typically downloaded as
    "since midnight through whenever I downloaded", so two exports on
    consecutive days overlap for the hours between midnight and the
    previous download time. Any row whose Network Reference Id is already
    in this set is skipped entirely (counted as `duplicate_source_skipped`)
    before it's categorized by Overall Status, so that overlap is never
    counted or reversed twice.
    """
    previously_reversed_ids = previously_reversed_ids or set()
    seen_reference_ids = seen_reference_ids or set()
    rows, col = _load_transactions(input_path)
    stats = ProcessingStats(total_rows=len(rows))

    if "Transaction Date" in col:
        date_idx = col["Transaction Date"]
        parsed_dates = (parse_transaction_datetime(r[date_idx]) for r in rows)
        latest = max((d for d in parsed_dates if d is not None), default=None)
        if latest is not None:
            try:
                from django.utils import timezone as _tz

                if _tz.is_naive(latest):
                    latest = _tz.make_aware(latest)
            except Exception:
                pass  # services.py stays usable outside Django too
        stats.data_through_at = latest

    def g(row, name):
        return row[col[name]]

    failed_kept_rows = []
    timeout_rows = []
    reversal_coop_rows = []
    reversal_ime_rows = []
    reversal_city_rows = []
    reason_breakdown: dict[str, int] = {}
    reason_breakdown_onus: dict[str, int] = {}
    reason_breakdown_offus: dict[str, int] = {}
    unrecognized_banks_seen: dict[str, int] = {}
    kept_reference_ids: set[str] = set()

    for row in rows:
        row_ref_id = normalize_reference_id(g(row, "Network Reference Id"))
        if row_ref_id and row_ref_id in seen_reference_ids:
            stats.duplicate_source_skipped += 1
            continue
        if row_ref_id:
            kept_reference_ids.add(row_ref_id)

        overall = str(g(row, "Overall Status") or "").strip().upper()

        if overall == "SUCCESS":
            stats.success_count += 1
            stats.success_amount += to_float(g(row, "Transaction Amount"))

        elif overall == "FAILED":
            stats.failed_total += 1
            stats.failed_amount += to_float(g(row, "Transaction Amount"))
            stats.failed_charge += to_float(resolve_charge_amount(row, col))
            reason = normalize_failure_reason(g(row, "Source Message"))
            reason_breakdown[reason] = reason_breakdown.get(reason, 0) + 1
            txn_amount = to_float(g(row, "Transaction Amount"))
            if is_on_us(g(row, "Debtor Bank"), g(row, "Creditor Bank")):
                stats.failed_onus_count += 1
                stats.failed_onus_amount += txn_amount
                reason_breakdown_onus[reason] = reason_breakdown_onus.get(reason, 0) + 1
            else:
                stats.failed_offus_count += 1
                stats.failed_offus_amount += txn_amount
                reason_breakdown_offus[reason] = reason_breakdown_offus.get(reason, 0) + 1
            if is_insufficient_funds(g(row, "Source Message")):
                stats.failed_insufficient_funds += 1
            else:
                failed_kept_rows.append(row)
                stats.failed_kept_amount += to_float(g(row, "Transaction Amount"))
                stats.failed_kept_charge += to_float(resolve_charge_amount(row, col))

        elif overall == "TIMEOUT":
            stats.timeout_count += 1
            stats.timeout_amount += to_float(g(row, "Transaction Amount"))
            timeout_rows.append(row)

        elif overall == "REVERSAL":
            stats.reversal_total += 1
            if is_manual_reversal(g(row, "Source Message")):
                ref_id = normalize_reference_id(g(row, "Network Reference Id"))
                if ref_id and ref_id in previously_reversed_ids:
                    stats.duplicate_skipped += 1
                    continue
                stats.reversal_manual_kept += 1
                stats.reversal_manual_amount += to_float(g(row, "Transaction Amount"))
                stats.reversal_manual_charge += to_float(resolve_charge_amount(row, col))
                debtor_bank_name = g(row, "Debtor Bank")
                _, recognized = resolve_debit_account(debtor_bank_name)
                if is_prabhu_bank(debtor_bank_name):
                    stats.prabhu_rerouted += 1
                if not recognized:
                    label = str(debtor_bank_name or "(blank)").strip() or "(blank)"
                    unrecognized_banks_seen[label] = unrecognized_banks_seen.get(label, 0) + 1
                aggregator = str(g(row, "Aggregator") or "").strip().upper()
                if aggregator == IME_AGGREGATOR:
                    reversal_ime_rows.append(row)
                elif aggregator == CITY_AGGREGATOR:
                    reversal_city_rows.append(row)
                else:
                    reversal_coop_rows.append(row)
            else:
                # System-initiated reversal — already reversed automatically
                # by the switch, so it doesn't need a manual reversal row in
                # the output file, but it's still tracked here for daily
                # reconciliation (how much moved via system reversal vs.
                # manual reversal).
                stats.reversal_system_count += 1
                stats.reversal_system_amount += to_float(g(row, "Transaction Amount"))
                stats.reversal_system_charge += to_float(resolve_charge_amount(row, col))

    stats.failed_kept = len(failed_kept_rows)
    stats.imeremit_count = len(reversal_ime_rows)
    stats.cityremit_count = len(reversal_city_rows)
    stats.coop_count = len(reversal_coop_rows)
    stats.unrecognized_debtor_bank_rows = sum(unrecognized_banks_seen.values())
    stats.unrecognized_debtor_banks = sorted(unrecognized_banks_seen.keys())
    stats.failed_reason_breakdown = reason_breakdown
    stats.failed_reason_breakdown_onus = reason_breakdown_onus
    stats.failed_reason_breakdown_offus = reason_breakdown_offus
    stats.kept_reference_ids = kept_reference_ids

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    _write_raw_sheet(out_wb, "failed", failed_kept_rows, col)
    _write_reversal_sheet_grouped_by_member(out_wb, "coop", reversal_coop_rows, col, stats)
    _write_reversal_sheet_flat(out_wb, "imeremit", reversal_ime_rows, col)
    _write_reversal_sheet_flat(out_wb, "cityremit", reversal_city_rows, col)
    _write_raw_sheet(out_wb, "timeout", timeout_rows, col)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    out_wb.save(output_path)

    return stats


# ---------------------------------------------------------------------------
# Double-reversal prevention: read Network Reference Ids out of a
# previously generated reversal workbook.
# ---------------------------------------------------------------------------

_REVERSAL_SHEET_NAMES = ("coop", "imeremit", "cityremit")


def extract_reversal_network_reference_ids(reversal_file_path: str | Path) -> set[str]:
    """Open a previously generated need_to_reversal_*.xlsx and collect every
    Network Reference Id across its reversal-style sheets (coop / imeremit /
    cityremit), so the next run can exclude them and avoid a double
    reversal. Silently returns an empty set if the file can't be read
    (e.g. it's missing or was deleted) rather than blocking new processing.
    """
    ids: set[str] = set()
    try:
        wb = openpyxl.load_workbook(reversal_file_path, data_only=True, read_only=True)
    except Exception:
        return ids

    try:
        for sheet_name in _REVERSAL_SHEET_NAMES:
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            header_seen = False
            ref_col_idx = None
            for row in ws.iter_rows(values_only=True):
                if all(v is None for v in row):
                    # blank separator row between coop member groups
                    header_seen = False
                    continue
                if not header_seen:
                    names = [str(c).strip() if c is not None else "" for c in row]
                    if "Network Reference Id" in names:
                        ref_col_idx = names.index("Network Reference Id")
                        header_seen = True
                    continue
                if ref_col_idx is not None and ref_col_idx < len(row):
                    ref_id = normalize_reference_id(row[ref_col_idx])
                    if ref_id:
                        ids.add(ref_id)
    finally:
        wb.close()

    return ids


# ---------------------------------------------------------------------------
# Bank statement cross-check
#
# The bank statement export (e.g. "global_BankStatement_<from>_<to>.csv")
# has columns: S.N, ENTRY TYPE (CR/DR), REMARKS, AMOUNT, DATE. A row's
# Network Reference Id shows up verbatim as one of the "/"- or ":"-
# separated tokens inside REMARKS, e.g. a Network Reference Id of
# "00000000J406" appears in a REMARKS value like:
#   "IME-90809428421-845:90809428421/00000000J406:IME REMIT/S43161794"
# and the trailing "/S43161794" is the switch's own ISO/session id for
# that statement line — used to tell whether a given credit into the
# parking account has *already* been reversed back out (a matching DR
# entry elsewhere in the statement carrying the same ISO id).
# ---------------------------------------------------------------------------

_STATEMENT_TOKEN_SPLIT_RE = re.compile(r"[/:]")
_STATEMENT_ISO_ID_RE = re.compile(r"/(S\d+)\s*$")


@dataclass
class BankStatementIndex:
    """In-memory index over a parsed bank statement, built once and reused
    for every row being checked."""

    # normalized token (e.g. a Network Reference Id) -> list of statement
    # rows ({"entry_type", "remarks", "amount", "date"}) whose REMARKS
    # contains that token.
    by_token: dict[str, list[dict]] = field(default_factory=dict)
    # ISO id (e.g. "S43161794") -> set of entry types ("CR"/"DR") seen
    # against that ISO id anywhere in the statement.
    iso_entry_types: dict[str, set[str]] = field(default_factory=dict)
    total_rows: int = 0


def _read_bank_statement_rows(path: str | Path) -> list[dict]:
    """Read a bank statement export (.csv or .xlsx) into a list of dicts
    with keys entry_type / remarks / amount / date. Column names are
    matched case-insensitively (S.N / ENTRY TYPE / REMARKS / AMOUNT / DATE)."""
    path = Path(path)
    suffix = path.suffix.lower()

    def _normalize_headers(raw_headers):
        return [str(h).strip().upper() if h is not None else "" for h in raw_headers]

    rows: list[dict] = []

    if suffix == ".csv":
        with open(path, "r", newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            try:
                header = _normalize_headers(next(reader))
            except StopIteration:
                raise ProcessingError("The bank statement file is empty.")
            try:
                entry_idx = header.index("ENTRY TYPE")
                remarks_idx = header.index("REMARKS")
            except ValueError:
                raise ProcessingError(
                    "The bank statement is missing an 'ENTRY TYPE' or "
                    "'REMARKS' column. Please upload the original bank "
                    "statement export."
                )
            amount_idx = header.index("AMOUNT") if "AMOUNT" in header else None
            date_idx = header.index("DATE") if "DATE" in header else None
            for raw in reader:
                if not raw or all(not str(c).strip() for c in raw):
                    continue
                rows.append(
                    {
                        "entry_type": raw[entry_idx].strip() if entry_idx < len(raw) else "",
                        "remarks": raw[remarks_idx].strip() if remarks_idx < len(raw) else "",
                        "amount": raw[amount_idx] if amount_idx is not None and amount_idx < len(raw) else None,
                        "date": raw[date_idx] if date_idx is not None and date_idx < len(raw) else None,
                    }
                )
    else:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            header = None
            entry_idx = remarks_idx = amount_idx = date_idx = None
            for raw in ws.iter_rows(values_only=True):
                if header is None:
                    names = _normalize_headers(raw)
                    if "ENTRY TYPE" in names and "REMARKS" in names:
                        header = names
                        entry_idx = header.index("ENTRY TYPE")
                        remarks_idx = header.index("REMARKS")
                        amount_idx = header.index("AMOUNT") if "AMOUNT" in header else None
                        date_idx = header.index("DATE") if "DATE" in header else None
                    continue
                if raw is None or all(v is None for v in raw):
                    continue
                rows.append(
                    {
                        "entry_type": str(raw[entry_idx]).strip() if entry_idx < len(raw) and raw[entry_idx] is not None else "",
                        "remarks": str(raw[remarks_idx]).strip() if remarks_idx < len(raw) and raw[remarks_idx] is not None else "",
                        "amount": raw[amount_idx] if amount_idx is not None and amount_idx < len(raw) else None,
                        "date": raw[date_idx] if date_idx is not None and date_idx < len(raw) else None,
                    }
                )
            if header is None:
                raise ProcessingError(
                    "Could not find the bank statement header row (expected "
                    "columns including 'ENTRY TYPE' and 'REMARKS')."
                )
        finally:
            wb.close()

    return rows


def build_bank_statement_index(path: str | Path) -> BankStatementIndex:
    rows = _read_bank_statement_rows(path)
    by_token: dict[str, list[dict]] = {}
    iso_entry_types: dict[str, set[str]] = {}

    for row in rows:
        remarks = row["remarks"] or ""
        entry_type = (row["entry_type"] or "").strip().upper()

        tokens = {t.strip().upper() for t in _STATEMENT_TOKEN_SPLIT_RE.split(remarks) if t.strip()}
        for tok in tokens:
            by_token.setdefault(tok, []).append(row)

        iso_match = _STATEMENT_ISO_ID_RE.search(remarks.strip())
        if iso_match:
            iso_id = iso_match.group(1).strip().upper()
            iso_entry_types.setdefault(iso_id, set()).add(entry_type)

    return BankStatementIndex(by_token=by_token, iso_entry_types=iso_entry_types, total_rows=len(rows))


def statement_entries_for_reference(index: BankStatementIndex, ref_id: Any) -> list[dict]:
    """Every bank statement row whose REMARKS contains this (normalized)
    Network Reference Id as a token."""
    return index.by_token.get(normalize_reference_id(ref_id), [])


def is_failed_but_credited(index: BankStatementIndex, ref_id: Any) -> bool:
    """A FAILED-status transaction whose Network Reference Id nonetheless
    shows up in the bank statement means money actually moved (credited
    into the parking account) despite the switch reporting it FAILED —
    this needs a manual reversal just like a normal failed-turned-reversal
    row would."""
    return len(statement_entries_for_reference(index, ref_id)) > 0


def is_already_reversed(index: BankStatementIndex, ref_id: Any) -> bool:
    """A manual-reversal row's Network Reference Id should show up as a CR
    entry in the bank statement (the original credit into the parking
    account). That CR entry's REMARKS ends with the switch's own ISO id
    (e.g. "/S43161794"). If that same ISO id also shows up anywhere inside
    a DR entry's REMARKS elsewhere in the statement (not necessarily at
    the end of it — a reversal's own REMARKS typically has the *original*
    ISO id somewhere in the middle, followed by the reversal transaction's
    own trailing id), the money has already gone back out — i.e. support
    (or the system) already reversed it, and this row on the manual
    reversal file would be a double reversal."""
    entries = statement_entries_for_reference(index, ref_id)
    for entry in entries:
        if (entry["entry_type"] or "").strip().upper() != "CR":
            continue
        iso_match = _STATEMENT_ISO_ID_RE.search((entry["remarks"] or "").strip())
        if not iso_match:
            continue
        iso_id = iso_match.group(1).strip().upper()
        # The ISO id is itself a "/"- or ":"-delimited token, so it was
        # already indexed by build_bank_statement_index() the same way a
        # Network Reference Id is — look it up directly rather than only
        # checking other rows' *trailing* ISO id.
        for candidate in index.by_token.get(iso_id, []):
            if (candidate["entry_type"] or "").strip().upper() == "DR":
                return True
    return False


@dataclass
class BankStatementCheckStats:
    statement_rows: int = 0
    failed_credited_count: int = 0
    failed_credited_amount: float = 0.0
    failed_credited_charge: float = 0.0
    already_reversed_count: int = 0
    already_reversed_amount: float = 0.0
    already_reversed_charge: float = 0.0
    already_reversed_by_sheet: dict[str, int] = field(default_factory=dict)
    # On-Us/Off-Us split of the failed-but-credited rows above, plus their
    # reason labels — used to subtract these back out of the "pure failed"
    # on-us/off-us breakdown on the log (a credited row isn't a pure
    # failure any more, so it shouldn't count toward that breakdown).
    failed_credited_onus_count: int = 0
    failed_credited_offus_count: int = 0
    failed_credited_onus_amount: float = 0.0
    failed_credited_offus_amount: float = 0.0
    failed_credited_reason_onus: dict[str, int] = field(default_factory=dict)
    failed_credited_reason_offus: dict[str, int] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        return {
            "failed_credited_count": self.failed_credited_count,
            "failed_credited_amount": self.failed_credited_amount,
            "failed_credited_charge": self.failed_credited_charge,
            "already_reversed_count": self.already_reversed_count,
            "already_reversed_amount": self.already_reversed_amount,
            "already_reversed_charge": self.already_reversed_charge,
        }


_RED_FLAG_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")


def _flag_row_red(ws, row_number):
    for cell in ws[row_number]:
        cell.fill = _RED_FLAG_FILL
        existing = cell.font
        cell.font = Font(bold=bool(existing and existing.bold), color="9C0006")


def apply_bank_statement_to_reversal_file(
    reversal_file_path: str | Path,
    statement_path: str | Path,
) -> BankStatementCheckStats:
    """Re-open an already-generated need_to_reversal_*.xlsx and cross-check
    it against a bank statement export:

      - "failed" sheet: any row whose Network Reference Id shows up in the
        statement gets red-flagged in place, and is also copied into a new
        "Failed but Credited" sheet laid out like the reversal sheets
        (same derived Debit/Credit Account Number + Narration columns)
        since it needs to be manually reversed.
      - "coop" / "imeremit" / "cityremit" sheets: any row whose original
        credit has already been reversed out (per `is_already_reversed`)
        gets red-flagged in place, so whoever works the file can see it
        was already handled and skip it (rather than double-reversing).

    The workbook is saved back to `reversal_file_path` in place. Rows are
    never deleted or reordered — only highlighted — so nothing about the
    file's existing structure changes for anyone already relying on it.
    """
    index = build_bank_statement_index(statement_path)

    try:
        wb = openpyxl.load_workbook(reversal_file_path)
    except Exception as exc:
        raise ProcessingError(f"Could not open the generated reversal file: {exc}") from exc

    stats = BankStatementCheckStats(statement_rows=index.total_rows)

    # --- "failed" sheet -----------------------------------------------
    failed_but_credited_rows: list[tuple] = []
    if "failed" in wb.sheetnames:
        ws = wb["failed"]
        header = [c.value for c in ws[1]]
        col = {name: i for i, name in enumerate(header) if name}
        if "Network Reference Id" in col:
            ref_idx = col["Network Reference Id"]
            for row_number in range(2, ws.max_row + 1):
                values = [ws.cell(row=row_number, column=c + 1).value for c in range(len(header))]
                if all(v is None for v in values):
                    continue
                ref_id = values[ref_idx]
                if not ref_id or not is_failed_but_credited(index, ref_id):
                    continue
                _flag_row_red(ws, row_number)
                stats.failed_credited_count += 1
                if "Transaction Amount" in col:
                    stats.failed_credited_amount += to_float(values[col["Transaction Amount"]])
                if "Charge Amount" in col and "Payment Processor" in col:
                    stats.failed_credited_charge += to_float(resolve_charge_amount(tuple(values), col))
                if "Debtor Bank" in col and "Creditor Bank" in col:
                    debtor_bank = values[col["Debtor Bank"]]
                    creditor_bank = values[col["Creditor Bank"]]
                    reason = normalize_failure_reason(
                        values[col["Source Message"]] if "Source Message" in col else None
                    )
                    if is_on_us(debtor_bank, creditor_bank):
                        stats.failed_credited_onus_count += 1
                        if "Transaction Amount" in col:
                            stats.failed_credited_onus_amount += to_float(values[col["Transaction Amount"]])
                        stats.failed_credited_reason_onus[reason] = (
                            stats.failed_credited_reason_onus.get(reason, 0) + 1
                        )
                    else:
                        stats.failed_credited_offus_count += 1
                        if "Transaction Amount" in col:
                            stats.failed_credited_offus_amount += to_float(values[col["Transaction Amount"]])
                        stats.failed_credited_reason_offus[reason] = (
                            stats.failed_credited_reason_offus.get(reason, 0) + 1
                        )
                failed_but_credited_rows.append(tuple(values))

    if failed_but_credited_rows:
        # Reuses the same "REQUIRED_COLUMNS-ordered raw row" -> reversal-row
        # transform used for the normal reversal sheets, since the
        # "failed" sheet is written with that same 23-column layout.
        raw_col = {name: i for i, name in enumerate(REQUIRED_COLUMNS)}
        if "Failed but Credited" in wb.sheetnames:
            del wb["Failed but Credited"]
        _write_reversal_sheet_flat(wb, "Failed but Credited", failed_but_credited_rows, raw_col)

    # --- coop / imeremit / cityremit sheets -----------------------------
    for sheet_name in _REVERSAL_SHEET_NAMES:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header = None
        col: dict[str, int] = {}
        for row_number in range(1, ws.max_row + 1):
            values = [ws.cell(row=row_number, column=c + 1).value for c in range(ws.max_column)]
            if all(v is None for v in values):
                header = None  # blank separator row between coop member groups
                continue
            if values[0] == "S NO":
                header = values
                col = {name: i for i, name in enumerate(header) if name}
                continue
            if header is None or "Network Reference Id" not in col:
                continue
            ref_id = values[col["Network Reference Id"]]
            if not ref_id or not is_already_reversed(index, ref_id):
                continue
            _flag_row_red(ws, row_number)
            stats.already_reversed_count += 1
            stats.already_reversed_by_sheet[sheet_name] = stats.already_reversed_by_sheet.get(sheet_name, 0) + 1
            if "Transaction Amount" in col:
                stats.already_reversed_amount += to_float(values[col["Transaction Amount"]])
            if "Charge Amount" in col:
                stats.already_reversed_charge += to_float(values[col["Charge Amount"]])

    Path(reversal_file_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(reversal_file_path)
    wb.close()

    return stats


# ---------------------------------------------------------------------------
# Sheet writers
# ---------------------------------------------------------------------------

_HEADER_FONT = Font(bold=True)
_YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
_THIN_SIDE = Side(style="thin", color="000000")
_ALL_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)


def _autofit(ws):
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        col_letter = get_column_letter(col_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 2, 10), 45)


def _style_header_row(ws, row_number, total_columns, highlight_trailing=0):
    """Bold + bordered header row; optionally highlight the last N columns
    (yellow fill) to call out generated/derived fields."""
    highlight_from = total_columns - highlight_trailing + 1
    for cell in ws[row_number]:
        cell.font = _HEADER_FONT
        cell.border = _ALL_BORDER
        if cell.column >= highlight_from:
            cell.fill = _YELLOW_FILL


def _style_data_row(ws, row_number):
    for cell in ws[row_number]:
        cell.border = _ALL_BORDER


# Columns that must always be written/kept as TEXT, never as a number.
# Long numeric-looking ids (17+ digits) silently lose trailing precision if
# Excel is allowed to store them as a number — this was the root cause of a
# past merchant-id mismatch between the source file and the generated
# reversal file, so every id-like column is force-formatted as text.
_ID_LIKE_COLUMN_NAMES = {
    "Member Transaction Id",
    "Network Reference Id",
    "Session Id",
    "Debit Account Number",
    "Credit Account Number",
}


def _force_text_cell(row_number, ws, col_idx):
    cell = ws.cell(row=row_number, column=col_idx)
    if cell.value is not None:
        cell.value = str(cell.value)
    cell.number_format = "@"


def _apply_id_text_format(ws, row_number, headers):
    for idx, name in enumerate(headers, start=1):
        if name in _ID_LIKE_COLUMN_NAMES:
            _force_text_cell(row_number, ws, idx)


def _write_raw_sheet(wb, sheet_name, rows, col):
    """Write rows using the original 23-column layout, with S NO
    re-serialized (1, 2, 3, ...) rather than the original file's S NO."""
    ws = wb.create_sheet(sheet_name)
    ws.append(REQUIRED_COLUMNS)
    _style_header_row(ws, ws.max_row, len(REQUIRED_COLUMNS))

    for i, row in enumerate(rows, start=1):
        values = [row[col[name]] for name in REQUIRED_COLUMNS]
        values[0] = i  # re-serialize S NO
        ws.append(values)
        _style_data_row(ws, ws.max_row)
        _apply_id_text_format(ws, ws.max_row, REQUIRED_COLUMNS)

    _autofit(ws)
    return ws


def _reversal_row_values(row, col, s_no):
    def g(name):
        return row[col[name]]

    narration = build_narration(g("Member Transaction Id"), g("Session Id"))
    charge_amount = resolve_charge_amount(row, col)
    debit_account, _recognized = resolve_debit_account(g("Debtor Bank"))
    return [
        s_no,
        g("Transaction Date"),
        g("Member Name"),
        g("Aggregator"),
        # Member Transaction Id / Network Reference Id are written as text
        # further down (_force_text_columns) so Excel never reinterprets a
        # long numeric-looking id as a number (which silently rounds/
        # truncates it past ~15 significant digits and was the cause of a
        # past merchant-id mismatch).
        g("Member Transaction Id"),
        g("Network Reference Id"),
        g("Session Id"),
        g("Payment Processor"),
        g("Transaction Amount"),
        charge_amount,
        g("Debtor Bank"),
        g("Creditor Bank"),
        debit_account,
        g("Debit Account Number"),  # money goes back to where it came from
        narration,
    ]


def _write_reversal_sheet_flat(wb, sheet_name, rows, col):
    """imeremit / cityremit: one continuous block, original relative order
    preserved (source file is already most-recent-first)."""
    ws = wb.create_sheet(sheet_name)
    ws.append(REVERSAL_SHEET_HEADERS)
    _style_header_row(ws, ws.max_row, len(REVERSAL_SHEET_HEADERS), HIGHLIGHTED_TRAILING_COLUMNS)

    for i, row in enumerate(rows, start=1):
        ws.append(_reversal_row_values(row, col, i))
        _style_data_row(ws, ws.max_row)
        _apply_id_text_format(ws, ws.max_row, REVERSAL_SHEET_HEADERS)

    _autofit(ws)
    return ws


def _write_reversal_sheet_grouped_by_member(wb, sheet_name, rows, col, stats):
    """coop: grouped by Member Name (case-insensitive alphabetical order),
    each group gets its own header row and S NO restarting at 1, separated
    by a single blank row."""
    ws = wb.create_sheet(sheet_name)

    groups: dict[str, list] = {}
    for row in rows:
        member = str(row[col["Member Name"]] or "").strip()
        groups.setdefault(member, []).append(row)

    stats.coop_member_count = len(groups)

    ordered_members = sorted(groups.keys(), key=lambda m: m.lower())

    first_group = True
    for member in ordered_members:
        if not first_group:
            ws.append([None] * len(REVERSAL_SHEET_HEADERS))
        first_group = False

        ws.append(REVERSAL_SHEET_HEADERS)
        _style_header_row(ws, ws.max_row, len(REVERSAL_SHEET_HEADERS), HIGHLIGHTED_TRAILING_COLUMNS)

        for i, row in enumerate(groups[member], start=1):
            ws.append(_reversal_row_values(row, col, i))
            _style_data_row(ws, ws.max_row)
            _apply_id_text_format(ws, ws.max_row, REVERSAL_SHEET_HEADERS)

    _autofit(ws)
    return ws


# ---------------------------------------------------------------------------
# Output filename helper
# ---------------------------------------------------------------------------

_DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})")


def _extract_date_str(source_filename: str) -> str:
    match = _DATE_RE.search(source_filename)
    return match.group(1) if match else date.today().isoformat()


def build_output_filename(source_filename: str) -> str:
    """need_to_reversal_<date>.xlsx"""
    return f"need_to_reversal_{_extract_date_str(source_filename)}.xlsx"


def build_uploaded_filename(source_filename: str) -> str:
    """ibft_txn_data_<date>.xlsx — every uploaded source file is renamed to
    this standardized name (date taken from the original filename if it
    contains one, otherwise today's date) so the audit trail is consistent
    regardless of what the file was originally called."""
    return f"ibft_txn_data_{_extract_date_str(source_filename)}.xlsx"
