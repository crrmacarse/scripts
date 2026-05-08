"""CC Analyzer Plus service.

Refactored from cc_analyzer_plus.py to be web-friendly:
- Reads PDF from in-memory bytes (file upload), not from a path on disk.
- Reads Money Manager data from either the Google Sheet (default) or an
  uploaded xlsx export.
- Returns a structured analysis (matched / missing / unused / inaccurate) so
  the UI can flag issues before any write to the destination sheet.
- Keeps the original Google Sheets write behaviour for the confirm step.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Iterable, Literal

from PyPDF2 import PdfReader
from dateutil.relativedelta import relativedelta
from gspread_formatting import CellFormat, TextFormat, format_cell_range

from app.services.google_auth import get_gspread_client


RowStatus = Literal["matched", "missing", "inaccurate", "duplicate"]


@dataclass
class MMEntry:
    date: str  # m/d/yy
    amount: str  # "1,234.56"
    description: str
    raw_period: str  # original m/d/Y from sheet (for display)

    @property
    def amount_value(self) -> float:
        return float(self.amount.replace(",", ""))

    def key(self) -> tuple[str, str, str]:
        return (self.date, self.amount, self.description)


@dataclass
class PdfRow:
    tran_date: str
    post_date: str
    description: str
    amount: str  # "1,234.56"

    @property
    def amount_value(self) -> float:
        return float(self.amount.replace(",", ""))


@dataclass
class AnalyzedRow:
    pdf: PdfRow
    status: RowStatus
    matched_mm: MMEntry | None = None
    nearest_candidates: list[MMEntry] = field(default_factory=list)
    note: str = ""
    mm_mentions: str = ""


@dataclass
class AnalysisResult:
    sheet_name: str
    billing_period: str
    cutoff_date: str  # ISO yyyy-mm-dd
    rows: list[AnalyzedRow]
    unused_mm: list[MMEntry]
    total_amount: float
    total_cr: float

    @property
    def grand_total(self) -> float:
        return self.total_amount - self.total_cr

    @property
    def issue_count(self) -> int:
        bad = sum(1 for r in self.rows if r.status != "matched")
        return bad + len(self.unused_mm)

    def serialise_for_confirm(self) -> dict:
        return {
            "sheet_name": self.sheet_name,
            "billing_period": self.billing_period,
            "cutoff_date": self.cutoff_date,
            "total_amount": self.total_amount,
            "total_cr": self.total_cr,
            "rows": [
                {
                    "tran_date": r.pdf.tran_date,
                    "post_date": r.pdf.post_date,
                    "description": r.pdf.description,
                    "amount": r.pdf.amount,
                    "note": r.note,
                    "mm_mentions": r.mm_mentions,
                }
                for r in self.rows
            ],
        }


_ACCOUNT_NAME = "Security Bank World"
_TYPE_EXPENSE = "Exp."

_MENTION_RE = re.compile(r"@\w+")
_PDF_TABLE_RE = re.compile(
    r"(\d{2}/\d{2}/\d{2})\s+(\d{2}/\d{2}/\d{2})\s+(.+?)\s+([\d,]+\.\d{2})(?:\s+(CR))?"
)


def _extract_mentions(description: str) -> str:
    return ", ".join(_MENTION_RE.findall(description or ""))


def _strip_mentions(description: str) -> str:
    return _MENTION_RE.sub("", description or "").strip()


def _parse_date_loose(value: str) -> datetime:
    for fmt in ("%m/%d/%y", "%m/%d/%Y"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    raise ValueError(f"Unrecognised date: {value!r}")


def _money_str(value: float) -> str:
    return f"{value:,.2f}"


# --- Money Manager loaders ---------------------------------------------------

def load_mm_from_google_sheet(cutoff_date: datetime) -> list[MMEntry]:
    client = get_gspread_client()
    sheet = client.open("Money Manager Data")
    worksheet = sheet.worksheet("Sheet1")
    data = worksheet.get_all_values()
    return _parse_mm_rows(data, cutoff_date)


def load_mm_from_xlsx(file_bytes: bytes, cutoff_date: datetime) -> list[MMEntry]:
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    data = [list(map(_xlsx_cell_to_str, row)) for row in ws.iter_rows(values_only=True)]
    return _parse_mm_rows(data, cutoff_date)


def _xlsx_cell_to_str(cell) -> str:
    if cell is None:
        return ""
    if isinstance(cell, datetime):
        return cell.strftime("%m/%d/%Y")
    return str(cell)


def _parse_mm_rows(data: list[list[str]], cutoff_date: datetime) -> list[MMEntry]:
    if not data:
        raise ValueError("Money Manager data appears empty.")

    header = data[0]
    try:
        period_idx = header.index("Period")
        accounts_idx = header.index("Accounts")
        amount_idx = header.index("Amount")
        type_idx = header.index("Income/Expense")
        description_idx = header.index("Description")
    except ValueError as exc:
        raise ValueError(f"Missing column in Money Manager data: {exc}") from exc

    cutoff_start = cutoff_date - relativedelta(months=1)

    entries: list[MMEntry] = []
    for row in data[1:]:
        max_idx = max(period_idx, amount_idx, accounts_idx, type_idx, description_idx)
        if len(row) <= max_idx:
            continue
        if row[accounts_idx] != _ACCOUNT_NAME:
            continue
        if row[type_idx] != _TYPE_EXPENSE:
            continue

        try:
            period_dt = _parse_date_loose(row[period_idx])
        except ValueError:
            continue
        if not (cutoff_start <= period_dt <= cutoff_date):
            continue

        # Original script noted MM data is consistently 1 day off from PDF.
        adjusted = (period_dt - relativedelta(days=1)).strftime("%m/%d/%y")
        try:
            amount_value = float(str(row[amount_idx]).replace(",", ""))
        except ValueError:
            continue

        entries.append(
            MMEntry(
                date=adjusted,
                amount=_money_str(amount_value),
                description=row[description_idx] or "",
                raw_period=period_dt.strftime("%m/%d/%Y"),
            )
        )
    return entries


# --- PDF extraction ----------------------------------------------------------

@dataclass
class PdfExtraction:
    rows: list[PdfRow]
    total_amount: float
    total_cr: float


def extract_pdf(file_bytes: bytes, password: str | None) -> PdfExtraction:
    reader = PdfReader(io.BytesIO(file_bytes))
    if reader.is_encrypted:
        if not password:
            raise ValueError("PDF is password protected. Provide a password.")
        try:
            ok = reader.decrypt(password)
        except Exception as exc:
            raise ValueError("Could not decrypt PDF.") from exc
        if not ok:
            raise ValueError("Incorrect PDF password.")

    rows: list[PdfRow] = []
    total_amount = 0.0
    total_cr = 0.0

    for page in reader.pages:
        text = page.extract_text() or ""
        for tran_date, post_date, description, amount, is_cr in _PDF_TABLE_RE.findall(text):
            amount_clean = amount.replace(",", "")
            try:
                amount_value = float(amount_clean)
            except ValueError:
                continue

            if is_cr:
                if not description.startswith("PAYMENT"):
                    total_cr += amount_value
                continue

            total_amount += amount_value
            rows.append(
                PdfRow(
                    tran_date=tran_date,
                    post_date=post_date,
                    description=description.strip(),
                    amount=_money_str(amount_value),
                )
            )

    return PdfExtraction(rows=rows, total_amount=total_amount, total_cr=total_cr)


# --- Matching / analysis ----------------------------------------------------

_MATCH_DAYS = 3
_INACCURATE_DAYS = 5
_INACCURATE_AMOUNT_TOLERANCE = 0.05  # 5%


def _exact_match(pdf_row: PdfRow, mm_pool: list[MMEntry]) -> MMEntry | None:
    pdf_date = _parse_date_loose(pdf_row.tran_date)
    for entry in mm_pool:
        try:
            mm_date = _parse_date_loose(entry.date)
        except ValueError:
            continue
        if (
            abs((pdf_date - mm_date).days) <= _MATCH_DAYS
            and entry.amount == pdf_row.amount
        ):
            return entry
    return None


def _nearby_candidates(pdf_row: PdfRow, mm_pool: Iterable[MMEntry]) -> list[MMEntry]:
    pdf_date = _parse_date_loose(pdf_row.tran_date)
    pdf_amount = pdf_row.amount_value
    out: list[MMEntry] = []
    for entry in mm_pool:
        try:
            mm_date = _parse_date_loose(entry.date)
        except ValueError:
            continue
        if abs((pdf_date - mm_date).days) > _INACCURATE_DAYS:
            continue
        if pdf_amount == 0:
            continue
        delta = abs(entry.amount_value - pdf_amount) / pdf_amount
        if delta <= _INACCURATE_AMOUNT_TOLERANCE:
            out.append(entry)
    return out


def analyse(
    *,
    pdf_extraction: PdfExtraction,
    mm_entries: list[MMEntry],
    sheet_name: str,
    billing_period: str,
    cutoff_date: datetime,
) -> AnalysisResult:
    mm_pool = list(mm_entries)
    analysed: list[AnalyzedRow] = []

    for pdf_row in pdf_extraction.rows:
        match = _exact_match(pdf_row, mm_pool)
        if match is not None:
            mm_pool.remove(match)
            note = _strip_mentions(match.description)
            mentions = _extract_mentions(match.description)
            analysed.append(
                AnalyzedRow(
                    pdf=pdf_row,
                    status="matched",
                    matched_mm=match,
                    note=note,
                    mm_mentions=mentions,
                )
            )
            continue

        candidates = _nearby_candidates(pdf_row, mm_pool)
        if candidates:
            consumed = candidates[0]
            mm_pool.remove(consumed)
            analysed.append(
                AnalyzedRow(
                    pdf=pdf_row,
                    status="inaccurate",
                    nearest_candidates=candidates,
                    note="[?] " + "; ".join(
                        f"MM {c.date} {c.amount} ({_strip_mentions(c.description)})"
                        for c in candidates
                    ),
                )
            )
            continue

        analysed.append(
            AnalyzedRow(pdf=pdf_row, status="missing", note="[!] not found in MM")
        )

    return AnalysisResult(
        sheet_name=sheet_name,
        billing_period=billing_period,
        cutoff_date=cutoff_date.strftime("%Y-%m-%d"),
        rows=analysed,
        unused_mm=mm_pool,
        total_amount=pdf_extraction.total_amount,
        total_cr=pdf_extraction.total_cr,
    )


# --- Sheet writer ------------------------------------------------------------

def push_to_google_sheet(payload: dict) -> str:
    """Write the analysed rows to a new tab on the destination sheet.

    `payload` matches AnalysisResult.serialise_for_confirm().
    Returns the spreadsheet URL.
    """
    client = get_gspread_client()
    sheet = client.open(payload["sheet_name"])

    billing_period = payload["billing_period"]
    if billing_period in [ws.title for ws in sheet.worksheets()]:
        raise ValueError(f"The billing period worksheet '{billing_period}' already exists.")

    rows = payload["rows"]
    new_ws = sheet.add_worksheet(title=billing_period, rows=str(len(rows) + 10), cols="20")

    worksheets = sheet.worksheets()
    sheet.reorder_worksheets([new_ws] + [ws for ws in worksheets if ws != new_ws])

    new_ws.append_row(
        ["Transaction", "Post date", "Merchant", "Amount", "Notes", "Shoulder", "C", "S"]
    )
    format_cell_range(new_ws, "1:1", CellFormat(textFormat=TextFormat(bold=True)))

    body = [
        [r["tran_date"], r["post_date"], r["description"], r["amount"], r["note"], r["mm_mentions"], "", ""]
        for r in rows
    ]
    body.sort(key=lambda x: _parse_date_loose(x[1]))

    if body:
        new_ws.append_rows(body, value_input_option="USER_ENTERED")

    total_amount = float(payload["total_amount"])
    total_cr = float(payload["total_cr"])
    grand_total = total_amount - total_cr

    new_ws.append_row(["", "", "TOTAL:", _money_str(total_amount)], value_input_option="USER_ENTERED")
    new_ws.append_row(["", "", "REIMBURSED TOTAL:", _money_str(total_cr)], value_input_option="USER_ENTERED")
    new_ws.append_row(["", "", "Grand Total:", _money_str(grand_total)], value_input_option="USER_ENTERED")

    return f"https://docs.google.com/spreadsheets/d/{sheet.id}"
