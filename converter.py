from pathlib import Path
import re

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


TRANSACTION_HEADER = [
    "GMT DDMM HH:MI",
    "TRAN DATE DDMM HH:MI",
    "PAN",
    "DEST",
    "MSG TYPE",
    "FUNC CODE",
    "ACT CODE",
    "TRAN TYPE",
    "TERMINAL ID",
    "RRN",
    "TRAN AMT",
    "INT FEE",
    "NET FEE",
]


def clean(value):
    if value is None:
        return ""

    return re.sub(r"\s+", " ", str(value)).strip()


def normalize(value):
    return clean(value).upper()


def detect_section(text):
    match = re.search(
        r"TRAN\s+TYPE\s*:\s*(.+)",
        text,
        re.IGNORECASE,
    )

    if match:
        return clean(match.group(1))

    return None


def is_transaction_header(row):
    if not row:
        return False

    text = " ".join(
        normalize(cell)
        for cell in row
    )

    required = [
        "GMT",
        "TRAN DATE",
        "PAN",
        "DEST",
        "MSG TYPE",
        "FUNC CODE",
        "ACT CODE",
        "TRAN TYPE",
        "TERMINAL ID",
        "RRN",
        "TRAN AMT",
        "INT FEE",
        "NET FEE",
    ]

    return all(item in text for item in required)


def find_transaction_table(tables):
    for table in tables:
        for row_index, row in enumerate(table):

            if is_transaction_header(row):
                return table, row_index

    return None, None


def extract_rows(table, header_index):
    rows = []

    for row in table[header_index + 1:]:

        if not row:
            continue

        row = [
            clean(cell)
            for cell in row
        ]

        if not any(row):
            continue

        if is_transaction_header(row):
            continue

        row = row[:13]

        if len(row) < 13:
            row.extend(
                [""] * (13 - len(row))
            )

        rows.append(row)

    return rows


def make_sheet_name(name, used_names):
    name = re.sub(
        r'[\\/*?:\[\]]',
        "-",
        name,
    )

    name = name.strip() or "Section"
    name = name[:31]

    original = name
    counter = 2

    while name in used_names:

        suffix = f" ({counter})"

        name = (
            original[:31 - len(suffix)]
            + suffix
        )

        counter += 1

    used_names.add(name)

    return name


def format_sheet(ws):
    ws.freeze_panes = "A2"

    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF",
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="1565C0",
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    for column in ws.columns:

        max_length = 0

        for cell in column:

            if cell.value is not None:
                max_length = max(
                    max_length,
                    len(str(cell.value)),
                )

        width = min(
            max(max_length + 2, 10),
            30,
        )

        letter = get_column_letter(
            column[0].column
        )

        ws.column_dimensions[
            letter
        ].width = width


def write_excel(sections, output_path):

    workbook = Workbook()

    workbook.remove(
        workbook.active
    )

    used_names = set()

    for section_name, rows in sections.items():

        sheet_name = make_sheet_name(
            section_name,
            used_names,
        )

        ws = workbook.create_sheet(
            sheet_name
        )

        ws.append(
            TRANSACTION_HEADER
        )

        for row in rows:
            ws.append(row)

        format_sheet(ws)

    workbook.save(output_path)


def convert_pdf_to_excel(
    pdf_path,
    output_path,
    progress_callback=None,
):

    pdf_path = Path(pdf_path)
    output_path = Path(output_path)

    if not pdf_path.exists():
        raise FileNotFoundError(
            f"PDF not found:\n{pdf_path}"
        )

    sections = {}
    current_section = None
    pages_processed = 0

    with pdfplumber.open(pdf_path) as pdf:

        total_pages = len(pdf.pages)

        for page_number, page in enumerate(
            pdf.pages,
            start=1,
        ):

            pages_processed = page_number

            # Tell the GUI about the current page.
            if progress_callback:
                progress_callback(
                    page_number,
                    total_pages,
                    "Reading PDF...",
                )

            page_text = (
                page.extract_text()
                or ""
            )

            # Detect a new section.
            detected_section = detect_section(
                page_text
            )

            if detected_section:

                current_section = (
                    detected_section
                )

                sections.setdefault(
                    current_section,
                    [],
                )

            # Extract tables.
            if progress_callback:
                progress_callback(
                    page_number,
                    total_pages,
                    "Reading tables...",
                )

            tables = page.extract_tables()

            table, header_index = (
                find_transaction_table(
                    tables
                )
            )

            if table is None:
                continue

            # Continuation page:
            # use the previous section.
            if current_section is None:
                continue

            rows = extract_rows(
                table,
                header_index,
            )

            sections[
                current_section
            ].extend(rows)

    if not sections:
        raise ValueError(
            "No transaction sections were found."
        )

    if progress_callback:
        progress_callback(
            pages_processed,
            total_pages,
            "Creating Excel file...",
        )

    write_excel(
        sections,
        output_path,
    )

    total_rows = sum(
        len(rows)
        for rows in sections.values()
    )

    # Purchase validation.
    purchase_rows = 0

    for section_name, rows in sections.items():

        if "PURCHASE" in section_name.upper():
            purchase_rows += len(rows)

    expected_purchase_rows = 3866

    purchase_valid = (
        purchase_rows
        == expected_purchase_rows
    )

    return {
        "pages": pages_processed,
        "sections": len(sections),
        "rows": total_rows,
        "purchase_rows": purchase_rows,
        "expected_purchase_rows": expected_purchase_rows,
        "purchase_valid": purchase_valid,
        "output": str(output_path),
    }


# Keep compatibility with app.py
extract_pdf_to_excel = convert_pdf_to_excel