import sys
import re
from pathlib import Path
from datetime import datetime, date

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    KeepTogether,
    Image,
)


SECTION_PATTERN = re.compile(r".+\s-\s.+")
DATE_PATTERNS = [
    re.compile(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b"),
    re.compile(r"\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b"),
    re.compile(
        r"\b(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s+"
        r"\d{1,2}\s+"
        r"(?:January|February|March|April|May|June|July|August|September|October|November|December)"
        r"\s+\d{4}\b",
        re.IGNORECASE,
    ),
    re.compile(
        r"\b\d{1,2}\s+"
        r"(?:January|February|March|April|May|June|July|August|September|October|November|December)"
        r"\s+\d{4}\b",
        re.IGNORECASE,
    ),
]


def is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    return str(value).strip() == ""


def format_nice_date(dt: datetime) -> str:
    if sys.platform == "win32":
        return dt.strftime("%A, %#d %B %Y")
    return dt.strftime("%A, %-d %B %Y")


def auto_race_name_from_filename(excel_path: str) -> str:
    stem = Path(excel_path).stem
    stem = stem.replace("_", " ").strip()
    stem = re.sub(r"\(\d+\)$", "", stem).strip()
    stem = re.sub(r"\s+", " ", stem)
    return stem if stem else "Race Results"


def detect_date_in_workbook(excel_path: str) -> str:
    raw = pd.read_excel(excel_path, header=None)
    raw = raw.where(pd.notna(raw), None)

    for _, row in raw.iterrows():
        for value in row.tolist():
            if isinstance(value, datetime):
                return format_nice_date(value)
            if isinstance(value, date):
                return format_nice_date(datetime.combine(value, datetime.min.time()))

    for _, row in raw.iterrows():
        for value in row.tolist():
            if is_blank(value):
                continue
            text = str(value).strip()
            for pattern in DATE_PATTERNS:
                match = pattern.search(text)
                if match:
                    return match.group(0)

    return ""


def auto_race_date(excel_path: str) -> str:
    detected = detect_date_in_workbook(excel_path)
    if detected:
        return detected

    mtime = Path(excel_path).stat().st_mtime
    return format_nice_date(datetime.fromtimestamp(mtime))


def looks_like_section_title(row_values) -> bool:
    non_blank = [str(v).strip() for v in row_values if not is_blank(v)]
    if len(non_blank) != 1:
        return False
    return bool(SECTION_PATTERN.match(non_blank[0]))


def looks_like_header_row(row_values) -> bool:
    values = [str(v).strip().lower() for v in row_values if not is_blank(v)]
    return {"place", "bib", "name"}.issubset(set(values))


def normalise_header(row_values):
    headers = []
    for v in row_values:
        headers.append("" if is_blank(v) else str(v).strip())
    return headers


def extract_sections_from_excel(excel_path: str):
    raw = pd.read_excel(excel_path, header=None)
    raw = raw.where(pd.notna(raw), None)

    sections = []
    current_title = None
    current_headers = None
    current_rows = []

    for _, row in raw.iterrows():
        row_values = row.tolist()

        if all(is_blank(v) for v in row_values):
            continue

        if looks_like_section_title(row_values):
            if current_title and current_headers and current_rows:
                sections.append((current_title, current_headers, current_rows))
            current_title = [str(v).strip() for v in row_values if not is_blank(v)][0]
            current_headers = None
            current_rows = []
            continue

        if current_title and looks_like_header_row(row_values):
            current_headers = normalise_header(row_values)
            continue

        if current_title and current_headers:
            trimmed = row_values[: len(current_headers)]
            if not all(is_blank(v) for v in trimmed):
                current_rows.append(trimmed)

    if current_title and current_headers and current_rows:
        sections.append((current_title, current_headers, current_rows))

    return sections


def map_columns(headers):
    mapped = {}
    for i, h in enumerate(headers):
        key = h.strip().lower()

        if key == "place":
            mapped["Place"] = i
        elif key == "bib":
            mapped["Bib"] = i
        elif key == "name":
            mapped["Name"] = i
        elif key == "category":
            mapped["Category"] = i
        elif key == "gender":
            mapped["Gender"] = i
        elif key in {"finish time", "time"}:
            mapped["Time"] = i

    return mapped


def row_is_dnf_or_dns(headers, row) -> bool:
    """
    Return True if a Webscorer row represents a DNF or DNS result.

    These usually appear in the Time column, but this checks the mapped output
    columns as a fallback so that rows are still excluded if Webscorer changes
    where the status text appears.
    """
    colmap = map_columns(headers)
    status_values = {"DNF", "DNS"}

    # Primary check: Time / Finish Time column
    if "Time" in colmap:
        idx = colmap["Time"]
        value = row[idx] if idx < len(row) else ""
        if str(value).strip().upper() in status_values:
            return True

    # Fallback check: any mapped output cell exactly equal to DNF or DNS
    for idx in colmap.values():
        value = row[idx] if idx < len(row) else ""
        if str(value).strip().upper() in status_values:
            return True

    return False


def build_table_data(headers, rows):
    colmap = map_columns(headers)

    preferred_order = ["Place", "Bib", "Name", "Category", "Gender", "Time"]
    present_cols = [c for c in preferred_order if c in colmap]

    output = [present_cols]

    for row in rows:
        if row_is_dnf_or_dns(headers, row):
            continue

        out_row = []
        for col in present_cols:
            idx = colmap[col]
            value = row[idx] if idx < len(row) else ""
            text = "" if is_blank(value) else str(value).strip()

            if col == "Gender":
                if text.lower() == "female":
                    text = "F"
                elif text.lower() == "male":
                    text = "M"

            out_row.append(text)
        output.append(out_row)

    return output


def add_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.grey)
    canvas.drawString(14 * mm, 8 * mm, "Sri Chinmoy Races")
    canvas.drawRightString(200 * mm, 8 * mm, f"Page {doc.page}")
    canvas.restoreState()


def make_table(table_data):
    num_cols = len(table_data[0])
    col_widths = []

    for col in table_data[0]:
        if col == "Place":
            col_widths.append(16 * mm)
        elif col == "Bib":
            col_widths.append(16 * mm)
        elif col == "Name":
            col_widths.append(48 * mm)
        elif col == "Category":
            col_widths.append(55 * mm)
        elif col == "Gender":
            col_widths.append(16 * mm)
        elif col == "Time":
            col_widths.append(24 * mm)
        else:
            col_widths.append((180 / max(1, num_cols)) * mm)

    table = Table(table_data, colWidths=col_widths, repeatRows=1)

    style_commands = [
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 5),
        ("TOPPADDING", (0, 0), (-1, 0), 2),
        ("LINEBELOW", (0, 0), (-1, 0), 0.6, colors.black),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 8.5),
        ("TEXTCOLOR", (0, 1), (-1, -1), colors.black),
        ("TOPPADDING", (0, 1), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 2),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]

    if "Gender" in table_data[0]:
        gender_idx = table_data[0].index("Gender")
        style_commands.append(("ALIGN", (gender_idx, 0), (gender_idx, -1), "CENTER"))

    if "Time" in table_data[0]:
        time_idx = table_data[0].index("Time")
        style_commands.append(("ALIGN", (time_idx, 0), (time_idx, -1), "RIGHT"))

    for r in range(1, len(table_data)):
        if r % 2 == 0:
            style_commands.append(("BACKGROUND", (0, r), (-1, r), colors.whitesmoke))

    table.setStyle(TableStyle(style_commands))
    return table


def split_title_parts(title: str):
    return [p.strip() for p in title.split(" - ")]


def get_distance_name(title: str) -> str:
    parts = split_title_parts(title)
    return parts[0] if parts else title.strip()


def clean_section_title(title: str) -> str:
    """
    Avoid headings like '4km - 4km Female 60-69 - Female'.

    Webscorer sometimes includes the race distance both as the section prefix and
    again at the start of the category name. This removes only the repeated
    distance at the start of later title parts.
    """
    parts = split_title_parts(title)
    if len(parts) < 2:
        return title.strip()

    distance = parts[0].strip()
    distance_pattern = re.compile(rf"^{re.escape(distance)}\s+", re.IGNORECASE)

    cleaned_parts = [distance]
    for part in parts[1:]:
        cleaned_part = distance_pattern.sub("", part.strip()).strip()
        cleaned_parts.append(cleaned_part)

    return " - ".join(part for part in cleaned_parts if part)


def build_distance_order(sections):
    """Preserve the order in which distances first appear in the workbook."""
    order = {}
    next_index = 0

    for title, headers, rows in sections:
        distance = get_distance_name(title)
        if distance not in order:
            order[distance] = next_index
            next_index += 1

    return order


def classify_section_type(title: str) -> int:
    """Lower number = earlier in the PDF."""
    parts = split_title_parts(title)

    if len(parts) >= 2:
        second = parts[1].strip().lower()

        if second == "overall":
            return 1
        if second == "female":
            return 2
        if second == "male":
            return 3
        if second in {"non-binary", "nonbinary"}:
            return 4

    return 10


def section_sort_key(title: str, distance_order_map: dict):
    distance = get_distance_name(title)
    return (
        distance_order_map.get(distance, 999),
        classify_section_type(title),
        title.lower(),
    )


def create_pdf(excel_file, output_pdf, race_name=None, race_date=None):
    sections = extract_sections_from_excel(excel_file)

    if not sections:
        raise ValueError("No result sections were found in the Excel file.")

    distance_order_map = build_distance_order(sections)
    sections = sorted(sections, key=lambda s: section_sort_key(s[0], distance_order_map))

    if not race_name:
        race_name = auto_race_name_from_filename(excel_file)

    if not race_date:
        race_date = auto_race_date(excel_file)

    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "RaceTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=19,
        alignment=TA_LEFT,
        spaceAfter=4,
    )

    subtitle_style = ParagraphStyle(
        "RaceDate",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=12,
        textColor=colors.grey,
        spaceAfter=10,
    )

    section_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=13,
        alignment=TA_LEFT,
        spaceBefore=6,
        spaceAfter=4,
        textColor=colors.black,
    )
    logo_path = Path(__file__).parent / "sri_chinmoy_races_logo.png"

    header_items = [
        Paragraph(race_name, title_style),
        Paragraph(race_date, subtitle_style) if race_date else Spacer(1, 4),
    ]

    if logo_path.exists():
        logo_width = 42 * mm
        results_width = 175 * mm

        logo = Image(str(logo_path))
        logo.drawWidth = logo_width
        logo.drawHeight = logo_width

        header_table = Table(
            [[header_items, logo]],
            colWidths=[results_width - logo_width, logo_width],
            hAlign="LEFT",
        )

        header_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (1, 0), (1, 0), "RIGHT"),
            ("BOX", (0, 0), (-1, -1), 0, colors.white),
            ("INNERGRID", (0, 0), (-1, -1), 0, colors.white),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))

        elements = [header_table, Spacer(1, 8)]
    else:
        elements = header_items

    first_section = True

    for title, headers, rows in sections:
        table_data = build_table_data(headers, rows)
        if len(table_data) <= 1:
            continue

        table = make_table(table_data)

        if not first_section:
            elements.append(Spacer(1, 8))
        first_section = False

        elements.append(Paragraph(clean_section_title(title), section_style))
        elements.append(Spacer(1, 2))
        elements.append(table)

    doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)

    return output_pdf

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print('Usage: python race_results_to_pdf.py "input.xlsx" "output.pdf" ["Race name"] ["Race date"]')
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    race_name = sys.argv[3] if len(sys.argv) > 3 else None
    race_date = sys.argv[4] if len(sys.argv) > 4 else None

    if not Path(input_file).exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    create_pdf(input_file, output_file, race_name, race_date)
    print(f"PDF created: {output_file}")
