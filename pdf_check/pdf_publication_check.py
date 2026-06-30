#!/usr/bin/env python3
"""Batch-check publication PDFs and create annotated review copies.

The script reads the authoritative page margins from a DOCX template, detects
large blank regions before the final PDF page, draws the correct margin box on
every page, and overlays detected whitespace in bright yellow at 70% opacity.

Outputs:
  <name>_checked.pdf   annotated review copy
  <name>_report.json  machine-readable per-PDF report
  batch_summary.csv   one-row summary per input PDF
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

import pdfplumber
from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
POINTS_PER_INCH = 72.0


@dataclass(frozen=True)
class Margins:
    top: float
    right: float
    bottom: float
    left: float
    page_width: float
    page_height: float


@dataclass(frozen=True)
class Region:
    x: float
    top: float
    width: float
    height: float
    columns: str

    @property
    def area_square_inches(self) -> float:
        return self.width * self.height / (POINTS_PER_INCH**2)


def _twips_to_points(value: str | None, default: float = 0.0) -> float:
    return float(value) / 20.0 if value is not None else default


def read_docx_sections(path: Path) -> list[Margins]:
    """Read section page sizes and margins directly from DOCX OOXML."""
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))

    sections: list[Margins] = []
    for section in root.iter(f"{W}sectPr"):
        size = section.find(f"{W}pgSz")
        margin = section.find(f"{W}pgMar")
        if size is None or margin is None:
            continue
        sections.append(
            Margins(
                top=_twips_to_points(margin.get(f"{W}top")),
                right=_twips_to_points(margin.get(f"{W}right")),
                bottom=_twips_to_points(margin.get(f"{W}bottom")),
                left=_twips_to_points(margin.get(f"{W}left")),
                page_width=_twips_to_points(size.get(f"{W}w")),
                page_height=_twips_to_points(size.get(f"{W}h")),
            )
        )
    if not sections:
        raise ValueError(f"No usable section margins found in {path}")
    return sections


def margins_for_page(
    sections: list[Margins], page_number: int, first_section_pages: int
) -> Margins:
    if len(sections) == 1:
        return sections[0]
    return sections[0] if page_number <= first_section_pages else sections[-1]


def find_pdftoppm(explicit: str | None) -> Path:
    candidates: list[Path] = []
    if explicit:
        candidates.append(Path(explicit))
    if os.environ.get("PDFTOPPM_PATH"):
        candidates.append(Path(os.environ["PDFTOPPM_PATH"]))

    found = shutil.which("pdftoppm.exe") or shutil.which("pdftoppm")
    if found:
        candidates.append(Path(found))

    # Codex bundled runtime: dependencies/python/python.exe is adjacent to
    # dependencies/native/poppler/Library/bin/pdftoppm.exe.
    python_dir = Path(sys.executable).resolve().parent
    dependencies_dir = python_dir.parent
    candidates.append(
        dependencies_dir / "native" / "poppler" / "Library" / "bin" / "pdftoppm.exe"
    )

    for candidate in candidates:
        if candidate.is_file() and candidate.suffix.lower() != ".cmd":
            return candidate.resolve()

    raise FileNotFoundError(
        "pdftoppm executable not found. Install Poppler or pass "
        "--pdftoppm PATH (PDFTOPPM_PATH is also supported)."
    )


def render_page_gray(
    pdftoppm: Path, pdf_path: Path, page_number: int, dpi: int, temp_dir: Path
) -> Image.Image:
    prefix = temp_dir / f"page-{page_number}"
    command = [
        str(pdftoppm),
        "-f",
        str(page_number),
        "-l",
        str(page_number),
        "-singlefile",
        "-gray",
        "-png",
        "-r",
        str(dpi),
        str(pdf_path),
        str(prefix),
    ]
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(
            f"pdftoppm failed for {pdf_path.name}, page {page_number}: "
            f"{result.stderr.strip()}"
        )
    image_path = prefix.with_suffix(".png")
    with Image.open(image_path) as rendered:
        return rendered.convert("L").copy()


def _expanded_occupied_rows(
    image: Image.Image,
    x0: int,
    x1: int,
    y0: int,
    y1: int,
    white_threshold: int,
    ink_pixels_per_row: int,
    padding_pixels: int,
) -> list[bool]:
    crop = image.crop((x0, y0, x1, y1))
    width, height = crop.size
    pixels = crop.load()
    occupied = [False] * height
    row_threshold = max(ink_pixels_per_row, int(width * 0.0005))

    for y in range(height):
        ink = 0
        for x in range(width):
            if pixels[x, y] < white_threshold:
                ink += 1
                if ink >= row_threshold:
                    occupied[y] = True
                    break

    if padding_pixels <= 0:
        return occupied

    expanded = [False] * height
    difference = [0] * (height + 1)
    for y, has_ink in enumerate(occupied):
        if not has_ink:
            continue
        start = max(0, y - padding_pixels)
        end = min(height, y + padding_pixels + 1)
        difference[start] += 1
        difference[end] -= 1
    active = 0
    for y in range(height):
        active += difference[y]
        expanded[y] = active > 0
    return expanded


def _state_runs(states: list[int]) -> Iterable[tuple[int, int, int]]:
    if not states:
        return
    start = 0
    state = states[0]
    for index in range(1, len(states) + 1):
        next_state = states[index] if index < len(states) else None
        if next_state != state:
            yield state, start, index
            start = index
            state = next_state


def detect_large_whitespace(
    image: Image.Image,
    page_width: float,
    page_height: float,
    margins: Margins,
    dpi: int,
    columns: int,
    gutter_points: float,
    min_height_points: float,
    white_threshold: int,
    ink_pixels_per_row: int,
    padding_points: float,
) -> list[Region]:
    """Detect large blank vertical runs in one- or two-column page panels."""
    scale_x = image.width / page_width
    scale_y = image.height / page_height
    x_left = max(0, round(margins.left * scale_x))
    x_right = min(image.width, round((page_width - margins.right) * scale_x))
    y_top = max(0, round(margins.top * scale_y))
    y_bottom = min(image.height, round((page_height - margins.bottom) * scale_y))
    padding_pixels = round(padding_points * scale_y)

    if x_right <= x_left or y_bottom <= y_top:
        return []

    panels: list[tuple[int, int]]
    if columns == 1:
        panels = [(x_left, x_right)]
    elif columns == 2:
        gutter_pixels = round(gutter_points * scale_x)
        center = (x_left + x_right) // 2
        panels = [
            (x_left, center - gutter_pixels // 2),
            (center + (gutter_pixels + 1) // 2, x_right),
        ]
    else:
        raise ValueError("--columns must be 1 or 2")

    occupied_by_panel = [
        _expanded_occupied_rows(
            image,
            panel_left,
            panel_right,
            y_top,
            y_bottom,
            white_threshold,
            ink_pixels_per_row,
            padding_pixels,
        )
        for panel_left, panel_right in panels
    ]

    if columns == 1:
        states = [0 if occupied_by_panel[0][y] else 3 for y in range(y_bottom - y_top)]
    else:
        states = []
        for y in range(y_bottom - y_top):
            state = 0
            if not occupied_by_panel[0][y]:
                state |= 1
            if not occupied_by_panel[1][y]:
                state |= 2
            states.append(state)

    min_height_pixels = round(min_height_points * scale_y)
    regions: list[Region] = []
    for state, start, end in _state_runs(states):
        if state == 0 or end - start < min_height_pixels:
            continue

        top_points = (y_top + start) / scale_y
        height_points = (end - start) / scale_y
        if state == 3 or columns == 1:
            region_x = margins.left
            region_width = page_width - margins.left - margins.right
            label = "full"
        elif state == 1:
            region_x = panels[0][0] / scale_x
            region_width = (panels[0][1] - panels[0][0]) / scale_x
            label = "left"
        else:
            region_x = panels[1][0] / scale_x
            region_width = (panels[1][1] - panels[1][0]) / scale_x
            label = "right"

        regions.append(
            Region(
                x=round(region_x, 2),
                top=round(top_points, 2),
                width=round(region_width, 2),
                height=round(height_points, 2),
                columns=label,
            )
        )
    return regions


def inspect_content_geometry(
    page: pdfplumber.page.Page,
    page_width: float,
    margins: Margins,
    tolerance: float,
) -> dict:
    """Inspect text and images for horizontal overflow past correct margins."""
    objects = list(page.chars) + list(page.images)
    if not objects:
        return {
            "status": "NO_CONTENT",
            "observed_bounds_points": None,
            "left_overflow_points": 0.0,
            "right_overflow_points": 0.0,
            "objects_outside_horizontal_margins": 0,
        }

    left_expected = margins.left
    right_expected = page_width - margins.right
    min_x = min(float(obj["x0"]) for obj in objects)
    max_x = max(float(obj["x1"]) for obj in objects)
    min_top = min(float(obj.get("top", 0.0)) for obj in objects)
    max_bottom = max(float(obj.get("bottom", 0.0)) for obj in objects)
    left_overflow = max(0.0, left_expected - min_x)
    right_overflow = max(0.0, max_x - right_expected)
    outside_count = sum(
        1
        for obj in objects
        if float(obj["x0"]) < left_expected - tolerance
        or float(obj["x1"]) > right_expected + tolerance
    )
    status = (
        "FAIL"
        if left_overflow > tolerance or right_overflow > tolerance
        else "PASS"
    )
    return {
        "status": status,
        "observed_bounds_points": {
            "left": round(min_x, 2),
            "right": round(max_x, 2),
            "top": round(min_top, 2),
            "bottom": round(max_bottom, 2),
        },
        "left_overflow_points": round(left_overflow, 2),
        "right_overflow_points": round(right_overflow, 2),
        "objects_outside_horizontal_margins": outside_count,
    }


def make_overlay(
    page_width: float,
    page_height: float,
    margins: Margins,
    regions: list[Region],
) -> PdfReader:
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=(page_width, page_height))

    # Correct publication margin box.
    pdf.saveState()
    pdf.setStrokeColorRGB(0.0, 0.35, 1.0)
    pdf.setLineWidth(1.2)
    pdf.setDash(5, 3)
    pdf.rect(
        margins.left,
        margins.bottom,
        page_width - margins.left - margins.right,
        page_height - margins.top - margins.bottom,
        stroke=1,
        fill=0,
    )
    pdf.restoreState()

    # Bright yellow, exactly 70% opacity, as requested.
    for region in regions:
        y = page_height - region.top - region.height
        pdf.saveState()
        if hasattr(pdf, "setFillAlpha"):
            pdf.setFillAlpha(0.70)
        pdf.setFillColorRGB(1.0, 1.0, 0.0)
        pdf.rect(region.x, y, region.width, region.height, stroke=0, fill=1)
        pdf.restoreState()

        pdf.saveState()
        pdf.setFillColorRGB(0.15, 0.15, 0.0)
        pdf.setFont("Helvetica-Bold", 7)
        pdf.drawString(region.x + 4, y + region.height - 10, "LARGE WHITESPACE")
        pdf.restoreState()

    pdf.showPage()
    pdf.save()
    buffer.seek(0)
    return PdfReader(buffer)


def annotate_pdf(
    source_path: Path,
    output_path: Path,
    sections: list[Margins],
    regions_by_page: list[list[Region]],
    first_section_pages: int,
) -> None:
    reader = PdfReader(source_path)
    writer = PdfWriter()
    for page_number, page in enumerate(reader.pages, start=1):
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        margins = margins_for_page(sections, page_number, first_section_pages)
        overlay = make_overlay(
            width, height, margins, regions_by_page[page_number - 1]
        ).pages[0]
        page.merge_page(overlay, over=True)
        writer.add_page(page)

    metadata = dict(reader.metadata or {})
    metadata["/PublicationCheck"] = "Correct margins boxed; large whitespace yellow at 70% opacity"
    writer.add_metadata({str(key): str(value) for key, value in metadata.items()})
    with output_path.open("wb") as handle:
        writer.write(handle)


def check_pdf(
    pdf_path: Path,
    output_dir: Path,
    sections: list[Margins],
    pdftoppm: Path,
    args: argparse.Namespace,
) -> dict:
    reader = PdfReader(pdf_path)
    page_count = len(reader.pages)
    regions_by_page: list[list[Region]] = [[] for _ in range(page_count)]
    page_reports: list[dict] = []

    with pdfplumber.open(pdf_path) as plumber_pdf, tempfile.TemporaryDirectory(
        prefix="pdf-publication-check-"
    ) as temp:
        temp_dir = Path(temp)
        for page_number, (pdf_page, geometry_page) in enumerate(
            zip(reader.pages, plumber_pdf.pages), start=1
        ):
            width = float(pdf_page.mediabox.width)
            height = float(pdf_page.mediabox.height)
            margins = margins_for_page(
                sections, page_number, args.first_section_pages
            )

            page_size_status = (
                "PASS"
                if abs(width - margins.page_width) <= args.page_size_tolerance
                and abs(height - margins.page_height) <= args.page_size_tolerance
                else "FAIL"
            )
            margin_check = inspect_content_geometry(
                geometry_page, width, margins, args.margin_tolerance
            )
            if page_size_status == "FAIL":
                margin_check["status"] = "FAIL"

            # Large trailing whitespace on the final page is normal at the end
            # of a paper, so only earlier pages are analyzed by default.
            should_analyze = (
                page_number < page_count or args.include_final_page
            )
            if should_analyze:
                image = render_page_gray(
                    pdftoppm, pdf_path, page_number, args.dpi, temp_dir
                )
                regions_by_page[page_number - 1] = detect_large_whitespace(
                    image=image,
                    page_width=width,
                    page_height=height,
                    margins=margins,
                    dpi=args.dpi,
                    columns=args.columns,
                    gutter_points=args.gutter,
                    min_height_points=args.min_whitespace_height,
                    white_threshold=args.white_threshold,
                    ink_pixels_per_row=args.ink_pixels_per_row,
                    padding_points=args.ink_padding,
                )

            page_reports.append(
                {
                    "page": page_number,
                    "page_size_points": {
                        "width": round(width, 2),
                        "height": round(height, 2),
                    },
                    "page_size_status": page_size_status,
                    "expected_margins_points": {
                        "top": round(margins.top, 2),
                        "right": round(margins.right, 2),
                        "bottom": round(margins.bottom, 2),
                        "left": round(margins.left, 2),
                    },
                    "margin_check": margin_check,
                    "whitespace_analyzed": should_analyze,
                    "whitespace_regions": [
                        {
                            **asdict(region),
                            "area_square_inches": round(
                                region.area_square_inches, 2
                            ),
                        }
                        for region in regions_by_page[page_number - 1]
                    ],
                }
            )

    annotated_name = f"{pdf_path.stem}_checked.pdf"
    report_name = f"{pdf_path.stem}_report.json"
    annotate_pdf(
        pdf_path,
        output_dir / annotated_name,
        sections,
        regions_by_page,
        args.first_section_pages,
    )

    whitespace_count = sum(len(regions) for regions in regions_by_page)
    margin_fail_pages = [
        report["page"]
        for report in page_reports
        if report["margin_check"]["status"] == "FAIL"
    ]
    report = {
        "source_pdf": str(pdf_path.resolve()),
        "annotated_pdf": str((output_dir / annotated_name).resolve()),
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "summary": {
            "page_count": page_count,
            "margin_status": "FAIL" if margin_fail_pages else "PASS",
            "margin_fail_pages": margin_fail_pages,
            "large_whitespace_region_count": whitespace_count,
            "whitespace_status": "REVIEW" if whitespace_count else "PASS",
        },
        "settings": {
            "final_page_excluded": not args.include_final_page,
            "columns": args.columns,
            "gutter_points": args.gutter,
            "minimum_whitespace_height_points": args.min_whitespace_height,
            "minimum_whitespace_height_inches": round(
                args.min_whitespace_height / POINTS_PER_INCH, 3
            ),
            "yellow_opacity": 0.70,
            "margin_tolerance_points": args.margin_tolerance,
            "dpi": args.dpi,
        },
        "pages": page_reports,
    }
    with (output_dir / report_name).open("w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2)
        handle.write("\n")
    return report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Batch-check PDFs for publication margins and large whitespace, "
            "then create annotated PDF copies and JSON reports."
        )
    )
    parser.add_argument(
        "input_dir",
        nargs="?",
        type=Path,
        default=Path("input"),
        help="Directory containing papers to check (default: input)",
    )
    parser.add_argument(
        "--template-docx",
        type=Path,
        default=Path("template.docx"),
        help="Authoritative DOCX template (default: template.docx)",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="Output directory (default: output)",
    )
    parser.add_argument("--pattern", default="*.pdf", help="Input glob pattern")
    parser.add_argument(
        "--first-section-pages",
        type=int,
        default=1,
        help="Pages using the DOCX's first section margins (default: 1)",
    )
    parser.add_argument(
        "--columns",
        type=int,
        choices=(1, 2),
        default=2,
        help="Expected publication column count (default: 2)",
    )
    parser.add_argument(
        "--gutter",
        type=float,
        default=18.0,
        help="Column gutter width in points (default: 18)",
    )
    parser.add_argument(
        "--min-whitespace-height",
        type=float,
        default=108.0,
        help="Minimum highlighted blank height in points (default: 108 = 1.5 in)",
    )
    parser.add_argument(
        "--include-final-page",
        action="store_true",
        help="Also flag large whitespace on the paper's final page",
    )
    parser.add_argument("--dpi", type=int, default=144, help="Raster analysis DPI")
    parser.add_argument(
        "--white-threshold",
        type=int,
        default=245,
        help="Pixels below this grayscale value count as ink (default: 245)",
    )
    parser.add_argument(
        "--ink-pixels-per-row",
        type=int,
        default=2,
        help="Minimum ink pixels that make a panel row occupied (default: 2)",
    )
    parser.add_argument(
        "--ink-padding",
        type=float,
        default=6.0,
        help="Vertical padding around detected ink, in points (default: 6)",
    )
    parser.add_argument(
        "--margin-tolerance",
        type=float,
        default=6.0,
        help="Allowed horizontal margin overflow in points (default: 6)",
    )
    parser.add_argument(
        "--page-size-tolerance",
        type=float,
        default=1.0,
        help="Allowed page-size difference in points (default: 1)",
    )
    parser.add_argument(
        "--pdftoppm",
        help="Path to Poppler pdftoppm executable",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    input_dir = args.input_dir.resolve()
    template_docx = args.template_docx.resolve()
    output_dir = args.output_dir.resolve()

    if not input_dir.is_dir():
        raise SystemExit(
            f"Input directory not found: {input_dir}. "
            "Create the folder and place the PDF papers in it."
        )
    if not template_docx.is_file():
        raise SystemExit(f"Template DOCX not found: {template_docx}")
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_paths = sorted(
        path
        for path in input_dir.glob(args.pattern)
        if path.is_file()
        and not path.stem.endswith("_checked")
        and not (
            output_dir != input_dir
            and output_dir in path.resolve().parents
        )
    )
    if not pdf_paths:
        raise SystemExit(
            f"No PDFs matching {args.pattern!r} found in {input_dir}"
        )

    sections = read_docx_sections(template_docx)
    pdftoppm = find_pdftoppm(args.pdftoppm)
    reports: list[dict] = []
    for pdf_path in pdf_paths:
        print(f"Checking {pdf_path.name} ...", flush=True)
        reports.append(
            check_pdf(pdf_path, output_dir, sections, pdftoppm, args)
        )

    summary_path = output_dir / "batch_summary.csv"
    with summary_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "source_pdf",
                "annotated_pdf",
                "page_count",
                "margin_status",
                "margin_fail_pages",
                "whitespace_status",
                "large_whitespace_region_count",
            ],
        )
        writer.writeheader()
        for report in reports:
            summary = report["summary"]
            writer.writerow(
                {
                    "source_pdf": report["source_pdf"],
                    "annotated_pdf": report["annotated_pdf"],
                    "page_count": summary["page_count"],
                    "margin_status": summary["margin_status"],
                    "margin_fail_pages": ",".join(
                        str(page) for page in summary["margin_fail_pages"]
                    ),
                    "whitespace_status": summary["whitespace_status"],
                    "large_whitespace_region_count": summary[
                        "large_whitespace_region_count"
                    ],
                }
            )

    print(f"Finished: {len(reports)} PDF(s). Output: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
