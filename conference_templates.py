from __future__ import annotations

import html
import io
import re
import unicodedata
import zipfile
from dataclasses import dataclass
from datetime import date
from pathlib import Path
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
HEADER_PATH = "word/header2.xml"
LATEX_MAIN_PATH = "main.tex"
DEFAULT_CONFERENCE_KEY = "escape-37-2027"
LATEX_CONFERENCE_KEY = "latex"
CUSTOM_CONFERENCE_KEY = "custom"

_PARAGRAPH_RE = re.compile(rb"<w:p(?=[\s>])[^>]*>.*?</w:p>", re.DOTALL)
_TEXT_RE = re.compile(rb"(<w:t(?=[\s>])[^>]*>)(.*?)(</w:t>)", re.DOTALL)
_SPACE_RE = re.compile(r"\s+")
_SLUG_RE = re.compile(r"[^a-z0-9]+")
_MONTH_NAMES = (
    "",
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)


@dataclass(frozen=True)
class ConferenceInfo:
    """The two editable conference lines in the PSE Press first-page header."""

    name: str
    location: str


CONFERENCE_PRESETS: dict[str, ConferenceInfo] = {
    "escape-37-2027": ConferenceInfo(
        name="ESCAPE 37 - European Symposium on Computer Aided Process Engineering",
        location="Trondheim, Norway, 6-9 June 2027",
    ),
    "pse-2027": ConferenceInfo(
        name="PSE 2027 – Process Systems Engineering",
        location="Mexico City, Mexico, 13-17 June 2027",
    ),
    "focapo-cpc-2027": ConferenceInfo(
        name="FOCAPO-CPC 2027",
        location="Tucson, Arizona, USA, 10-14 January 2027",
    ),
}


def normalize_header_text(value: str) -> str:
    return _SPACE_RE.sub(" ", value).strip()


def validate_conference(conference: ConferenceInfo) -> ConferenceInfo:
    name = normalize_header_text(conference.name)
    location = normalize_header_text(conference.location)
    if not name:
        raise ValueError("Conference name is required.")
    if not location:
        raise ValueError("Conference location and dates are required.")
    return ConferenceInfo(name=name, location=location)


def format_date_range(start_date: date, end_date: date) -> str:
    if end_date < start_date:
        raise ValueError("Conference end date must be on or after the start date.")
    if start_date == end_date:
        return f"{start_date.day} {_MONTH_NAMES[start_date.month]} {start_date.year}"
    if start_date.year == end_date.year and start_date.month == end_date.month:
        return f"{start_date.day}-{end_date.day} {_MONTH_NAMES[start_date.month]} {start_date.year}"
    if start_date.year == end_date.year:
        return (
            f"{start_date.day} {_MONTH_NAMES[start_date.month]}-"
            f"{end_date.day} {_MONTH_NAMES[end_date.month]} {start_date.year}"
        )
    return (
        f"{start_date.day} {_MONTH_NAMES[start_date.month]} {start_date.year}-"
        f"{end_date.day} {_MONTH_NAMES[end_date.month]} {end_date.year}"
    )


def conference_from_form(
    *,
    name: str,
    city: str,
    region: str,
    country: str,
    start_date: date,
    end_date: date,
) -> ConferenceInfo:
    normalized_city = normalize_header_text(city)
    normalized_region = normalize_header_text(region)
    normalized_country = normalize_header_text(country)
    missing = [
        label
        for label, value in (
            ("Conference name", normalize_header_text(name)),
            ("City", normalized_city),
            ("Country", normalized_country),
        )
        if not value
    ]
    if missing:
        raise ValueError(f"Required field(s) missing: {', '.join(missing)}.")
    location_parts = [normalized_city]
    if normalized_region:
        location_parts.append(normalized_region)
    location_parts.append(normalized_country)
    location_parts.append(format_date_range(start_date, end_date))
    return validate_conference(
        ConferenceInfo(
            name=name,
            location=", ".join(location_parts),
        )
    )


def conference_slug(name: str, *, max_length: int = 60) -> str:
    ascii_name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
    slug = _SLUG_RE.sub("-", ascii_name.lower()).strip("-")
    slug = slug[:max_length].rstrip("-")
    return slug or "conference"


def latex_escape_text(value: str) -> str:
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    return "".join(replacements.get(character, character) for character in value)


def resolve_conference_selection(
    selection: str,
    *,
    custom_name: str = "",
    custom_location: str = "",
) -> ConferenceInfo | None:
    normalized_selection = selection.strip().lower()
    if normalized_selection == LATEX_CONFERENCE_KEY:
        return None
    if normalized_selection == CUSTOM_CONFERENCE_KEY:
        return validate_conference(ConferenceInfo(custom_name, custom_location))
    try:
        return CONFERENCE_PRESETS[normalized_selection]
    except KeyError as exc:
        choices = ", ".join((*CONFERENCE_PRESETS, LATEX_CONFERENCE_KEY, CUSTOM_CONFERENCE_KEY))
        raise ValueError(f"Unknown conference selection {selection!r}. Choose one of: {choices}.") from exc


def _header_paragraph_texts(header_xml: bytes) -> list[str]:
    root = ET.fromstring(header_xml)
    return [
        "".join(text.text or "" for text in paragraph.iter(f"{{{W_NS}}}t"))
        for paragraph in root.iter(f"{{{W_NS}}}p")
    ]


def _replace_paragraph_text(paragraph_xml: bytes, value: str) -> bytes:
    matches = list(_TEXT_RE.finditer(paragraph_xml))
    if not matches:
        raise ValueError("The conference header paragraph does not contain an editable Word text node.")
    escaped_value = html.escape(value, quote=False).encode("utf-8")
    pieces: list[bytes] = []
    cursor = 0
    for index, match in enumerate(matches):
        pieces.append(paragraph_xml[cursor : match.start()])
        pieces.append(match.group(1))
        if index == 0:
            pieces.append(escaped_value)
        pieces.append(match.group(3))
        cursor = match.end()
    pieces.append(paragraph_xml[cursor:])
    return b"".join(pieces)


def update_conference_header_xml(header_xml: bytes, conference: ConferenceInfo) -> bytes:
    conference = validate_conference(conference)
    paragraphs = list(_PARAGRAPH_RE.finditer(header_xml))
    paragraph_texts = _header_paragraph_texts(header_xml)
    if len(paragraphs) < 5 or len(paragraph_texts) < 5:
        raise ValueError("The Word template does not contain the expected five-paragraph first-page header.")
    if normalize_header_text(paragraph_texts[1]) != "Original Research Article":
        raise ValueError("The Word template first-page header has an unexpected article-type slot.")
    if normalize_header_text(paragraph_texts[2]) != "Peer Reviewed Conference Proceeding":
        raise ValueError("The Word template first-page header has an unexpected review-type slot.")

    replacements = ((3, conference.name), (4, conference.location))
    output = header_xml
    for paragraph_index, value in reversed(replacements):
        match = paragraphs[paragraph_index]
        updated_paragraph = _replace_paragraph_text(match.group(0), value)
        output = output[: match.start()] + updated_paragraph + output[match.end() :]
    return output


def update_conference_header_entries(entries: dict[str, bytes], conference: ConferenceInfo) -> None:
    if HEADER_PATH not in entries:
        raise ValueError(f"The Word template is missing {HEADER_PATH}.")
    entries[HEADER_PATH] = update_conference_header_xml(entries[HEADER_PATH], conference)


def generate_word_template(template_bytes: bytes, conference: ConferenceInfo) -> bytes:
    source = io.BytesIO(template_bytes)
    output = io.BytesIO()
    with zipfile.ZipFile(source, "r") as input_archive:
        if HEADER_PATH not in input_archive.namelist():
            raise ValueError(f"The Word template is missing {HEADER_PATH}.")
        with zipfile.ZipFile(output, "w") as output_archive:
            output_archive.comment = input_archive.comment
            for info in input_archive.infolist():
                data = input_archive.read(info.filename)
                if info.filename == HEADER_PATH:
                    data = update_conference_header_xml(data, conference)
                output_archive.writestr(info, data)
    return output.getvalue()


def custom_conference_latex(conference: ConferenceInfo) -> str:
    conference = validate_conference(conference)
    return (
        f"\\PSESetConference{{{latex_escape_text(conference.name)}}}"
        f"{{{latex_escape_text(conference.location)}}}"
    )


def update_latex_main_conference(tex_source: str, conference: ConferenceInfo) -> str:
    replacement = custom_conference_latex(conference)
    default_command = rf"\PSESelectConference{{{DEFAULT_CONFERENCE_KEY}}}"
    if tex_source.count(default_command) != 1:
        raise ValueError(
            "The LaTeX template must contain exactly one active default conference selector "
            f"({default_command})."
        )
    return tex_source.replace(default_command, replacement, 1)


def generate_latex_template_archive(template_zip_bytes: bytes, conference: ConferenceInfo) -> bytes:
    source = io.BytesIO(template_zip_bytes)
    output = io.BytesIO()
    with zipfile.ZipFile(source, "r") as input_archive:
        if LATEX_MAIN_PATH not in input_archive.namelist():
            raise ValueError(f"The LaTeX template archive is missing {LATEX_MAIN_PATH}.")
        with zipfile.ZipFile(output, "w") as output_archive:
            output_archive.comment = input_archive.comment
            for info in input_archive.infolist():
                data = input_archive.read(info.filename)
                if info.filename == LATEX_MAIN_PATH:
                    tex_source = data.decode("utf-8")
                    data = update_latex_main_conference(tex_source, conference).encode("utf-8")
                output_archive.writestr(info, data)
    return output.getvalue()


def generate_template_filenames(conference: ConferenceInfo) -> tuple[str, str]:
    slug = conference_slug(validate_conference(conference).name)
    return f"{slug}-word-template.docx", f"{slug}-latex-template.zip"


def load_template_bytes(path: Path) -> bytes:
    if not path.is_file():
        raise FileNotFoundError(f"Template file not found: {path}")
    return path.read_bytes()
