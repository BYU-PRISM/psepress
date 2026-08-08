from __future__ import annotations

import io
import re
import tempfile
import zipfile
from pathlib import Path, PurePosixPath

import streamlit as st

from conference_templates import (
    CONFERENCE_PRESETS,
    CUSTOM_CONFERENCE_KEY,
    DEFAULT_CONFERENCE_KEY,
    LATEX_CONFERENCE_KEY,
    ConferenceInfo,
    conference_from_form,
    generate_latex_template_archive,
    generate_template_filenames,
    generate_word_template,
    latex_escape_text,
    resolve_conference_selection,
)
from latex_to_word import DocxTemplateConverter


APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "template.docx"
LATEX_TEMPLATE_ZIP_PATH = APP_DIR / "latex_template" / "latex_template.zip"
KEYWORDS_PATH = APP_DIR / "PSEkeywords.txt"
SCT_LOGO_PATH = APP_DIR / "systems-control-transactions.png"
PSE_PRESS_LOGO_PATH = APP_DIR / "pse-press.png"


def _load_approved_keywords() -> list[str]:
    if not KEYWORDS_PATH.is_file():
        return []
    keywords: list[str] = []
    seen: set[str] = set()
    for raw_line in KEYWORDS_PATH.read_text(encoding="utf-8", errors="replace").splitlines():
        for part in [item.strip() for item in raw_line.split("\t")]:
            if not part:
                continue
            if part in seen:
                continue
            seen.add(part)
            keywords.append(part)
    return keywords


def _parse_keyword_text(value: str) -> list[str]:
    keywords: list[str] = []
    seen: set[str] = set()
    for raw_part in re.split(r"[,;\r\n]+", value):
        keyword = raw_part.strip()
        if not keyword or keyword in seen:
            continue
        seen.add(keyword)
        keywords.append(keyword)
    return keywords


def _merge_keywords(*groups: list[str]) -> list[str]:
    merged: list[str] = []
    seen: set[str] = set()
    for group in groups:
        for keyword in group:
            if keyword in seen:
                continue
            seen.add(keyword)
            merged.append(keyword)
    return merged


def _override_keywords_in_tex(tex_path: Path, selected_keywords: list[str]) -> None:
    keyword_text = ", ".join(latex_escape_text(keyword) for keyword in selected_keywords)
    override = f"\\renewcommand{{\\PaperKeywords}}{{%\n{keyword_text}}}"
    tex_source = tex_path.read_text(encoding="utf-8")
    pattern = re.compile(r"\\(?:re)?newcommand\{\\PaperKeywords\}\{.*?\}", re.DOTALL)
    if pattern.search(tex_source):
        tex_source = pattern.sub(lambda _: override, tex_source, count=1)
    elif "\\begin{document}" in tex_source:
        tex_source = tex_source.replace("\\begin{document}", override + "\n\n\\begin{document}", 1)
    else:
        tex_source = tex_source.rstrip() + "\n\n" + override + "\n"
    tex_path.write_text(tex_source, encoding="utf-8")


def _visible_tex_members(archive_bytes: bytes) -> list[str]:
    with zipfile.ZipFile(io.BytesIO(archive_bytes)) as archive:
        members: list[str] = []
        for info in archive.infolist():
            if info.is_dir():
                continue
            rel_path = PurePosixPath(info.filename)
            if "__MACOSX" in rel_path.parts:
                continue
            if rel_path.name.startswith("."):
                continue
            if rel_path.suffix.lower() == ".tex":
                members.append(str(rel_path))
    return sorted(set(members))


def _safe_extract_archive(archive_bytes: bytes, destination: Path) -> None:
    destination = destination.resolve()
    with zipfile.ZipFile(io.BytesIO(archive_bytes)) as archive:
        for info in archive.infolist():
            rel_path = PurePosixPath(info.filename)
            if not rel_path.parts:
                continue
            if "__MACOSX" in rel_path.parts:
                continue

            target_path = destination.joinpath(*rel_path.parts).resolve()
            if destination not in (target_path, *target_path.parents):
                raise ValueError(f"Unsafe archive path: {info.filename}")

            if info.is_dir():
                target_path.mkdir(parents=True, exist_ok=True)
                continue

            target_path.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(info) as source, target_path.open("wb") as sink:
                sink.write(source.read())


def _default_tex_choice(options: list[str]) -> str:
    for preferred in ("main.tex", "template.tex"):
        for option in options:
            if PurePosixPath(option).name == preferred:
                return option
    return options[0]


def _conference_option_label(selection: str) -> str:
    if selection == CUSTOM_CONFERENCE_KEY:
        return "Custom conference"
    if selection == LATEX_CONFERENCE_KEY:
        return "Use conference from LaTeX"
    return CONFERENCE_PRESETS[selection].name


def _convert_archive(
    archive_bytes: bytes,
    selected_tex: str,
    selected_keywords: list[str] | None = None,
    conference: ConferenceInfo | None = CONFERENCE_PRESETS[DEFAULT_CONFERENCE_KEY],
) -> tuple[bytes, str]:
    with tempfile.TemporaryDirectory() as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        extract_root = temp_dir / "latex_project"
        extract_root.mkdir(parents=True, exist_ok=True)
        _safe_extract_archive(archive_bytes, extract_root)

        tex_path = extract_root.joinpath(*PurePosixPath(selected_tex).parts)
        if not tex_path.is_file():
            raise FileNotFoundError(f"Selected TeX file was not found after extraction: {selected_tex}")
        if selected_keywords:
            _override_keywords_in_tex(tex_path, selected_keywords)

        output_name = f"{tex_path.stem}-from-latex.docx"
        output_path = temp_dir / output_name

        DocxTemplateConverter(
            template_path=TEMPLATE_PATH,
            tex_path=tex_path,
            output_path=output_path,
            conference=conference,
        ).convert()

        return output_path.read_bytes(), output_name


def _render_hero() -> None:
    st.markdown(
        """
        <style>
        :root {
            --pse-hero-border: rgba(10, 85, 168, 0.18);
            --pse-hero-background:
                linear-gradient(135deg, rgba(8, 88, 170, 0.10), rgba(255, 255, 255, 0.96) 42%),
                linear-gradient(180deg, rgba(6, 170, 178, 0.07), rgba(255, 255, 255, 0.99));
            --pse-hero-shadow: 0 14px 34px rgba(10, 34, 66, 0.08);
            --pse-hero-kicker-bg: rgba(6, 170, 178, 0.12);
            --pse-hero-kicker-text: #0a5aa8;
            --pse-hero-title: #0b3a75;
            --pse-hero-copy: rgba(20, 34, 51, 0.90);
        }
        @media (prefers-color-scheme: dark) {
            :root {
                --pse-hero-border: rgba(113, 160, 224, 0.28);
                --pse-hero-background:
                    linear-gradient(135deg, rgba(18, 72, 132, 0.52), rgba(17, 24, 39, 0.96) 45%),
                    linear-gradient(180deg, rgba(6, 170, 178, 0.18), rgba(17, 24, 39, 0.98));
                --pse-hero-shadow: 0 18px 40px rgba(0, 0, 0, 0.30);
                --pse-hero-kicker-bg: rgba(159, 216, 255, 0.12);
                --pse-hero-kicker-text: #9fd8ff;
                --pse-hero-title: #f2f6fb;
                --pse-hero-copy: rgba(242, 246, 251, 0.92);
            }
        }
        .pse-hero {
            padding: 1.2rem 1.25rem 1.1rem 1.25rem;
            border: 1px solid var(--pse-hero-border);
            border-radius: 20px;
            background: var(--pse-hero-background);
            box-shadow: var(--pse-hero-shadow);
            margin-bottom: 1.2rem;
            color: var(--text-color, inherit);
        }
        .pse-hero-kicker {
            display: inline-block;
            padding: 0.28rem 0.62rem;
            border-radius: 999px;
            background: var(--pse-hero-kicker-bg);
            color: var(--pse-hero-kicker-text);
            font-size: 0.82rem;
            font-weight: 700;
            letter-spacing: 0.02em;
            text-transform: uppercase;
            margin-bottom: 0.7rem;
        }
        .pse-hero-title {
            font-size: 2rem;
            font-weight: 700;
            line-height: 1.08;
            color: var(--pse-hero-title);
            margin: 0 0 0.45rem 0;
        }
        .pse-hero-copy {
            font-size: 1rem;
            line-height: 1.55;
            color: var(--pse-hero-copy);
            margin: 0;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown('<div class="pse-hero">', unsafe_allow_html=True)
    col_left, col_body, col_right = st.columns([1.9, 3.3, 1.0], vertical_alignment="center")
    with col_left:
        if SCT_LOGO_PATH.is_file():
            st.image(str(SCT_LOGO_PATH), width="stretch")
    with col_body:
        st.markdown('<div class="pse-hero-kicker">PSE Press Workflow</div>', unsafe_allow_html=True)
        st.markdown('<div class="pse-hero-title">Conference Templates &amp; LaTeX to Word</div>', unsafe_allow_html=True)
        st.markdown(
            '<p class="pse-hero-copy">Generate conference-specific Word and LaTeX templates, or convert a LaTeX project archive into a Word manuscript while preserving the PSE Press structure.</p>',
            unsafe_allow_html=True,
        )
    with col_right:
        if PSE_PRESS_LOGO_PATH.is_file():
            st.image(str(PSE_PRESS_LOGO_PATH), width="stretch")
    st.markdown("</div>", unsafe_allow_html=True)


st.set_page_config(
    page_title="PSE Press Templates and Converter",
    page_icon="📄",
    layout="centered",
)

_render_hero()

if not TEMPLATE_PATH.is_file():
    st.error(f"Missing bundled Word template: {TEMPLATE_PATH}")
    st.stop()
if not LATEX_TEMPLATE_ZIP_PATH.is_file():
    st.error(f"Missing bundled LaTeX template: {LATEX_TEMPLATE_ZIP_PATH}")
    st.stop()

st.subheader("Generate Conference Templates")
st.caption(
    "Enter the conference header information once, then download both a Word template and a complete LaTeX project."
)

with st.form("conference_template_generator"):
    generator_name = st.text_input(
        "Conference name",
        value="",
        placeholder="Example: ESCAPE 37 - European Symposium on Computer Aided Process Engineering",
    )
    location_left, location_middle, location_right = st.columns([1.15, 1.0, 1.0])
    with location_left:
        generator_city = st.text_input("City", value="", placeholder="Trondheim")
    with location_middle:
        generator_region = st.text_input("State or region (optional)", value="", placeholder="Trøndelag")
    with location_right:
        generator_country = st.text_input("Country", value="", placeholder="Norway")
    date_left, date_right = st.columns(2)
    with date_left:
        generator_start_date = st.date_input("Start date", value=None)
    with date_right:
        generator_end_date = st.date_input("End date", value=None)
    generate_templates = st.form_submit_button("Generate templates", type="primary")

if generate_templates:
    try:
        if generator_start_date is None or generator_end_date is None:
            raise ValueError("Conference start date and end date are required.")
        generated_conference = conference_from_form(
            name=generator_name,
            city=generator_city,
            region=generator_region,
            country=generator_country,
            start_date=generator_start_date,
            end_date=generator_end_date,
        )
        word_name, latex_name = generate_template_filenames(generated_conference)
        generated_word = generate_word_template(TEMPLATE_PATH.read_bytes(), generated_conference)
        generated_latex = generate_latex_template_archive(
            LATEX_TEMPLATE_ZIP_PATH.read_bytes(),
            generated_conference,
        )
    except Exception as exc:
        st.error(str(exc))
    else:
        st.session_state["generated_conference_templates"] = {
            "conference": generated_conference,
            "word_name": word_name,
            "word_bytes": generated_word,
            "latex_name": latex_name,
            "latex_bytes": generated_latex,
        }
        st.success("Conference templates are ready to download.")

generated_templates = st.session_state.get("generated_conference_templates")
if generated_templates:
    st.caption(
        f"Header: {generated_templates['conference'].name} | "
        f"{generated_templates['conference'].location}"
    )
    word_download, latex_download = st.columns(2)
    with word_download:
        st.download_button(
            "Download Word template",
            data=generated_templates["word_bytes"],
            file_name=generated_templates["word_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with latex_download:
        st.download_button(
            "Download LaTeX template",
            data=generated_templates["latex_bytes"],
            file_name=generated_templates["latex_name"],
            mime="application/zip",
            use_container_width=True,
        )

st.divider()
st.subheader("Convert LaTeX to Word")

with st.expander("What to include in the zip", expanded=True):
    st.markdown(
        "- Your main `.tex` file, usually `main.tex`\n"
        "- `refs.bib` if your manuscript uses it\n"
        "- Any figures or other files referenced by the LaTeX source\n"
        "- Any additional `.tex` files pulled in with `\\input{...}`"
    )

approved_keywords = _load_approved_keywords()

uploaded_file = st.file_uploader(
    "Upload a zip archive",
    type=["zip"],
    help="The archive should preserve the same relative file structure your manuscript uses locally.",
)

if uploaded_file is not None:
    archive_bytes = uploaded_file.getvalue()

    try:
        tex_members = _visible_tex_members(archive_bytes)
    except zipfile.BadZipFile:
        st.error("That file is not a valid zip archive.")
        st.stop()

    if not tex_members:
        st.error("No `.tex` files were found in the uploaded archive.")
        st.stop()

    selected_tex = st.selectbox(
        "TeX file to convert",
        options=tex_members,
        index=tex_members.index(_default_tex_choice(tex_members)),
        help="If your archive contains multiple `.tex` files, choose the manuscript entry point.",
    )

    conference_selection = st.selectbox(
        "Conference",
        options=[*CONFERENCE_PRESETS, CUSTOM_CONFERENCE_KEY, LATEX_CONFERENCE_KEY],
        index=0,
        format_func=_conference_option_label,
        help=(
            "Choose a built-in conference, enter custom header text, or use the conference command "
            "already present in the selected LaTeX manuscript."
        ),
    )
    custom_conference_name = ""
    custom_conference_location = ""
    if conference_selection == CUSTOM_CONFERENCE_KEY:
        custom_conference_name = st.text_input(
            "Custom conference name",
            value="",
            placeholder="Conference name as it should appear in the Word header",
        )
        custom_conference_location = st.text_input(
            "Custom conference location and dates",
            value="",
            placeholder="City, Country, 6-9 June 2027",
        )

    keyword_mode = st.radio(
        "Keywords",
        options=("Use keywords already in manuscript", "Set keywords for this conversion"),
        index=0,
        help="Leave the manuscript keywords unchanged, or override them for this conversion with recommended and/or custom keywords.",
    )

    selected_keywords: list[str] = []
    additional_keywords_text = ""
    if keyword_mode == "Set keywords for this conversion":
        if approved_keywords:
            selected_keywords = st.multiselect(
                "Recommended keywords",
                options=approved_keywords,
                default=[],
                help="Optional suggestions from `PSEkeywords.txt`. You can choose any number.",
            )
        additional_keywords_text = st.text_area(
            "Other keywords",
            value="",
            height=110,
            help="Add any keywords you want. Separate them with commas or new lines.",
            placeholder="Example: Process Safety, Digital Twins, New custom keyword",
        )
        final_keywords = _merge_keywords(selected_keywords, _parse_keyword_text(additional_keywords_text))
        if final_keywords:
            st.caption("Keywords to use for this conversion: " + ", ".join(final_keywords))
        else:
            st.caption("Choose any recommended keywords and/or type your own custom keywords.")

    if st.button("Convert to Word", type="primary"):
        final_keywords = _merge_keywords(selected_keywords, _parse_keyword_text(additional_keywords_text))
        if keyword_mode == "Set keywords for this conversion" and not final_keywords:
            st.error("Add at least one keyword, or use the manuscript keywords option.")
            st.stop()
        try:
            selected_conference = resolve_conference_selection(
                conference_selection,
                custom_name=custom_conference_name,
                custom_location=custom_conference_location,
            )
        except ValueError as exc:
            st.error(str(exc))
            st.stop()
        with st.spinner("Converting the archive to DOCX..."):
            try:
                output_bytes, output_name = _convert_archive(
                    archive_bytes,
                    selected_tex,
                    final_keywords if keyword_mode == "Set keywords for this conversion" else None,
                    selected_conference,
                )
            except Exception as exc:
                st.error("Conversion failed.")
                st.exception(exc)
            else:
                st.session_state["output_bytes"] = output_bytes
                st.session_state["output_name"] = output_name
                st.success("Conversion finished.")

if "output_bytes" in st.session_state and "output_name" in st.session_state:
    st.download_button(
        "Download Word document",
        data=st.session_state["output_bytes"],
        file_name=st.session_state["output_name"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
