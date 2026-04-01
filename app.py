from __future__ import annotations

import io
import re
import tempfile
import zipfile
from pathlib import Path, PurePosixPath

import streamlit as st

from latex_to_word import DocxTemplateConverter


APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "template.docx"
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


def _latex_escape_text(value: str) -> str:
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
    }
    return "".join(replacements.get(ch, ch) for ch in value)


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
    keyword_text = ", ".join(_latex_escape_text(keyword) for keyword in selected_keywords)
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


def _convert_archive(archive_bytes: bytes, selected_tex: str, selected_keywords: list[str] | None = None) -> tuple[bytes, str]:
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
        ).convert()

        return output_path.read_bytes(), output_name


def _render_hero() -> None:
    st.markdown(
        """
        <style>
        .pse-hero {
            padding: 1.2rem 1.25rem 1.1rem 1.25rem;
            border: 1px solid rgba(10, 85, 168, 0.18);
            border-radius: 20px;
            background:
                linear-gradient(135deg, rgba(8, 88, 170, 0.08), rgba(255, 255, 255, 0.95) 42%),
                linear-gradient(180deg, rgba(6, 170, 178, 0.06), rgba(255, 255, 255, 0.99));
            margin-bottom: 1.2rem;
        }
        .pse-hero-kicker {
            display: inline-block;
            padding: 0.28rem 0.62rem;
            border-radius: 999px;
            background: rgba(6, 170, 178, 0.12);
            color: #0a5aa8;
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
            color: #0b3a75;
            margin: 0 0 0.45rem 0;
        }
        .pse-hero-copy {
            font-size: 1rem;
            line-height: 1.55;
            color: #213547;
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
            st.image(str(SCT_LOGO_PATH), use_container_width=True)
    with col_body:
        st.markdown('<div class="pse-hero-kicker">PSE Press Workflow</div>', unsafe_allow_html=True)
        st.markdown('<div class="pse-hero-title">LaTeX to Word Converter</div>', unsafe_allow_html=True)
        st.markdown(
            '<p class="pse-hero-copy">Upload a LaTeX project archive and generate a Word manuscript that preserves the PSE Press template structure while rebuilding the editable content sections.</p>',
            unsafe_allow_html=True,
        )
    with col_right:
        if PSE_PRESS_LOGO_PATH.is_file():
            st.image(str(PSE_PRESS_LOGO_PATH), use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)


st.set_page_config(
    page_title="LaTeX to Word Converter",
    page_icon="📄",
    layout="centered",
)

_render_hero()

with st.expander("What to include in the zip", expanded=True):
    st.markdown(
        "- Your main `.tex` file, usually `main.tex`\n"
        "- `refs.bib` if your manuscript uses it\n"
        "- Any figures or other files referenced by the LaTeX source\n"
        "- Any additional `.tex` files pulled in with `\\input{...}`"
    )

if not TEMPLATE_PATH.is_file():
    st.error(f"Missing bundled template: {TEMPLATE_PATH}")
    st.stop()

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
        with st.spinner("Converting the archive to DOCX..."):
            try:
                output_bytes, output_name = _convert_archive(
                    archive_bytes,
                    selected_tex,
                    final_keywords if keyword_mode == "Set keywords for this conversion" else None,
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
