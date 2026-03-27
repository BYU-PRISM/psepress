from __future__ import annotations

import io
import tempfile
import zipfile
from pathlib import Path, PurePosixPath

import streamlit as st

from latex_to_word import DocxTemplateConverter


APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "template.docx"


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


def _convert_archive(archive_bytes: bytes, selected_tex: str) -> tuple[bytes, str]:
    with tempfile.TemporaryDirectory() as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        extract_root = temp_dir / "latex_project"
        extract_root.mkdir(parents=True, exist_ok=True)
        _safe_extract_archive(archive_bytes, extract_root)

        tex_path = extract_root.joinpath(*PurePosixPath(selected_tex).parts)
        if not tex_path.is_file():
            raise FileNotFoundError(f"Selected TeX file was not found after extraction: {selected_tex}")

        output_name = f"{tex_path.stem}-from-latex.docx"
        output_path = temp_dir / output_name

        DocxTemplateConverter(
            template_path=TEMPLATE_PATH,
            tex_path=tex_path,
            output_path=output_path,
        ).convert()

        return output_path.read_bytes(), output_name


st.set_page_config(
    page_title="LaTeX to Word Converter",
    page_icon="📄",
    layout="centered",
)

st.title("LaTeX to Word Converter")
st.write(
    "Upload a zip archive containing your LaTeX manuscript files, and this app will "
    "generate a Word document using the bundled `template.docx` and the existing template-aware converter."
)

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

    if st.button("Convert to Word", type="primary"):
        with st.spinner("Converting the archive to DOCX..."):
            try:
                output_bytes, output_name = _convert_archive(archive_bytes, selected_tex)
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
