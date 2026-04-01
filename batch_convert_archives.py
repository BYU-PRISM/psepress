from __future__ import annotations

import argparse
import csv
import io
import secrets
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath

from latex_to_word import DocxTemplateConverter


APP_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = APP_DIR / "template.docx"


@dataclass
class ConversionResult:
    archive: str
    status: str
    selected_tex: str = ""
    output: str = ""
    error: str = ""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Batch-convert zipped LaTeX submissions into DOCX files using template.docx."
    )
    parser.add_argument("--input-dir", required=True, help="Folder containing submission zip archives.")
    parser.add_argument("--output-dir", required=True, help="Folder where DOCX files should be written.")
    parser.add_argument(
        "--template",
        default=str(DEFAULT_TEMPLATE),
        help="Path to the Word template DOCX. Defaults to template.docx in the app directory.",
    )
    parser.add_argument(
        "--pattern",
        default="*.zip",
        help="Archive filename pattern to process. Defaults to *.zip.",
    )
    parser.add_argument(
        "--report",
        default="",
        help="Optional CSV report path. Defaults to conversion-report.csv inside the output directory.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing DOCX outputs if they already exist.",
    )
    return parser.parse_args()


def visible_tex_members(archive_bytes: bytes) -> list[str]:
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


def default_tex_choice(options: list[str]) -> str:
    for preferred in ("main.tex", "template.tex"):
        for option in options:
            if PurePosixPath(option).name == preferred:
                return option
    return options[0]


def safe_extract_archive(archive_bytes: bytes, destination: Path) -> None:
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


def convert_archive(archive_path: Path, template_path: Path, output_path: Path, temp_root: Path) -> ConversionResult:
    archive_bytes = archive_path.read_bytes()
    try:
        tex_members = visible_tex_members(archive_bytes)
    except zipfile.BadZipFile as exc:
        return ConversionResult(archive=archive_path.name, status="failed", error=f"Invalid zip archive: {exc}")

    if not tex_members:
        return ConversionResult(archive=archive_path.name, status="failed", error="No .tex files found in archive.")

    selected_tex = default_tex_choice(tex_members)
    temp_root.mkdir(parents=True, exist_ok=True)
    work_dir = temp_root / f"{archive_path.stem}_{secrets.token_hex(4)}"
    try:
        extract_root = work_dir / "latex_project"
        extract_root.mkdir(parents=True, exist_ok=True)
        safe_extract_archive(archive_bytes, extract_root)

        tex_path = extract_root.joinpath(*PurePosixPath(selected_tex).parts)
        if not tex_path.is_file():
            return ConversionResult(
                archive=archive_path.name,
                status="failed",
                selected_tex=selected_tex,
                error=f"Selected TeX file was not found after extraction: {selected_tex}",
            )

        DocxTemplateConverter(
            template_path=template_path,
            tex_path=tex_path,
            output_path=output_path,
        ).convert()
    except Exception as exc:  # noqa: BLE001
        return ConversionResult(
            archive=archive_path.name,
            status="failed",
            selected_tex=selected_tex,
            output=str(output_path),
            error=str(exc),
        )
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    return ConversionResult(
        archive=archive_path.name,
        status="ok",
        selected_tex=selected_tex,
        output=str(output_path),
    )


def write_report(report_path: Path, results: list[ConversionResult]) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with report_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["archive", "status", "selected_tex", "output", "error"],
        )
        writer.writeheader()
        for result in results:
            writer.writerow(
                {
                    "archive": result.archive,
                    "status": result.status,
                    "selected_tex": result.selected_tex,
                    "output": result.output,
                    "error": result.error,
                }
            )


def main() -> None:
    args = parse_args()
    input_dir = Path(args.input_dir).resolve()
    output_dir = Path(args.output_dir).resolve()
    template_path = Path(args.template).resolve()
    report_path = Path(args.report).resolve() if args.report else output_dir / "conversion-report.csv"
    temp_root = output_dir / ".psepress_tmp"

    if not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")
    if not template_path.is_file():
        raise FileNotFoundError(f"Template DOCX not found: {template_path}")

    output_dir.mkdir(parents=True, exist_ok=True)
    archives = sorted(input_dir.glob(args.pattern))
    if not archives:
        print(f"No archives matched {args.pattern!r} in {input_dir}")
        write_report(report_path, [])
        print(f"Wrote empty report: {report_path}")
        return

    results: list[ConversionResult] = []
    print(f"Found {len(archives)} archive(s) in {input_dir}")
    for index, archive_path in enumerate(archives, start=1):
        output_path = output_dir / f"{archive_path.stem}.docx"
        print(f"[{index}/{len(archives)}] {archive_path.name}")
        if output_path.exists() and not args.overwrite:
            result = ConversionResult(
                archive=archive_path.name,
                status="skipped",
                output=str(output_path),
                error="Output already exists. Use --overwrite to replace it.",
            )
        else:
            result = convert_archive(archive_path, template_path, output_path, temp_root)
        results.append(result)
        if result.status == "ok":
            print(f"  ok -> {output_path.name} ({result.selected_tex})")
        elif result.status == "skipped":
            print(f"  skipped -> {output_path.name}")
        else:
            print(f"  failed -> {result.error}")

    write_report(report_path, results)
    if temp_root.exists() and not any(temp_root.iterdir()):
        temp_root.rmdir()
    ok_count = sum(result.status == "ok" for result in results)
    skipped_count = sum(result.status == "skipped" for result in results)
    failed_count = sum(result.status == "failed" for result in results)
    print(f"Completed. ok={ok_count} skipped={skipped_count} failed={failed_count}")
    print(f"Report: {report_path}")


if __name__ == "__main__":
    main()
