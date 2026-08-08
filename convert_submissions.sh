#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TEMPLATE_DEFAULT="${SCRIPT_DIR}/template.docx"

if [[ $# -lt 2 ]]; then
  echo "Usage: $(basename "$0") INPUT_DIR OUTPUT_DIR [--conference KEY] [--conference-name NAME] [--conference-location TEXT] [--template PATH] [--pattern GLOB] [--report PATH] [--overwrite]" >&2
  exit 1
fi

INPUT_DIR="$1"
OUTPUT_DIR="$2"
shift 2

python3 "${SCRIPT_DIR}/batch_convert_archives.py" \
  --input-dir "${INPUT_DIR}" \
  --output-dir "${OUTPUT_DIR}" \
  --template "${TEMPLATE_DEFAULT}" \
  "$@"
