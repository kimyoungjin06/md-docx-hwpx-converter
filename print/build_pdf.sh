#!/usr/bin/env bash
set -euo pipefail

QUARTO_DIR="${1:-${QUARTO_DIR:-$PWD}}"
QUARTO_DIR="$(cd "$QUARTO_DIR" && pwd)"

if [[ ! -f "$QUARTO_DIR/_quarto.yml" && ! -f "$QUARTO_DIR/_quarto.yaml" ]]; then
  echo "ERROR: Quarto project not found (missing _quarto.yml) in: $QUARTO_DIR" >&2
  echo "Usage: $0 <quarto_project_dir>" >&2
  exit 2
fi

cd "$QUARTO_DIR"

PATH="$HOME/.local/bin:$PATH"

rm -rf _site_pdf
mkdir -p _site_pdf

quarto render --profile pdf --to pdf

# Quarto may emit PDFs to either _site_pdf/ or the project root depending on version/options.
if [[ -f "_site_pdf/kriss_databook_print.pdf.pdf" ]]; then
  mv "_site_pdf/kriss_databook_print.pdf.pdf" "_site_pdf/kriss_databook_print.pdf"
fi

if [[ -f "index.pdf" ]]; then
  for ext in aux log pdf tex toc out; do
    if [[ -f "index.${ext}" ]]; then
      mv "index.${ext}" "_site_pdf/index.${ext}"
    fi
  done
  if [[ -f "_site_pdf/index.pdf" ]]; then
    mv "_site_pdf/index.pdf" "_site_pdf/kriss_databook_print.pdf"
  fi
fi

echo "Done. PDF output: _site_pdf/kriss_databook_print.pdf"
