#!/usr/bin/env bash
set -euo pipefail

# Review-only DOCX → PDF renderer.
# - Path of least resistance for 빠른 검토용(편집/인쇄 최종본 용도 아님)
# - pandoc(DOCX→HTML) + Headless Chrome(HTML→PDF) 기반

if [[ $# -lt 2 ]]; then
  echo "Usage: $0 <input.docx> <output.pdf>" >&2
  exit 2
fi

INPUT_DOCX="$1"
OUTPUT_PDF="$2"

if ! command -v pandoc >/dev/null 2>&1; then
  echo "pandoc not found in PATH" >&2
  exit 1
fi

CHROME_BIN=""
for candidate in google-chrome chromium chromium-browser; do
  if command -v "$candidate" >/dev/null 2>&1; then
    CHROME_BIN="$candidate"
    break
  fi
done
if [[ -z "$CHROME_BIN" ]]; then
  echo "Chrome/Chromium not found (need google-chrome or chromium)" >&2
  exit 1
fi

TMP_DIR="$(mktemp -d -t kriss_docx_pdf_XXXXXX)"
cleanup() { rm -rf "$TMP_DIR"; }
trap cleanup EXIT

cat >"$TMP_DIR/header.html" <<'HTML'
<style>
@page {
  size: A4;
  margin: 20mm;
}
body {
  font-family: "Noto Sans CJK KR", "Noto Sans", sans-serif;
  font-size: 11pt;
  line-height: 1.45;
}
img {
  max-width: 100%;
  height: auto;
}
table {
  width: 100%;
  border-collapse: collapse;
}
th, td {
  border: 1px solid #999;
  padding: 4px 6px;
  vertical-align: top;
}
</style>
HTML

# NOTE: DOCX→PDF 직접 변환은 LaTeX 테이블 렌더 문제가 있어, HTML로 변환 후 Headless Chrome으로 PDF를 생성한다.
# 수식(LaTeX)은 MathJax로 렌더링해 PDF에 반영한다(네트워크 필요).
pandoc "$INPUT_DOCX" \
  -s --standalone \
  --extract-media="$TMP_DIR/media" \
  --mathjax="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js" \
  --metadata title="KRISS 최종보고서(검토용)" \
  --include-in-header="$TMP_DIR/header.html" \
  -o "$TMP_DIR/doc.html"

"$CHROME_BIN" \
  --headless \
  --disable-gpu \
  --no-sandbox \
  --allow-file-access-from-files \
  --virtual-time-budget=60000 \
  --print-to-pdf="$OUTPUT_PDF" \
  --print-to-pdf-no-header \
  "file://$TMP_DIR/doc.html"

if [[ ! -f "$OUTPUT_PDF" ]]; then
  echo "PDF generation failed: $OUTPUT_PDF" >&2
  exit 1
fi

echo "Wrote: $OUTPUT_PDF" >&2
