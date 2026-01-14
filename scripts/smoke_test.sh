#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

TEMPLATE_HWPX=""
WITH_PDF="0"
KEEP_WORKDIR="0"

usage() {
  cat <<'USAGE' >&2
Usage: scripts/smoke_test.sh [--template-hwpx PATH] [--with-pdf]

Runs a basic environment + conversion smoke test:
- Checks: pandoc, python, (optional) Chrome/Chromium
- Validates python deps: lxml, Pillow
- Generates a tiny sample .docx via pandoc
- (Optional) DOCX -> HWPX -> DOCX roundtrip if --template-hwpx is provided
- (Optional) DOCX -> PDF via print/docx_to_review_pdf.sh if --with-pdf is set
- (Optional) Keep intermediate files with --keep
USAGE
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --template-hwpx)
      TEMPLATE_HWPX="${2:-}"
      shift 2
      ;;
    --with-pdf)
      WITH_PDF="1"
      shift 1
      ;;
    --keep)
      KEEP_WORKDIR="1"
      shift 1
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown arg: $1" >&2
      usage
      exit 2
      ;;
  esac
done

echo "[smoke] root: $ROOT" >&2

if ! command -v pandoc >/dev/null 2>&1; then
  echo "[smoke] ERROR: pandoc not found in PATH" >&2
  exit 1
fi

PY="$ROOT/.venv/bin/python"
if [[ ! -x "$PY" ]]; then
  # Convenience: when this toolkit is vendored inside another repo, allow using the parent venv.
  if [[ -x "$ROOT/../.venv/bin/python" ]]; then
    PY="$ROOT/../.venv/bin/python"
    echo "[smoke] WARN: using parent repo .venv: $PY" >&2
  elif command -v python3 >/dev/null 2>&1; then
    PY="python3"
    echo "[smoke] WARN: .venv not found; falling back to python3 (run 'make venv' for reproducibility)" >&2
  else
    echo "[smoke] ERROR: python3 not found and .venv missing" >&2
    exit 1
  fi
fi

"$PY" - <<'PY'
import sys
try:
    import lxml  # noqa: F401
    from PIL import Image  # noqa: F401
except Exception as e:
    print("[smoke] ERROR: missing python deps (lxml, Pillow). Run: make venv", file=sys.stderr)
    raise
print("[smoke] python deps ok", file=sys.stderr)
PY

CHROME_BIN=""
for candidate in google-chrome chromium chromium-browser; do
  if command -v "$candidate" >/dev/null 2>&1; then
    CHROME_BIN="$candidate"
    break
  fi
done
if [[ -n "$CHROME_BIN" ]]; then
  echo "[smoke] chrome: $CHROME_BIN" >&2
else
  echo "[smoke] chrome: (not found)" >&2
fi

WORKDIR="$ROOT/.tmp_smoke_test"
SUCCESS="0"
cleanup() {
  if [[ "$KEEP_WORKDIR" == "1" ]]; then
    echo "[smoke] kept workdir: $WORKDIR" >&2
    return
  fi
  if [[ "$SUCCESS" == "1" ]]; then
    rm -rf "$WORKDIR"
    return
  fi
  echo "[smoke] keeping workdir for debugging: $WORKDIR" >&2
}
trap cleanup EXIT

rm -rf "$WORKDIR"
mkdir -p "$WORKDIR"

cat >"$WORKDIR/sample.md" <<'MD'
# 제 1 장 테스트

## 제 1 절 개요

### 1. 소제목

본문 텍스트.

- 항목 1
- 항목 2

| 구분 | 값 |
| --- | --- |
| A | 1 |
| B | 2 |
MD

pandoc "$WORKDIR/sample.md" -o "$WORKDIR/sample.docx"
test -s "$WORKDIR/sample.docx"
echo "[smoke] generated: $WORKDIR/sample.docx" >&2

if [[ -n "$TEMPLATE_HWPX" ]]; then
  if [[ ! -f "$TEMPLATE_HWPX" ]]; then
    echo "[smoke] ERROR: template hwpx not found: $TEMPLATE_HWPX" >&2
    exit 1
  fi

  "$PY" "print/make_hwpx_template.py" "$TEMPLATE_HWPX" --output "$WORKDIR/template.hwpx"
  test -s "$WORKDIR/template.hwpx"

  "$PY" "print/convert_docx_to_hwpx.py" "$WORKDIR/sample.docx" \
    --output "$WORKDIR/out.hwpx" \
    --template-hwpx "$WORKDIR/template.hwpx" \
    --mode outline
  test -s "$WORKDIR/out.hwpx"

  "$PY" "print/convert_hwpx_to_docx.py" "$WORKDIR/out.hwpx" --output "$WORKDIR/roundtrip.docx"
  test -s "$WORKDIR/roundtrip.docx"

  echo "[smoke] roundtrip ok: $WORKDIR/out.hwpx -> $WORKDIR/roundtrip.docx" >&2
fi

if [[ "$WITH_PDF" == "1" ]]; then
  if [[ -z "$CHROME_BIN" ]]; then
    echo "[smoke] ERROR: --with-pdf requires Chrome/Chromium" >&2
    exit 1
  fi
  bash "print/docx_to_review_pdf.sh" "$WORKDIR/sample.docx" "$WORKDIR/out.pdf"
  test -s "$WORKDIR/out.pdf"
  echo "[smoke] pdf ok: $WORKDIR/out.pdf" >&2
fi

SUCCESS="1"
echo "[smoke] OK" >&2
