.PHONY: help venv check check-pdf check-template check-template-pdf clean

help:
	@echo "Targets:"
	@echo "  make venv                 Create .venv + install requirements"
	@echo "  make check                Run basic smoke test"
	@echo "  make check-pdf            Smoke test + DOCX->PDF (requires Chrome/Chromium)"
	@echo "  make check-template        Smoke test + DOCX<->HWPX (requires TEMPLATE_HWPX=... )"
	@echo "  make check-template-pdf    Template roundtrip + PDF (TEMPLATE_HWPX=... )"
	@echo "  make clean                Remove .venv + temp outputs"

venv:
	python3 -m venv .venv
	.venv/bin/python -m pip install -U pip
	.venv/bin/pip install -r requirements.txt

check:
	bash scripts/smoke_test.sh

check-pdf:
	bash scripts/smoke_test.sh --with-pdf

check-template:
	@test -n "$(TEMPLATE_HWPX)" || (echo "ERROR: set TEMPLATE_HWPX=/path/to/template.hwpx" >&2; exit 2)
	bash scripts/smoke_test.sh --template-hwpx "$(TEMPLATE_HWPX)"

check-template-pdf:
	@test -n "$(TEMPLATE_HWPX)" || (echo "ERROR: set TEMPLATE_HWPX=/path/to/template.hwpx" >&2; exit 2)
	bash scripts/smoke_test.sh --template-hwpx "$(TEMPLATE_HWPX)" --with-pdf

clean:
	rm -rf .venv
	rm -rf .tmp_smoke_test
	rm -f print/*.docx print/*.pdf print/*.hwpx
	rm -rf print/*.docx2hwpx print/*.hwpx2docx

