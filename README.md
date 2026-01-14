# KRISS DOCX/HWPX/PDF Print Toolkit

이 폴더는 `final_report_site/print/`에서 사용하던 **DOCX↔HWPX 변환**, **DOCX 머지**, **검토용 PDF 생성** 유틸을 별도 repo로 분리하기 위해 복사해둔 것입니다.

- 워크플로/사용법: `print/README.md`
- 스크립트 위치: `print/`

## Quickstart

```bash
make venv
make check
```

템플릿 기반 DOCX↔HWPX 라운드트립까지 확인하려면:

```bash
TEMPLATE_HWPX="template.hwpx" make check-template
```
