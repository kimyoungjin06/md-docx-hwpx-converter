# Print/Conversion Workflow (DOCX/HWPX/PDF)

> 실행 기준: 이 toolkit repo root에서 커맨드 실행.

## 단일 원본(편집) 원칙

- 최종 편집/머지의 단일 원본은 **DOCX**로 둔다: `print/kriss_master_merged.docx`
- 검토용 인쇄 PDF는 위 DOCX에서 뽑는다: `print/kriss_master_merged.pdf`

## 빠른 워크플로(추천)

1. (필요 시) HWPX → DOCX: `print/convert_hwpx_to_docx.py`
2. (선택) 데이터북 DOCX 생성(Quarto 프로젝트에서 생성)
3. 최종보고서 DOCX + 데이터북 DOCX 머지: `print/merge_docx_master.py`
4. 머지된 DOCX → 검토용 PDF: `print/docx_to_review_pdf.sh`
5. (옵션) DOCX → HWPX(한글 편집용): `print/convert_docx_to_hwpx.py`

## 의존성(필수)

- `pandoc` (HWPX/DOCX 변환 및 머지에 사용)
- `.venv` (python 패키지: `lxml`, `Pillow`)
- `google-chrome` 또는 `chromium` (DOCX → PDF 검토용 출력)
- `quarto` (데이터북 PDF/DOCX 빌드)

## 데이터북 PDF/DOCX 빌드(Quarto)

> 이 toolkit에는 Quarto 프로젝트가 포함되어 있지 않다. 데이터북을 Quarto로 빌드하는 경우, Quarto 프로젝트 디렉터리에서 실행한다.

- (예) Quarto 프로젝트에서 PDF 빌드: `quarto render --profile pdf --to pdf`
- (예) Quarto 프로젝트에서 DOCX 빌드: `quarto render --profile pdf --to docx`

## HWPX → DOCX 변환(로컬 작성 보고서)

`.hwpx` 보고서를 편집자 전달용 `.docx`로 “최소한의 문서 구조(제목/본문/표/그림)” 형태로 변환한다.

```bash
.venv/bin/python print/convert_hwpx_to_docx.py "input.hwpx" --output "output.docx"
```

## DOCX → HWPX 변환(한글 편집용, best-effort)

Hancom Office 없이도 열리는 형태를 목표로, **제목(개요/Outline) + 텍스트 중심**으로 DOCX를 HWPX로 변환한다.

- 지원(우선순위): `Heading 1~3` 기반 개요(Outline), 본문 텍스트, 리스트, 표(텍스트로 풀기), 그림(placeholder)
- 제한: 표/그림/수식의 “완전한 레이아웃” 보존은 Hancom Office/전용 변환기가 필요하다.

```bash
.venv/bin/python print/convert_docx_to_hwpx.py "input.docx" --output "output.hwpx" --template-hwpx "template.hwpx"
```

### HWPX 템플릿 만들기(권장)

`convert_docx_to_hwpx.py`는 스타일/설정이 들어있는 **template HWPX**가 필요하다.  
원본 HWPX를 그대로 템플릿으로 쓰면 `BinData/`가 포함되어 output이 수십 MB로 커질 수 있으므로, 아래 스크립트로 “가벼운 템플릿”을 만들어 쓰는 것을 권장한다.

```bash
.venv/bin/python print/make_hwpx_template.py "source.hwpx" --output "template.hwpx"
```

옵션:

- 개요(Outline) 우선(기본): `--mode outline`
- “제 장/제 절 …” 스타일 매핑: `--mode report` (기본으로 제목 선행 번호는 제거해서 중복을 피함)
  - 번호를 제거하지 않으려면: `--no-strip-heading-prefix`
- 중간 산출물 보관: `--keep-intermediate` (suffix: `.docx2hwpx`)

## 파일명 리네임 정책(예정)

- 한글 토큰을 영문 토큰으로 치환(예: `표준과학영역` → `stdscience`)
- 공백은 `_`로 치환
- 중복 파일명은 `_1`, `_2` 등 suffix로 처리
- 리네임 매핑은 `print/asset_rename_map.csv`에 기록

### Quarto 이미지(asset) 리네임 스크립트

LaTeX/PDF 빌드에서 경로/폰트 이슈를 줄이기 위해, `Final_Report_MD/report_assets/*.png` 파일명을 ASCII로 정리하고 문서 내 참조를 자동 갱신한다.

```bash
.venv/bin/python print/rename_report_images.py --dry-run --asset-dir "Final_Report_MD/report_assets"
```

- 실제 적용: `--dry-run` 제거
- 숫자 정렬(선택): `--pad-suffix-numbers 2` (예: `image_6.png` → `image_06.png`)

## 최종보고서(DOCX) + 데이터북(DOCX) 머지(편집용 단일 원본)

최종 편집/인쇄를 위해 단일 원본은 **DOCX**로 두되, 최종보고서의 “3장(분석 본문)” 구간은 최신성 있는 **데이터북 DOCX**로 덮어쓴다.

- 머지 스크립트: `final_report_site/print/merge_docx_master.py`
- 머지 스크립트: `print/merge_docx_master.py`
- 산출물(예시): `print/kriss_master_merged.docx`

```bash
.venv/bin/python print/merge_docx_master.py \\
  --master-docx "final_report.docx" \\
  --databook-docx "databook.docx" \\
  --output "print/kriss_master_merged.docx"
```

머지 시 적용되는 기본 처리:

- 3장 내부 3개 절을 데이터북 내용으로 덮어쓰기
- 캐러셀로 묶였던 반복 그림은 대표 1개만 본문에 남기고 `부록 > 추가 그림(캐러셀 해제본)`으로 이동
- 본문 URL은 `(데이터북 참조)`로 통일(최종 참고문헌 구간은 유지)
- 본문 그림 캡션은 `그림 3-1, 3-2, ...` 형태로 재번호(편집자가 InDesign에서 자유롭게 재조정 가능)

## DOCX → PDF(검토용)

머지된 DOCX를 빠르게 확인하기 위한 검토용 PDF를 생성한다.

```bash
bash print/docx_to_review_pdf.sh "print/kriss_master_merged.docx" "print/kriss_master_merged.pdf"
```

## 주의사항(요약)

- Quarto → PDF는 환경에 따라 LaTeX(TeX Live/TinyTeX) 설치가 필요할 수 있다.
- `docx_to_review_pdf.sh`는 MathJax CDN을 사용하므로 기본적으로 네트워크가 필요하다(오프라인이면 `--mathjax` 제거 후 재시도).
- DOCX ↔ HWPX 변환은 best-effort이며, 표/그림/수식은 “열리는 수준”을 목표로 한다(최종 제본용 레이아웃은 편집툴에서 조정 필요).
