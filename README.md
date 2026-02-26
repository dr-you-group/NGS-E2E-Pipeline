# NGS E2E Pipeline

**NGS(차세대 염기서열 분석) 보고서 자동 생성 파이프라인**

Excel 기반의 NGS 분석 데이터를 업로드하면, 웹 기반 HTML 보고서 미리보기 및 PPTX 형식의 보고서를 자동으로 생성합니다.

---

## ✨ 주요 기능

| 기능 | 설명 |
|------|------|
| **Excel 파싱** | NGS 분석 결과 Excel 파일(.xlsx)을 자동으로 파싱하여 구조화된 데이터로 변환 |
| **HTML 보고서** | 브라우저에서 실시간으로 보고서를 미리보기 (동적 페이지네이션 지원) |
| **PPTX 보고서** | 템플릿 기반 PowerPoint 보고서 자동 생성 및 다운로드 |
| **검체 검색** | 병리번호 기반 실시간 검색 및 보고서 조회 |
| **드래그 앤 드롭 업로드** | 다수의 Excel 파일을 드래그 앤 드롭으로 일괄 업로드 |
| **패널 지원** | SA (Solid Assay) / GE (Gene Expression) 패널 및 V2 키트 버전 지원 |

### 지원하는 변이 유형

- **SNV** (Single Nucleotide Variant)
- **Fusion** (유전자 융합)
- **CNV** (Copy Number Variant)
- **LR BRCA** (Large Rearrangement BRCA)
- **Splice** (스플라이스 변이)
- **Biomarkers** — TMB (Tumor Mutational Burden), MSI (Microsatellite Instability)

---

## 기술 스택

- **Backend**: [FastAPI](https://fastapi.tiangolo.com/) + [Uvicorn](https://www.uvicorn.org/)
- **Database**: SQLite
- **Excel 파싱**: Pandas, openpyxl
- **PPTX 생성**: python-pptx
- **Frontend**: Vanilla HTML/CSS/JavaScript (Jinja2 템플릿)

---

## 프로젝트 구조

```
NGS-E2E-Pipeline/
├── app.py                  # FastAPI 애플리케이션 엔트리포인트
├── config.py               # 경로 및 로깅 설정
├── database.py             # SQLite DB 초기화 및 연결 관리
│
├── routers/                # API 라우터 (엔드포인트 정의)
│   ├── static.py           #   메인 페이지, Specification, Gene Content 조회
│   ├── reports.py          #   보고서 조회 및 검색 API
│   ├── upload.py           #   Excel 파일 업로드 및 DB 저장
│   └── downloads.py        #   PPTX 보고서 다운로드
│
├── services/               # 비즈니스 로직
│   ├── excel_parser.py     #   NGS Excel 파일 파싱 (NGS_EXCEL2DB 클래스)
│   ├── report_service.py   #   리포트 데이터 추출 및 가공
│   ├── pptx_generator.py   #   PPTX 보고서 생성 엔진 (NGS_PPT_Generator)
│   ├── make_pptx_result.py #   PPTX 결과 생성 유틸리티 (CLI 지원)
│   └── file_service.py     #   파일 저장/삭제 유틸리티
│
├── templates/              # Jinja2 HTML 템플릿
│   ├── index.html          #   메인 검색 & 업로드 페이지
│   ├── report.html         #   HTML 보고서 뷰어
│   ├── *_Specification*.html   #   패널별 검사 사양 (SA/GE, V1/V2)
│   └── *_Gene_Content_*.html   #   패널별 유전자 목록 (DNA/RNA)
│
├── static/                 # 정적 파일
│   ├── css/styles.css      #   스타일시트
│   └── js/script.js        #   프론트엔드 로직 (검색, 업로드, 페이지네이션)
│
├── resources/              # PPTX 보고서 템플릿
│   ├── NGS_GE_report_baseline.pptx
│   ├── NGS_GE_report_baseline_v2.pptx
│   ├── NGS_SA_report_baseline.pptx
│   └── NGS_SA_report_baseline_v2.pptx
│
├── json/                   # 파싱된 보고서 JSON 백업 (gitignored)
└── tmp/                    # 임시 파일 저장소 (gitignored)
```

---

## 시작하기

### 요구사항

- Python 3.13+

### 설치

```bash
# 저장소 클론
git clone https://github.com/dr-you-group/NGS-E2E-Pipeline.git
cd NGS-E2E-Pipeline

# 가상환경 생성 및 활성화
python -m venv venv
source venv/bin/activate   # macOS/Linux
# venv\Scripts\activate    # Windows

# 의존성 설치
pip install fastapi uvicorn pandas openpyxl python-pptx jinja2 python-multipart
```

### 실행

```bash
python app.py
```

서버가 `http://0.0.0.0:1234` 에서 시작됩니다.

---

## 사용 방법

### 1. Excel 업로드

메인 페이지(`http://localhost:1234`)에서 NGS 분석 결과 Excel 파일을 **드래그 앤 드롭**하거나 **클릭하여 선택**합니다.

### 2. HTML 보고서 확인

업로드 완료 후 **[결과 보기]** 버튼을 클릭하면 브라우저에서 보고서를 미리볼 수 있습니다.

### 3. PPTX 다운로드

HTML 보고서 화면에서 PPTX 다운로드 기능을 통해 PowerPoint 보고서를 자동 생성하고 다운로드합니다.

### 4. 보고서 검색

메인 페이지의 검색창에 **병리번호(Specimen ID)**를 입력하면 실시간으로 보고서를 검색할 수 있습니다.

---

## API 엔드포인트

| Method | Endpoint | 설명 |
|--------|----------|------|
| `GET` | `/` | 메인 페이지 |
| `POST` | `/api/upload-excel` | Excel 파일 업로드 |
| `GET` | `/api/search?q={query}` | 보고서 검색 |
| `GET` | `/api/reports` | 전체 보고서 목록 조회 |
| `GET` | `/report/{specimen_id}` | HTML 보고서 조회 |
| `POST` | `/generate-report` | 보고서 생성 (Form 제출) |
| `POST` | `/api/download-pptx` | PPTX 보고서 다운로드 |
| `GET` | `/api/specification/{panel_type}` | 검사 사양 HTML 조회 |
| `GET` | `/api/gene-content/{content_type}` | 유전자 목록 HTML 조회 |