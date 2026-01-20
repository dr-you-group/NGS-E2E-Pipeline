import os
import sqlite3
import time
import traceback
import shutil
import config
import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, Form, File, UploadFile, Depends
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import json
from services.excel_parser import NGS_EXCEL2DB

try:
    from services.pptx_generator import NGS_PPT_Generator
except ImportError as e:
    logging.warning("경고: ppt_generator 모듈을 찾을 수 없습니다. PPT 다운로드 기능이 작동하지 않을 수 있습니다.")

config.setup_logging()
logger = logging.getLogger("app")

@asynccontextmanager
async def lifespan(app: FastAPI):
    # 디렉토리 확인 및 생성
    if not config.JSON_DIR.exists():
        config.JSON_DIR.mkdir(parents=True, exist_ok=True)

    if not config.TMP_DIR.exists():
        config.TMP_DIR.mkdir(parents=True, exist_ok=True)
        logger.info(f"임시 파일 디렉토리 생성: {config.TMP_DIR}")

    # DB 초기화
    init_db()

    yield

    # 앱 종료 시 실행할 로직이 있다면 여기에 작성
app = FastAPI(lifespan=lifespan)

# 현재 스크립트 파일의 경로
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app.mount("/static", StaticFiles(directory=config.STATIC_DIR), name="static")
templates = Jinja2Templates(directory=config.TEMPLATE_DIR)

def init_db():
    if not config.DB_PATH.exists():
        conn = sqlite3.connect(config.DB_PATH)
        cursor = conn.cursor()
        cursor.execute('''
                       CREATE TABLE IF NOT EXISTS reports
                       (
                           id
                           INTEGER
                           PRIMARY
                           KEY
                           AUTOINCREMENT,
                           specimen_id
                           TEXT
                           UNIQUE,
                           report_data
                           TEXT,
                           created_at
                           TIMESTAMP
                           DEFAULT
                           CURRENT_TIMESTAMP
                       )
                       ''')

        conn.commit()
        conn.close()

def get_db():
    # check_same_thread=False: 비동기/동기 혼용 시 스레드 에러 방지 옵션
    conn = sqlite3.connect(config.DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()

def safe_remove_file(file_path: str, max_retries: int = 5, delay: float = 0.5) -> bool:
    """
    파일을 안전하게 삭제하는 함수 (Windows 파일 잠금 문제 해결)
    
    Args:
        file_path: 삭제할 파일 경로
        max_retries: 최대 재시도 횟수
        delay: 재시도 간격 (초)
    
    Returns:
        bool: 삭제 성공 여부
    """
    if not os.path.exists(file_path):
        return True

    for attempt in range(max_retries):
        try:
            os.remove(file_path)
            logger.info(f"임시 파일 삭제 완료: {file_path}")
            return True
        except PermissionError as e:
            if attempt < max_retries - 1:
                logger.warning(f"파일 삭제 재시도 {attempt + 1}/{max_retries}: {file_path}")
                time.sleep(delay)
            else:
                logger.error(f"파일 삭제 실패 (최대 재시도 도달): {file_path} - {e}")
                return False
        except Exception as e:
            logger.error(f"파일 삭제 중 예상치 못한 오류: {file_path} - {e}")
            return False
    return False


def process_table_data_with_split_info(rows, max_rows_first_page=8):
    """
    rows 데이터를 headers와 data로 분리하고 분할 정보를 추가하는 함수
    첫 페이지에 표시할 최대 행 수를 고려하여 split_at 정보 제공
    """
    if not rows or len(rows) <= 1:
        return {"headers": [], "data": [], "split_at": None}

    headers = rows[0]
    data = rows[1:]

    # 데이터가 많을 경우 분할 위치 계산
    split_at = None
    if len(data) > max_rows_first_page:
        # 첫 페이지에 max_rows_first_page개, 나머지는 다음 페이지로
        split_at = max_rows_first_page

    return {"headers": headers, "data": data, "split_at": split_at}


# 테이블 데이터 처리 함수
def process_table_data(rows):
    """
    rows 데이터를 headers와 data로 분리하는 함수
    rows[0]은 headers로, rows[1:]는 data로 변환
    """
    if not rows or len(rows) <= 1:
        return {"headers": [], "data": []}

    headers = rows[0]
    data = rows[1:]

    return {"headers": headers, "data": data}


@app.get("/", response_class=HTMLResponse)
async def main_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/report/{specimen_id}", response_class=HTMLResponse)
async def show_report(request: Request, specimen_id: str, conn: sqlite3.Connection = Depends(get_db)):
    cursor = conn.cursor()
    cursor.execute("SELECT report_data FROM reports WHERE specimen_id = ?", (specimen_id,))
    result = cursor.fetchone()

    if result:
        report_data = json.loads(result["report_data"])
        return templates.TemplateResponse(
            "report.html",
            {
                "request": request,
                "specimen_id": specimen_id,
                "report_data": report_data,
                "debug": False
            }
        )
    else:
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "error": f"보고서를 찾을 수 없습니다: {specimen_id}"
            }
        )


@app.post("/generate-report", response_class=HTMLResponse)
async def generate_report(request: Request, specimen_id: str = Form(...), conn: sqlite3.Connection = Depends(get_db)):
    cursor = conn.cursor()
    cursor.execute("SELECT report_data FROM reports WHERE specimen_id = ?", (specimen_id,))
    result = cursor.fetchone()

    if result:
        report_data = json.loads(result["report_data"])
        return templates.TemplateResponse(
            "report.html",
            {
                "request": request,
                "specimen_id": specimen_id,
                "report_data": report_data,
                "debug": False
            }
        )
    else:
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "error": f"보고서를 찾을 수 없습니다: {specimen_id}"
            }
        )


@app.post("/api/download-pptx")
def download_pptx(specimen_id: str = Form(...), conn: sqlite3.Connection = Depends(get_db)):
    """
    특정 검체의 PPT 보고서를 생성하여 다운로드합니다.
    """
    cursor = conn.cursor()

    try:
        # 1. DB에서 리포트 데이터 조회
        cursor.execute("SELECT report_data FROM reports WHERE specimen_id = ?", (specimen_id,))
        result = cursor.fetchone()

        if not result:
            return JSONResponse({"success": False, "error": f"보고서를 찾을 수 없습니다: {specimen_id}"}, status_code=404)

        # 2. JSON 데이터 파싱
        report_data = json.loads(result["report_data"])

        # 3. PPT 생성 (메모리 상에서)
        generator = NGS_PPT_Generator()
        ppt_buffer = generator.generate(report_data)

        # 4. 파일 다운로드 응답 (StreamingResponse 사용)
        filename = f"{specimen_id}.pptx"

        return StreamingResponse(
            ppt_buffer,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f"attachment; filename={filename}",
                "Access-Control-Expose-Headers": "Content-Disposition"
            }
        )

    except Exception as e:
        logger.error(f"PPT 생성 중 오류 발생: {e}")
        return JSONResponse({"success": False, "error": str(e)}, status_code=500)


def save_json_file(specimen_id, report_data):
    """JSON 파일로 저장하는 함수"""
    try:
        json_file_path = config.JSON_DIR / f"{specimen_id}.json"
        with open(json_file_path, "w", encoding="utf-8") as json_file:
            json.dump(report_data, json_file, indent=4, ensure_ascii=False)
        logger.info(f"JSON 파일 저장 완료: {json_file_path}")
        return True
    except Exception as e:
        logger.error(f"JSON 파일 저장 실패: {str(e)}")
        return False


def extract_report_data(parser) -> dict:
    report_data = {}

    # 1. 기본 정보
    report_data["clinical_info"] = parser.get_Clinical_Info()
    report_data["biomarkers"] = parser.get_Biomarkers()
    report_data["failed_gene"] = parser.get_Failed_Gene()
    report_data["comments"] = parser.get_Comments()
    report_data["diagnostic_info"] = parser.get_Diagnostic_Info()
    report_data["filter_history"] = parser.get_Filter_History()
    report_data["drna_qubit"] = parser.get_DRNA_Qubit_Density()
    report_data["analysis_program"] = parser.get_Analysis_Program()
    report_data["diagnosis_user"] = parser.get_Diagnosis_User_Registration()
    report_data["panel_type"] = parser.panel

    # 2. QC 데이터
    report_data["qc"] = process_table_data(parser.get_QC())

    # 3. 변이 데이터 (Split Info 적용)
    # SNV
    h, rows = parser.get_SNV('VCS')
    report_data["snv_clinical"] = {"highlight": h, **process_table_data_with_split_info(rows, 8)}

    h, rows = parser.get_SNV('VUS')
    report_data["snv_unknown"] = {"highlight": h, **process_table_data_with_split_info(rows, 8)}

    # Fusion
    h, rows = parser.get_Fusion('VCS')
    report_data["fusion_clinical"] = {"highlight": h, **process_table_data(rows)}

    h, rows = parser.get_Fusion('VUS')
    report_data["fusion_unknown"] = {"highlight": h, **process_table_data(rows)}

    # CNV
    h, rows = parser.get_CNV('VCS')
    report_data["cnv_clinical"] = {"highlight": h, **process_table_data_with_split_info(rows, 10)}

    h, rows = parser.get_CNV('VUS')
    report_data["cnv_unknown"] = {"highlight": h, **process_table_data_with_split_info(rows, 10)}

    # LR BRCA
    h, rows = parser.get_LR_BRCA('VCS')
    report_data["lr_brca_clinical"] = {"highlight": h, **process_table_data(rows)}

    h, rows = parser.get_LR_BRCA('VUS')
    report_data["lr_brca_unknown"] = {"highlight": h, **process_table_data(rows)}

    # Splice
    h, rows = parser.get_Splice('VCS')
    report_data["splice_clinical"] = {"highlight": h, **process_table_data(rows)}

    h, rows = parser.get_Splice('VUS')
    report_data["splice_unknown"] = {"highlight": h, **process_table_data(rows)}

    return report_data


@app.post("/api/upload-excel")
def upload_excel(file: UploadFile = File(...), conn: sqlite3.Connection = Depends(get_db)):
    temp_file_path = config.TMP_DIR / f"temp_{file.filename}"
    parser = None

    try:
        # 동기 방식으로 임시 파일 저장 (shutil)
        with open(temp_file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        logger.info(f"엑셀 파일 임시 저장 완료: {temp_file_path}")

        # Excel 파일 파싱 (Path 객체를 문자열로 변환하여 전달)
        parser = NGS_EXCEL2DB(str(temp_file_path))

        # 데이터 추출 로직 분리 호출
        report_data = extract_report_data(parser)

        # 검체 정보 확인
        specimen_id = report_data["clinical_info"].get("검체 정보")
        if not specimen_id:
            raise ValueError("엑셀 파일에서 '검체 정보(Specimen ID)'를 찾을 수 없습니다.")

        # 로깅 (print -> logger)
        logger.info(f"\n=== 업로드 처리 시작: {file.filename} ===")
        logger.info(f"Target Specimen ID: {specimen_id}")

        pathology_num = parser.clinical_dict.get('병리번호', 'NOT FOUND')
        logger.debug(f"병리번호: {pathology_num}")  # debug 레벨 권장

        cursor = conn.cursor()

        # 중복 확인
        cursor.execute("SELECT specimen_id FROM reports WHERE specimen_id = ?", (specimen_id,))
        if cursor.fetchone():
            logger.warning(f"경고: {specimen_id} 보고서가 이미 존재합니다. 덮어쓰기(Replace)를 수행합니다.")

        # DB 저장 (Insert or Replace)
        cursor.execute(
            "INSERT OR REPLACE INTO reports (specimen_id, report_data) VALUES (?, ?)",
            (specimen_id, json.dumps(report_data))
        )
        conn.commit()

        logger.info(f"데이터베이스 저장 완료: {specimen_id}")
        logger.info(f"======================================\n")

        # JSON 파일 백업
        json_saved = save_json_file(specimen_id, report_data)

        # 리소스 정리
        parser.close()
        safe_remove_file(str(temp_file_path))

        return JSONResponse({
            "success": True,
            "specimen_id": specimen_id,
            "json_saved": json_saved
        })

    except Exception as e:
        # 예외 발생 시 정리
        if parser:
            try:
                parser.close()
            except:
                pass

        # os.path.exists -> Pathlib.exists()
        if temp_file_path.exists():
            safe_remove_file(str(temp_file_path))

        # print/traceback -> logger.error/exception
        logger.error(f"엑셀 업로드 중 치명적 오류: {e}")
        # logger.exception은 스택 트레이스를 자동으로 로그에 남깁니다.
        # logger.exception(e)

        return JSONResponse({"success": False, "error": str(e)})


@app.get("/api/search")
async def search_reports(q: str = "", conn: sqlite3.Connection = Depends(get_db)):
    if not q or len(q.strip()) < 1:
        return JSONResponse({"success": True, "results": []})

    search_term = q.strip().lower()
    cursor = conn.cursor()

    # specimen_id에 검색어가 포함된 보고서 찾기
    cursor.execute(
        "SELECT specimen_id, report_data FROM reports WHERE LOWER(specimen_id) LIKE ? ORDER BY specimen_id LIMIT 10",
        (f"%{search_term}%",)
    )
    rows = cursor.fetchall()

    results = []
    for row in rows:
        try:
            report_data = json.loads(row["report_data"])
            clinical_info = report_data.get("clinical_info", {})
            diagnosis_user = report_data.get("diagnosis_user", {})

            result_item = {
                "specimen_id": row["specimen_id"],
                "원발장기": clinical_info.get("원발 장기", "N/A"),
                "진단": clinical_info.get("진단", "N/A"),
                "signed1": diagnosis_user.get("Signed by", "N/A").split(", ")[1] if ", " in diagnosis_user.get(
                    "Signed by", "") else diagnosis_user.get("Signed by", "N/A")
            }
            results.append(result_item)
        except (json.JSONDecodeError, KeyError) as e:
            print(f"데이터 파싱 오류: {e}")
            continue

    return JSONResponse({"success": True, "results": results})


@app.get("/api/reports")
async def get_reports(conn: sqlite3.Connection = Depends(get_db)):
    cursor = conn.cursor()

    cursor.execute("SELECT specimen_id, created_at FROM reports ORDER BY created_at DESC")
    rows = cursor.fetchall()

    reports = [{"specimen_id": row["specimen_id"], "created_at": row["created_at"]} for row in rows]

    # JSON 파일 목록도 확인
    json_files = []
    if config.JSON_DIR.exists():
        json_files = [f for f in config.JSON_DIR.glob('*.json')]

    logger.info(f"\n=== 보고서 목록 호출 ===")
    logger.info(f"DB에 저장된 보고서 수: {len(reports)}")
    logger.info(f"JSON 파일 수: {len(json_files)}")
    logger.info(f"DB specimen_ids: {[r['specimen_id'] for r in reports[:10]]}")
    logger.info(f"JSON 파일명: {json_files[:10]}")
    logger.info(f"==================\n")

    return JSONResponse({"success": True, "reports": reports, "json_files": json_files})


@app.get("/api/specification/{panel_type}")
async def get_specification(panel_type: str):
    """
    패널 타입에 따른 Specification HTML 반환
    """
    if panel_type not in ['GE', 'SA']:
        return JSONResponse({"error": "Invalid panel type"}, status_code=400)

    spec_file = config.TEMPLATE_DIR / f"{panel_type}_Specification.html"

    if not spec_file.exists():
        return JSONResponse({"error": "Specification file not found"}, status_code=404)

    return FileResponse(spec_file, media_type="text/html")


@app.get("/api/gene-content/{content_type}")
async def get_gene_content(content_type: str):
    """
    Gene Content HTML 반환
    """
    valid_types = ['GE_Gene_Content_DRNA', 'SA_Gene_Content_DNA', 'SA_Gene_Content_RNA']

    if content_type not in valid_types:
        return JSONResponse({"error": "Invalid content type"}, status_code=400)

    gene_content_file = config.TEMPLATE_DIR / f"{content_type}.html"

    if not gene_content_file.exists():
        return JSONResponse({"error": "Gene content file not found"}, status_code=404)

    return FileResponse(gene_content_file, media_type="text/html")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:app", host="0.0.0.0", port=1234, reload=True, workers=1)
