from fastapi import APIRouter, Request, Form, File, UploadFile, Depends
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
import sqlite3
import json
import logging
import shutil
import config
from database import get_db
from services.excel_parser import NGS_EXCEL2DB
from services.report_service import extract_report_data
from services.file_service import save_json_file, safe_remove_file

router = APIRouter()
templates = Jinja2Templates(directory=config.TEMPLATE_DIR)
logger = logging.getLogger("app")

@router.post("/generate-report", response_class=HTMLResponse)
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

@router.post("/api/upload-excel")
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

        # 검체 정보 확인 (specimen_id는 병리번호만 사용)
        specimen_id = parser.clinical_dict.get("병리번호", "").strip()
        if not specimen_id:
            raise ValueError("엑셀 파일에서 '병리번호(Specimen ID)'를 찾을 수 없습니다.")

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
        
        return JSONResponse({"success": False, "error": str(e)})
