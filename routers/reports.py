from fastapi import APIRouter, Request, Depends
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
import sqlite3
import json
import logging
import config
from database import get_db

router = APIRouter()
templates = Jinja2Templates(directory=config.TEMPLATE_DIR)
logger = logging.getLogger("app")

@router.get("/report/{specimen_id}", response_class=HTMLResponse)
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

@router.get("/api/search")
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

@router.get("/api/reports")
async def get_reports(conn: sqlite3.Connection = Depends(get_db)):
    cursor = conn.cursor()

    cursor.execute("SELECT specimen_id, created_at FROM reports ORDER BY created_at DESC")
    rows = cursor.fetchall()

    reports = [{"specimen_id": row["specimen_id"], "created_at": row["created_at"]} for row in rows]

    # JSON 파일 목록도 확인
    json_files = []
    if config.JSON_DIR.exists():
        json_files = [f.name for f in config.JSON_DIR.glob('*.json')]

    logger.info(f"\n=== 보고서 목록 호출 ===")
    logger.info(f"DB에 저장된 보고서 수: {len(reports)}")
    logger.info(f"JSON 파일 수: {len(json_files)}")
    logger.info(f"DB specimen_ids: {[r['specimen_id'] for r in reports[:10]]}")
    logger.info(f"JSON 파일명: {json_files[:10]}")
    logger.info(f"==================\n")

    return JSONResponse({"success": True, "reports": reports, "json_files": json_files})
