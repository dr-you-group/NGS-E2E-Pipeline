from fastapi import APIRouter, Form, Depends
from fastapi.responses import JSONResponse, StreamingResponse
import sqlite3
import json
import logging
from database import get_db

logger = logging.getLogger("app")

try:
    from services.pptx_generator import NGS_PPT_Generator
except ImportError as e:
    logger.warning("경고: ppt_generator 모듈을 찾을 수 없습니다. PPT 다운로드 기능이 작동하지 않을 수 있습니다.")
    NGS_PPT_Generator = None

router = APIRouter()

@router.post("/api/download-pptx")
def download_pptx(specimen_id: str = Form(...), conn: sqlite3.Connection = Depends(get_db)):
    """
    특정 검체의 PPT 보고서를 생성하여 다운로드합니다.
    """
    if NGS_PPT_Generator is None:
        return JSONResponse({"success": False, "error": "PPT 생성 모듈이 로드되지 않았습니다."}, status_code=500)

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
        from datetime import datetime
        
        panel_type = report_data.get('panel_type', 'GE') # Default to GE if not present
        
        # Desired format: {specimen_id}_{Type}_report_{yymmdd}_auto.pptx
        # Extract sequence date from report_data and format it
        formatted_date = ""
        sequence_date_str = report_data.get('sequence_date', '').strip()
        
        if sequence_date_str:
            try:
                # Try parsing "2025-12-03" -> "251203"
                dt = datetime.strptime(sequence_date_str, "%Y-%m-%d")
                formatted_date = dt.strftime("%y%m%d")
            except Exception as e:
                logger.warning(f"Sequence Date 파싱 실패 ({sequence_date_str}): {e}")
                formatted_date = ""
                
        date_suffix = f"_{formatted_date}" if formatted_date else ""
        
        # Check if v2 report
        is_v2 = report_data.get('is_v2', False)
        report_str = "v2report" if is_v2 else "report"
        
        filename = f"{specimen_id}_{panel_type}_{report_str}{date_suffix}_auto.pptx"

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
