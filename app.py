from fastapi import FastAPI, Request, Form, File, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import os
import sqlite3
import json
import traceback
from utils import NGS_EXCEL2DB

app = FastAPI()

# 현재 스크립트 파일의 경로
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))

# SQLite 데이터베이스 설정
DB_PATH = os.path.join(BASE_DIR, "ngs_reports.db")
JSON_DIR = os.path.join(BASE_DIR, "json")

# JSON 디렉토리 확인 및 생성
if not os.path.exists(JSON_DIR):
    os.makedirs(JSON_DIR)

def init_db():
    if not os.path.exists(DB_PATH):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            specimen_id TEXT UNIQUE,
            report_data TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        ''')
        
        conn.commit()
        conn.close()

init_db()

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
async def read_item(request: Request):
    return templates.TemplateResponse("report.html", {"request": request})

@app.post("/generate-report", response_class=HTMLResponse)
async def generate_report(request: Request, specimen_id: str = Form(...)):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
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
                "show_report": True,
                "debug": False  # 디버그 모드 설정
            }
        )
    else:
        return templates.TemplateResponse(
            "report.html", 
            {
                "request": request,
                "error": f"보고서를 찾을 수 없습니다: {specimen_id}"
            }
        )

@app.get("/generate-report", response_class=HTMLResponse)
async def get_generate_report(request: Request, specimen_id: str = None):
    if not specimen_id:
        return templates.TemplateResponse("report.html", {"request": request})
    
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
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
                "show_report": True,
                "debug": False  # 디버그 모드 설정
            }
        )
    else:
        return templates.TemplateResponse(
            "report.html", 
            {
                "request": request,
                "error": f"보고서를 찾을 수 없습니다: {specimen_id}"
            }
        )

def save_json_file(specimen_id, report_data):
    """JSON 파일로 저장하는 함수"""
    try:
        json_file_path = os.path.join(JSON_DIR, f"{specimen_id}.json")
        with open(json_file_path, "w", encoding="utf-8") as json_file:
            json.dump(report_data, json_file, indent=4, ensure_ascii=False)
        print(f"JSON 파일 저장 완료: {json_file_path}")
        return True
    except Exception as e:
        print(f"JSON 파일 저장 실패: {str(e)}")
        traceback.print_exc()
        return False

@app.post("/api/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    temp_file_path = os.path.join(BASE_DIR, f"temp_{file.filename}")
    
    try:
        # 임시 파일에 저장
        with open(temp_file_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)
        
        print(f"엑셀 파일 임시 저장 완료: {temp_file_path}")
        
        # Excel 파일 파싱
        parser = NGS_EXCEL2DB(temp_file_path)
        
        # 보고서 데이터 생성
        report_data = {}
        
        # 검체 정보
        report_data["clinical_info"] = parser.get_Clinical_Info()
        
        # SNV 데이터 - 수정된 값 사용
        snv_clinical_highlight, snv_clinical_rows = parser.get_SNV('VCS')
        report_data["snv_clinical"] = {
            "highlight": snv_clinical_highlight,
            **process_table_data(snv_clinical_rows)
        }
        
        snv_unknown_highlight, snv_unknown_rows = parser.get_SNV('VUS')
        report_data["snv_unknown"] = {
            "highlight": snv_unknown_highlight,
            **process_table_data(snv_unknown_rows)
        }
        
        # Fusion 데이터 - 수정된 값 사용
        fusion_clinical_highlight, fusion_clinical_rows = parser.get_Fusion('VCS')
        report_data["fusion_clinical"] = {
            "highlight": fusion_clinical_highlight,
            **process_table_data(fusion_clinical_rows)
        }
        
        fusion_unknown_highlight, fusion_unknown_rows = parser.get_Fusion('VUS')
        report_data["fusion_unknown"] = {
            "highlight": fusion_unknown_highlight,
            **process_table_data(fusion_unknown_rows)
        }
        
        # CNV 데이터 - 수정된 값 사용
        cnv_clinical_highlight, cnv_clinical_rows = parser.get_CNV('VCS')
        report_data["cnv_clinical"] = {
            "highlight": cnv_clinical_highlight,
            **process_table_data(cnv_clinical_rows)
        }
        
        cnv_unknown_highlight, cnv_unknown_rows = parser.get_CNV('VUS')
        report_data["cnv_unknown"] = {
            "highlight": cnv_unknown_highlight,
            **process_table_data(cnv_unknown_rows)
        }
        
        # LR BRCA 데이터 - 수정된 값 사용
        lr_brca_clinical_highlight, lr_brca_clinical_rows = parser.get_LR_BRCA('VCS')
        report_data["lr_brca_clinical"] = {
            "highlight": lr_brca_clinical_highlight,
            **process_table_data(lr_brca_clinical_rows)
        }
        
        lr_brca_unknown_highlight, lr_brca_unknown_rows = parser.get_LR_BRCA('VUS')
        report_data["lr_brca_unknown"] = {
            "highlight": lr_brca_unknown_highlight,
            **process_table_data(lr_brca_unknown_rows)
        }
        
        # Splice 데이터 - 수정된 값 사용
        splice_clinical_highlight, splice_clinical_rows = parser.get_Splice('VCS')
        report_data["splice_clinical"] = {
            "highlight": splice_clinical_highlight,
            **process_table_data(splice_clinical_rows)
        }
        
        splice_unknown_highlight, splice_unknown_rows = parser.get_Splice('VUS')
        report_data["splice_unknown"] = {
            "highlight": splice_unknown_highlight,
            **process_table_data(splice_unknown_rows)
        }
        
        # Biomarkers
        report_data["biomarkers"] = parser.get_Biomarkers()
        
        # Failed Gene
        report_data["failed_gene"] = parser.get_Failed_Gene()
        
        # Comments
        report_data["comments"] = parser.get_Comments()
        
        # 검사정보
        report_data["diagnostic_info"] = parser.get_Diagnostic_Info()
        
        # Filter History
        report_data["filter_history"] = parser.get_Filter_History()
        
        # DNA, RNA
        report_data["drna_qubit"] = parser.get_DRNA_Qubit_Density()
        
        # QC
        qc_rows = parser.get_QC()
        report_data["qc"] = process_table_data(qc_rows)
        
        # Analysis Program
        report_data["analysis_program"] = parser.get_Analysis_Program()
        
        # 사용자 정보
        report_data["diagnosis_user"] = parser.get_Diagnosis_User_Registration()
        
        # Panel Type
        report_data["panel_type"] = parser.panel
        
        # 데이터베이스에 저장
        specimen_id = report_data["clinical_info"]["검체 정보"]
        
        print(f"\n=== 업로드 디버깅 ===")
        print(f"업로드 파일명: {file.filename}")
        print(f"추출된 specimen_id: {specimen_id}")
        print(f"병리번호 (clinical_dict): {parser.clinical_dict.get('병리번호', 'NOT FOUND')}")
        
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # 기존 데이터 확인
        cursor.execute("SELECT specimen_id FROM reports WHERE specimen_id = ?", (specimen_id,))
        existing = cursor.fetchone()
        if existing:
            print(f"경고: {specimen_id}는 이미 존재합니다. 덮어쓰기됩니다.")
        
        # 저장 전 전체 데이터 확인
        cursor.execute("SELECT COUNT(*) as count FROM reports")
        before_count = cursor.fetchone()[0]
        print(f"저장 전 전체 보고서 수: {before_count}")
        
        cursor.execute(
            "INSERT OR REPLACE INTO reports (specimen_id, report_data) VALUES (?, ?)",
            (specimen_id, json.dumps(report_data))
        )
        
        conn.commit()
        
        # 저장 후 전체 데이터 확인
        cursor.execute("SELECT COUNT(*) as count FROM reports")
        after_count = cursor.fetchone()[0]
        print(f"저장 후 전체 보고서 수: {after_count}")
        
        # 현재 저장된 모든 specimen_id 출력
        cursor.execute("SELECT specimen_id FROM reports ORDER BY created_at DESC LIMIT 10")
        all_ids = cursor.fetchall()
        print(f"최근 10개 specimen_id: {[row[0] for row in all_ids]}")
        
        conn.close()
        
        print(f"데이터베이스에 저장 완료: {specimen_id}")
        print(f"==================\n")
        
        # JSON 파일로도 저장
        json_saved = save_json_file(specimen_id, report_data)
        
        # 임시 파일 삭제
        os.remove(temp_file_path)
        
        return JSONResponse({
            "success": True, 
            "specimen_id": specimen_id,
            "json_saved": json_saved
        })
    
    except Exception as e:
        # 오류 발생 시 임시 파일 삭제
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)
        
        print(f"엑셀 업로드 중 오류 발생: {str(e)}")
        traceback.print_exc()
        
        return JSONResponse({"success": False, "error": str(e)})

@app.get("/api/reports")
async def get_reports():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute("SELECT specimen_id, created_at FROM reports ORDER BY created_at DESC")
    rows = cursor.fetchall()
    
    reports = [{"specimen_id": row["specimen_id"], "created_at": row["created_at"]} for row in rows]
    
    conn.close()
    
    # JSON 파일 목록도 확인
    json_files = []
    if os.path.exists(JSON_DIR):
        json_files = [f for f in os.listdir(JSON_DIR) if f.endswith('.json')]
    
    print(f"\n=== 보고서 목록 호출 ===")
    print(f"DB에 저장된 보고서 수: {len(reports)}")
    print(f"JSON 파일 수: {len(json_files)}")
    print(f"DB specimen_ids: {[r['specimen_id'] for r in reports[:10]]}")
    print(f"JSON 파일명: {json_files[:10]}")
    print(f"==================\n")
    
    return JSONResponse({"success": True, "reports": reports, "json_files": json_files})

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True, workers=1)