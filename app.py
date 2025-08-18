from fastapi import FastAPI, Request, Form, File, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import os
import sqlite3
import json
import traceback
import time
from utils import NGS_EXCEL2DB

app = FastAPI()

# 현재 스크립트 파일의 경로
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))

# SQLite 데이터베이스 설정
DB_PATH = os.path.join(BASE_DIR, "ngs_reports.db")
JSON_DIR = os.path.join(BASE_DIR, "json")
TMP_DIR = os.path.join(BASE_DIR, "tmp")

# 디렉토리 확인 및 생성
if not os.path.exists(JSON_DIR):
    os.makedirs(JSON_DIR)

if not os.path.exists(TMP_DIR):
    os.makedirs(TMP_DIR)
    print(f"임시 파일 디렉토리 생성: {TMP_DIR}")

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
            print(f"임시 파일 삭제 완료: {file_path}")
            return True
        except PermissionError as e:
            if attempt < max_retries - 1:
                print(f"파일 삭제 재시도 {attempt + 1}/{max_retries}: {file_path}")
                time.sleep(delay)
            else:
                print(f"파일 삭제 실패 (최대 재시도 도달): {file_path} - {e}")
                return False
        except Exception as e:
            print(f"파일 삭제 중 예상치 못한 오류: {file_path} - {e}")
            return False
    
    return False

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
async def show_report(request: Request, specimen_id: str):
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
    temp_file_path = os.path.join(TMP_DIR, f"temp_{file.filename}")
    
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
        
        # SNV 데이터 - 수정된 값 사용 (분할 정보 포함)
        snv_clinical_highlight, snv_clinical_rows = parser.get_SNV('VCS')
        snv_clinical_processed = process_table_data_with_split_info(snv_clinical_rows, max_rows_first_page=8)
        report_data["snv_clinical"] = {
            "highlight": snv_clinical_highlight,
            **snv_clinical_processed
        }
        
        snv_unknown_highlight, snv_unknown_rows = parser.get_SNV('VUS')
        snv_unknown_processed = process_table_data_with_split_info(snv_unknown_rows, max_rows_first_page=8)
        report_data["snv_unknown"] = {
            "highlight": snv_unknown_highlight,
            **snv_unknown_processed
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
        
        # CNV 데이터 - 수정된 값 사용 (분할 정보 포함)
        cnv_clinical_highlight, cnv_clinical_rows = parser.get_CNV('VCS')
        cnv_clinical_processed = process_table_data_with_split_info(cnv_clinical_rows, max_rows_first_page=10)
        report_data["cnv_clinical"] = {
            "highlight": cnv_clinical_highlight,
            **cnv_clinical_processed
        }
        
        cnv_unknown_highlight, cnv_unknown_rows = parser.get_CNV('VUS')
        cnv_unknown_processed = process_table_data_with_split_info(cnv_unknown_rows, max_rows_first_page=10)
        report_data["cnv_unknown"] = {
            "highlight": cnv_unknown_highlight,
            **cnv_unknown_processed
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
        
        # Excel 파일 닫기 (파일 잠금 해제)
        parser.close()
        
        # 임시 파일 삭제
        safe_remove_file(temp_file_path)
        
        return JSONResponse({
            "success": True, 
            "specimen_id": specimen_id,
            "json_saved": json_saved
        })
    
    except Exception as e:
        # 오류 발생 시 Excel 파일 닫기 및 임시 파일 삭제
        try:
            if 'parser' in locals():
                parser.close()
        except:
            pass
        
        if os.path.exists(temp_file_path):
            safe_remove_file(temp_file_path)
        
        print(f"엑셀 업로드 중 오류 발생: {str(e)}")
        traceback.print_exc()
        
        return JSONResponse({"success": False, "error": str(e)})

@app.get("/api/search")
async def search_reports(q: str = ""):
    if not q or len(q.strip()) < 1:
        return JSONResponse({"success": True, "results": []})
    
    search_term = q.strip().lower()
    
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
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
                "signed1": diagnosis_user.get("Signed by", "N/A").split(", ")[1] if ", " in diagnosis_user.get("Signed by", "") else diagnosis_user.get("Signed by", "N/A")
            }
            results.append(result_item)
        except (json.JSONDecodeError, KeyError) as e:
            print(f"데이터 파싱 오류: {e}")
            continue
    
    conn.close()
    
    return JSONResponse({"success": True, "results": results})

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

@app.get("/api/specification/{panel_type}")
async def get_specification(panel_type: str):
    """
    패널 타입에 따른 Specification HTML 반환
    """
    if panel_type not in ['GE', 'SA']:
        return JSONResponse({"error": "Invalid panel type"}, status_code=400)
    
    spec_file = os.path.join(BASE_DIR, "templates", f"{panel_type}_Specification.html")
    
    if not os.path.exists(spec_file):
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
    
    gene_content_file = os.path.join(BASE_DIR, "templates", f"{content_type}.html")
    
    if not os.path.exists(gene_content_file):
        return JSONResponse({"error": "Gene content file not found"}, status_code=404)
    
    return FileResponse(gene_content_file, media_type="text/html")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=1234, reload=True, workers=1)