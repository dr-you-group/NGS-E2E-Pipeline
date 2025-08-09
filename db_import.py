import os
import sys
import json
import sqlite3
import traceback
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from utils import NGS_EXCEL2DB

# 현재 스크립트 파일의 경로
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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

def process_table_data(rows, name=""):
    """
    rows 데이터를 headers와 data로 분리하는 함수
    rows[0]은 headers로, rows[1:]는 data로 변환
    """
    print(f"\n처리 중인 테이블: {name}")
    print(f"원본 rows: {rows}")
    
    if not rows or len(rows) < 1:
        print(f"{name}: 데이터가 없거나 부족합니다.")
        return {"headers": [], "data": []}
    
    headers = rows[0]
    print(f"추출된 headers: {headers}")
    
    data = rows[1:] if len(rows) > 1 else []
    print(f"추출된 data 개수: {len(data)}")
    if data:
        print(f"첫 번째 data 행: {data[0] if data else '없음'}")
    
    return {"headers": headers, "data": data}

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

def extract_variant_data(parser, variant_type, clinical_significance, name):
    """변이 데이터를 추출하고 디버깅하는 함수"""
    print(f"\n{name} ({clinical_significance}):")
    
    try:
        if variant_type == 'SNV':
            highlight, rows = parser.get_SNV(clinical_significance)
        elif variant_type == 'Fusion':
            highlight, rows = parser.get_Fusion(clinical_significance)
        elif variant_type == 'CNV':
            highlight, rows = parser.get_CNV(clinical_significance)
        elif variant_type == 'LR_BRCA':
            highlight, rows = parser.get_LR_BRCA(clinical_significance)
        elif variant_type == 'Splice':
            highlight, rows = parser.get_Splice(clinical_significance)
        else:
            print(f"  알 수 없는 변이 타입: {variant_type}")
            return None
        
        print(f"  Highlight: {highlight}")
        print(f"  Rows 타입: {type(rows)}")
        print(f"  Rows 길이: {len(rows) if rows is not None else 'None'}")
        if rows and len(rows) > 0:
            print(f"  Headers: {rows[0]}")
            if len(rows) > 1:
                print(f"  첫 번째 데이터 행: {rows[1]}")
        else:
            print("  데이터 행 없음")
        
        processed = process_table_data(rows, f"{name}_{clinical_significance}")
        return {
            "highlight": highlight,
            **processed
        }
    except Exception as e:
        print(f"  {name} 데이터 추출 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return {
            "highlight": "",
            "headers": [],
            "data": [],
            "error": str(e)
        }

def excel_to_db(excel_file_path):
    try:
        init_db()
        
        # NGS_EXCEL2DB 클래스 사용하여 데이터 파싱
        parser = NGS_EXCEL2DB(excel_file_path)
        
        # 보고서 데이터 생성
        report_data = {}
        
        # 검체 정보
        report_data["clinical_info"] = parser.get_Clinical_Info()
        
        # SNV 데이터 - 수정된 값 사용
        report_data["snv_clinical"] = extract_variant_data(parser, 'SNV', 'VCS', 'SNV')
        report_data["snv_unknown"] = extract_variant_data(parser, 'SNV', 'VUS', 'SNV')
        
        # Fusion 데이터 - 수정된 값 사용
        report_data["fusion_clinical"] = extract_variant_data(parser, 'Fusion', 'VCS', 'Fusion')
        report_data["fusion_unknown"] = extract_variant_data(parser, 'Fusion', 'VUS', 'Fusion')
        
        # CNV 데이터 - 수정된 값 사용
        report_data["cnv_clinical"] = extract_variant_data(parser, 'CNV', 'VCS', 'CNV')
        report_data["cnv_unknown"] = extract_variant_data(parser, 'CNV', 'VUS', 'CNV')
        
        # LR BRCA 데이터 - 수정된 값 사용
        report_data["lr_brca_clinical"] = extract_variant_data(parser, 'LR_BRCA', 'VCS', 'LR_BRCA')
        report_data["lr_brca_unknown"] = extract_variant_data(parser, 'LR_BRCA', 'VUS', 'LR_BRCA')
        
        # Splice 데이터 - 수정된 값 사용
        report_data["splice_clinical"] = extract_variant_data(parser, 'Splice', 'VCS', 'Splice')
        report_data["splice_unknown"] = extract_variant_data(parser, 'Splice', 'VUS', 'Splice')
        
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
        report_data["qc"] = process_table_data(qc_rows, "QC")
        
        # Analysis Program
        report_data["analysis_program"] = parser.get_Analysis_Program()
        
        # 사용자 정보
        report_data["diagnosis_user"] = parser.get_Diagnosis_User_Registration()
        
        # Panel Type
        report_data["panel_type"] = parser.panel
        
        # 데이터베이스에 저장
        specimen_id = report_data["clinical_info"]["검체 정보"]
        
        print(f"\n최종 데이터 확인:")
        for key, value in report_data.items():
            if isinstance(value, dict) and "data" in value:
                print(f"{key}: {len(value.get('data', []))}개 데이터 행")
        
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        cursor.execute(
            "INSERT OR REPLACE INTO reports (specimen_id, report_data) VALUES (?, ?)",
            (specimen_id, json.dumps(report_data))
        )
        
        conn.commit()
        conn.close()
        
        print(f"데이터베이스에 저장 완료: {specimen_id}")
        
        # JSON 파일로도 저장
        json_saved = save_json_file(specimen_id, report_data)
        
        return {"success": True, "specimen_id": specimen_id, "json_saved": json_saved}
    except Exception as e:
        print(f"엑셀 파싱 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return {"success": False, "error": str(e)}

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python db_import.py <excel_file_path>")
        sys.exit(1)
    
    excel_file_path = sys.argv[1]
    result = excel_to_db(excel_file_path)
    
    if result["success"]:
        json_status = "JSON 파일 저장 성공" if result["json_saved"] else "JSON 파일 저장 실패"
        print(f"성공적으로 데이터를 가져왔습니다. 검체 ID: {result['specimen_id']} ({json_status})")
    else:
        print(f"오류: {result['error']}")