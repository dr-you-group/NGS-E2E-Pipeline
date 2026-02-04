
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

    # 3. 변이 데이터
    # SNV
    h, rows = parser.get_SNV('VCS')
    report_data["snv_clinical"] = {"highlight": h, **process_table_data(rows)}

    h, rows = parser.get_SNV('VUS')
    report_data["snv_unknown"] = {"highlight": h, **process_table_data(rows)}

    # Fusion
    h, rows = parser.get_Fusion('VCS')
    report_data["fusion_clinical"] = {"highlight": h, **process_table_data(rows)}

    h, rows = parser.get_Fusion('VUS')
    report_data["fusion_unknown"] = {"highlight": h, **process_table_data(rows)}

    # CNV
    h, rows = parser.get_CNV('VCS')
    report_data["cnv_clinical"] = {"highlight": h, **process_table_data(rows)}

    h, rows = parser.get_CNV('VUS')
    report_data["cnv_unknown"] = {"highlight": h, **process_table_data(rows)}

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
