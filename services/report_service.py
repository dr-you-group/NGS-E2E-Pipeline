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
