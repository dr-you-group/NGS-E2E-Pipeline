import pandas as pd
from typing import Tuple, List, Dict, Any
import warnings
warnings.filterwarnings('ignore')

class NGS_EXCEL2DB:
    def __init__(self, file):
        self.df = pd.ExcelFile(file, engine='openpyxl')
        self._file_path = file
        self.Clinical_Information = self.df.parse('clinical_information', header=None, dtype=str).fillna('')
        self.clinical_dict = dict(zip(self.Clinical_Information.iloc[:, 0], self.Clinical_Information.iloc[:, 1]))
        self.NGS_QC = self.df.parse('NGS_QC', header=None, dtype=str).fillna('')
        self.SNV = self.df.parse('SNV', dtype=str).fillna('')
        self.CNV = self.df.parse('CNV', dtype=str).fillna('')
        
        # CNVarm 시트를 안전하게 로드
        try:
            self.CNVarm = self.df.parse('CNVarm', dtype=str).fillna('')
        except Exception as e:
            print(f"CNVarm 시트 로드 실패, 빈 DataFrame 사용: {e}")
            self.CNVarm = pd.DataFrame()
            
        self.CNV_allFC = self.df.parse('CNV_allFC', dtype=str).fillna('')
        self.LR_BRCA = self.df.parse('LR_BRCA', dtype=str).fillna('')
        self.Fusion = self.df.parse('Fusion', header=1, dtype=str).fillna('')
        self.Splice = self.df.parse('Splice', dtype=str).fillna('')
        self.IO = self.df.parse('IO', header=1, dtype=str).fillna('')
        self.panel = 'SA' if '.SA.' in self.clinical_dict["검체 유형"] else 'GE'
        
        # V2 리포트 여부 확인 (NGS_QC 시트 E4 셀)
        self.is_v2 = False
        try:
            if self.NGS_QC.shape[0] > 3 and self.NGS_QC.shape[1] > 4:
                cell_e4 = str(self.NGS_QC.iloc[3, 4]).strip()
                if cell_e4 == "TSO500_v2":
                    self.is_v2 = True
        except Exception as e:
            print(f"V2 판별 중 오류 발생: {e}")

    def _parse_highlight_structure(self, highlight_text: str, gene_names: List[str] = None) -> List[Dict[str, Any]]:
        if not highlight_text:
            return []

        import re

        # 유전자명 기반 정규식 패턴 준비
        gene_pattern = None
        if gene_names:
            # 개별 유전자명 추출 (Fusion의 '-' 구분자 분리 포함), 빈 문자열 제거
            individual_genes = set()
            for g in gene_names:
                if not g:
                    continue
                for part in g.split('-'):
                    stripped = part.strip()
                    if stripped:
                        individual_genes.add(stripped)

            if individual_genes:
                # 길이 순 정렬 (긴 것 우선 매칭)
                sorted_genes = sorted(individual_genes, key=len, reverse=True)
                escaped = '|'.join(re.escape(g) for g in sorted_genes)
                # 유전자명이 ::로 연결된 경우도 하나의 이탤릭 블록으로 처리
                gene_pattern = re.compile(f'((?:{escaped})(?:::(?:{escaped}))*)')

        structured_items = []
        items = highlight_text.split(', ')

        for idx, item in enumerate(items):
            if gene_pattern:
                segments = self._split_by_gene_pattern(item, gene_pattern)
                structured_items.extend(segments)
            else:
                # 폴백: 첫 공백 기준 분리 (기존 로직)
                parts = item.split(' ', 1)
                structured_items.append({"text": parts[0], "style": "italic"})
                if len(parts) > 1:
                    structured_items.append({"text": " " + parts[1], "style": "normal"})

            # 항목 간 구분자 (쉼표)
            if idx < len(items) - 1:
                structured_items.append({"text": ", ", "style": "normal"})

        return structured_items

    def _split_by_gene_pattern(self, text: str, gene_pattern) -> List[Dict[str, Any]]:
        """텍스트에서 유전자명을 찾아 이탤릭/정자체 세그먼트로 분리"""
        segments = []
        last_end = 0

        for match in gene_pattern.finditer(text):
            start, end = match.span()
            if start > last_end:
                segments.append({"text": text[last_end:start], "style": "normal"})
            segments.append({"text": match.group(), "style": "italic"})
            last_end = end

        if last_end < len(text):
            segments.append({"text": text[last_end:], "style": "normal"})

        # 매칭 없으면 폴백 (첫 공백 기준)
        if not segments:
            parts = text.split(' ', 1)
            segments = [{"text": parts[0], "style": "italic"}]
            if len(parts) > 1:
                segments.append({"text": " " + parts[1], "style": "normal"})

        return segments

        return structured_items
    
    def close(self):
        """Excel 파일을 명시적으로 닫아 파일 잠금을 해제합니다."""
        try:
            if hasattr(self.df, 'close'):
                self.df.close()
            print(f"Excel 파일 닫기 완료: {self._file_path}")
        except Exception as e:
            print(f"Excel 파일 닫기 실패: {e}")
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
    

    # 검체정보
    def get_Clinical_Info(self) -> Dict:
        sub_block = self.clinical_dict.get("Sub.", "").strip()
        specimen_display = f'{self.clinical_dict["병리번호"]} {sub_block}' if sub_block else self.clinical_dict["병리번호"]
        return {
            "검체 정보": specimen_display,
            "성별": self.clinical_dict["성별"],
            "나이": self.clinical_dict["나이"],
            "Unit NO.": self.clinical_dict["Unit NO."],
            "환자명": self.clinical_dict["환자명"],
            "채취 장기": self.clinical_dict["채취 장기"],
            "원발 장기": self.clinical_dict["원발 장기"],
            "진단":self.clinical_dict["진단"],
            "의뢰의": self.clinical_dict["의뢰의"],
            "의뢰의 소속": self.clinical_dict["의뢰의 소속"],
            "검체 유형": self.clinical_dict["검체 유형"].split('.')[0],
            "검체의 적절성여부": self.clinical_dict["검체의 적절성"],
            "검체접수일": self.clinical_dict["검체접수일"],
            "결과보고일": self.clinical_dict["Report date"],
        }
    

    # SNVs & Indels
    def get_SNV(self, data_type: str) -> Tuple[List[Dict[str, List]], List]:
        SNV_Data = self.SNV[self.SNV['Clinical_significance'] == data_type]
        raw_highlight = ', '.join(SNV_Data[SNV_Data["highlight"] != '']['highlight'].tolist())
        gene_names = SNV_Data['Gene'].tolist()
        SNV_Highlight = self._parse_highlight_structure(raw_highlight, gene_names)
        SNV_Row = ['Gene', 'Consequence', 'AA Change', 'VAF', 'HGVSc', 'HGVSp']
        # VAF 값 소수 2번째 자리까지 반올림 처리
        SNV_Data_processed = SNV_Data[SNV_Row].copy()
        SNV_Data_processed['VAF'] = SNV_Data_processed['VAF'].apply(
            lambda x: f"{float(x):.2f}" if pd.notna(x) else x
        )
        return SNV_Highlight, [SNV_Row] + SNV_Data_processed.values.tolist()
    

    # Fusion Gene
    def get_Fusion(self, data_type:str) -> Tuple[List[Dict[str, List]], List]:
        Fusion_Data = self.Fusion[self.Fusion['Clinical_significance'] == data_type]
        raw_highlight = ', '.join(Fusion_Data[Fusion_Data["highlight"] != '']['highlight'].tolist())
        gene_names = Fusion_Data['Gene fusion'].tolist()
        Fusion_Highlight = self._parse_highlight_structure(raw_highlight, gene_names)
        Fusion_Row = ['Gene fusion', 'Breakpoint 1', 'Breakpoint 2', 'Fusion supporting reads']
        return Fusion_Highlight, [Fusion_Row] + Fusion_Data[Fusion_Row].values.tolist()
    

    # Copy number variation
    def get_CNV(self, data_type:str) -> Tuple[List[Dict[str, List]], List]:
        CNV_Data = self.CNV[self.CNV['Clinical_significance'] == data_type]
        raw_highlight = ', '.join(CNV_Data[CNV_Data["highlight"] != '']['highlight'].tolist())
        gene_names = CNV_Data['Gene'].tolist()
        CNV_Highlight = self._parse_highlight_structure(raw_highlight, gene_names)
        CNV_Row = ['Gene', 'Location', 'Fold Change', 'Estimated copy number']
        return CNV_Highlight, [CNV_Row] + CNV_Data[CNV_Row].values.tolist()


    # Large rearrangements in BRCA1/2
    def get_LR_BRCA(self, data_type: str) -> Tuple[List[Dict[str, List]], List]:
        LR_BRCA_Data = self.LR_BRCA[self.LR_BRCA['Clinical_significance'] == data_type]
        raw_highlight = ', '.join(LR_BRCA_Data[LR_BRCA_Data["highlight"] != '']['highlight'].tolist())
        gene_names = LR_BRCA_Data['Gene'].tolist()
        LR_BRCA_Data_Highlight = self._parse_highlight_structure(raw_highlight, gene_names)
        LR_BRCA_Row = ['Gene', 'Location', 'Affected exon', 'Fold Change', 'Estimated copy number']
        return LR_BRCA_Data_Highlight, [LR_BRCA_Row] + LR_BRCA_Data[LR_BRCA_Row].values.tolist()


    # Splice Variant
    def get_Splice(self, data_type: str) -> Tuple[List[Dict[str, List]], List]:
        Splice_Data = self.Splice[self.Splice['Clinical_significance'] == data_type]
        raw_highlight = ', '.join(Splice_Data[Splice_Data["highlight"] != '']['highlight'].tolist())
        gene_names = Splice_Data['Gene'].tolist()
        Splice_Highlight = self._parse_highlight_structure(raw_highlight, gene_names)
        Splice_Row = ['Gene', 'Affected exon', 'Breakpoint 1', 'Breakpoint 2', 'Splice supporting reads']
        return Splice_Highlight, [Splice_Row] + Splice_Data[Splice_Row].values.tolist()
    

    # Other BioMarkers
    def get_Biomarkers(self) -> Dict:
        """
        IO 시트에서 TMB, MSI 값을 추출합니다.
        Value: TMB(Row 0), MSI(Row 7)
        Status (Categorical): TMB(Row 3), MSI(Row 10) - High/Low 여부 판단용
        """
        biomarkers = {
            'TMB': {
                'value': self.IO['Value'][0],
                'unit': '/Megabase',
                'status': self.IO['Value'][3]  # Categorical Result (High/Low/Stable etc)
            },
            'MSI': {
                'value': self.IO['Value'][7],
                'unit': '%',
                'status': self.IO['Value'][10], # Categorical Result (High/Low/Stable etc)
                'usable_msi_sites': self.IO['Value'][9]  # Usable MSI Sites (B12)
            }
        }
        
        # V2 전용 추가: Tumor Fraction (B18), Ploidy (B19)
        if self.is_v2:
            try:
                # tumor % 키 디버깅
                tumor_keys = [k for k in self.clinical_dict.keys() if 'tumor' in str(k).lower()]
                print(f"[DEBUG] clinical_dict tumor-related keys: {tumor_keys}")
                pathological_val = ''
                for k in self.clinical_dict:
                    if 'tumor' in str(k).lower():
                        pathological_val = str(self.clinical_dict[k]).strip()
                        print(f"[DEBUG] Found tumor key: '{k}' -> '{pathological_val}'")
                        break
                
                biomarkers['Tumor_Fraction'] = {
                    'value': self.IO['Value'][15], # Row 15 -> B18 (SNP based estimation)
                    'pathological': pathological_val, # Pathological estimation
                    'unit': ''
                }
                biomarkers['Ploidy'] = {
                    'value': self.IO['Value'][16], # Row 16 -> B19
                    'unit': ''
                }
                biomarkers['GIS'] = {
                    'value': self.IO['Value'][14], # Row 14 -> B17 (Genomic Instability Score)
                    'unit': ''
                }
            except Exception as e:
                print(f"V2 Biomarker (Tumor Fraction/Ploidy) 추출 실패: {e}")
                
        return biomarkers

    # Failed_Gene
    def get_Failed_Gene(self) -> str:
        return 'None'


    # Comments
    def get_Comments(self) -> List:
        SNV_Comments = self.SNV[self.SNV["Comment"] != '']["Comment"].tolist()
        CNV_Comments = self.CNV[self.CNV["Comment"] != '']["Comment"].tolist()
        # CNVarm_Comments = self.CNVarm[self.CNVarm["Comment"] != '']["Comment"].tolist()
        LR_BRCA_Comments = self.LR_BRCA[self.LR_BRCA["Comment"] != '']["Comment"].tolist()
        Fusion_Comments = self.Fusion[self.Fusion["Comment"] != '']["Comment"].tolist() 
        Splice_Comments = self.Splice[self.Splice["comment"] != '']["comment"].tolist()
        Comments_List = SNV_Comments+CNV_Comments+LR_BRCA_Comments+Fusion_Comments+Splice_Comments
        return Comments_List
    

    # Sequence Date
    def get_Sequence_Date(self) -> str:
        """분석 일자 (NGS_QC B2 셀) 추출"""
        try:
            if self.NGS_QC.shape[0] > 1 and self.NGS_QC.shape[1] > 1:
                val = self.NGS_QC.iloc[1, 1]
                if pd.isna(val) or str(val).strip() == '':
                    return ""
                return str(val).strip()
            return ""
        except Exception as e:
            print(f"Sequence Date 추출 실패: {e}")
            return ""


    # Run Name (V2 전용)
    def get_Run_Name(self) -> str:
        """Run Name (NGS_QC B1 셀) 추출 - V2 전용"""
        if not self.is_v2:
            return ""
        try:
            if self.NGS_QC.shape[0] > 0 and self.NGS_QC.shape[1] > 1:
                val = self.NGS_QC.iloc[0, 1]
                if pd.isna(val) or str(val).strip() == '':
                    return ""
                return str(val).strip()
            return ""
        except Exception as e:
            print(f"Run Name 추출 실패: {e}")
            return ""
    

    # 검사 정보
    def get_Diagnostic_Info(self) -> Dict:
        InstrumentType = self.NGS_QC[4][1] + " Dx [Illumina]"
        if self.panel == 'GE':
            di = {
                "검사시약":"AllPrep DNA/RNA FFPE Kit (50) [Qiagen], TruSight™ Oncology 500 kit [Illumina]",
                "검사방법":"NGS targeted DNA/RNA sequencing (Library : Hybrid capture)",
                "검사기기":InstrumentType,
                "Reference genome": "Homo_sapiens/ UCSC/ hg19"
            }
        else:
            di = {
                "검사시약":"AllPrep DNA/RNA FFPE Kit (50) [Qiagen], TruSight™ Oncology 500 kit [Illumina], TruSight™ RNA Fusion Panel [Illumina]",
                "검사방법":"NGS targeted DNA/RNA sequencing (Library : Hybrid capture)",
                "검사기기":InstrumentType,
                "Reference genome": "Homo_sapiens/ UCSC/ hg19"
            }
        return di
    

    # Filter History
    def get_Filter_History(self) -> Dict:
        return {
            'Include': 'Exonic, Illumina Q.C Filter PASS, Fold change <0.5 or >1.5',
            'Exclude': 'Synonymous, VAF <3%, total depth <100, refer depth=0'
        }
    
    
    # DNA, RNA (Qubit 농도)
    def get_DRNA_Qubit_Density(self) -> Dict:
        return {
            'DNA': str(self.clinical_dict["DNA conc.(ng/ul)"]),
            'RNA': str(self.clinical_dict["RNA conc.(ng/ul)"])
        }


    # Q.C
    def get_QC(self) -> List[List]:
        return [
            ["Metric (UOM)", 'LSL Guideline', 'Value'],
            [self.NGS_QC[1][7], '80', self.NGS_QC[3][7]],
            [self.NGS_QC[1][8], '80', self.NGS_QC[3][8]],
            [self.NGS_QC[1][9], '80', self.NGS_QC[3][9]],
        ]
    

    # Analysis Program
    def get_Analysis_Program(self):
        return "DRAGEN TSO500 ( Workflow Version : 2.5.2 )"


    # Tested, Signed, Analyzed by, 분자 접수 번호
    def get_Diagnosis_User_Registration(self):
        return {
            'Tested by': f'{self.clinical_dict["Tester1"]}, {self.clinical_dict["Tester2"]}',
            'Signed by': f'{self.clinical_dict["Signed2"]}, {self.clinical_dict["Signed1"]}',
            'Analyzed by': '이청',
            '분자접수번호': self.clinical_dict["분자접수번호"]
        }


def split_variants_into_pages(variants_data: Dict[str, Any], max_items_per_page: int = 5) -> List[Dict[str, Any]]:
    """
    변이 데이터를 페이지별로 분할합니다.
    
    Args:
        variants_data: 변이 데이터 딕셔너리
        max_items_per_page: 페이지당 최대 항목 수
    
    Returns:
        페이지별로 분할된 데이터 리스트
    """
    pages = []
    current_page = {'clinical': {}, 'unknown': {}}
    item_count = 0
    
    # Clinical significance 변이들
    for key in ['snv_clinical', 'fusion_clinical', 'cnv_clinical', 'lr_brca_clinical', 'splice_clinical']:
        if key in variants_data and variants_data[key]['data']:
            if item_count + len(variants_data[key]['data']) > max_items_per_page:
                if current_page['clinical']:
                    pages.append(current_page)
                current_page = {'clinical': {}, 'unknown': {}}
                item_count = 0
            
            current_page['clinical'][key] = variants_data[key]
            item_count += len(variants_data[key]['data'])
    
    # Unknown significance 변이들
    for key in ['snv_unknown', 'fusion_unknown', 'cnv_unknown', 'lr_brca_unknown', 'splice_unknown']:
        if key in variants_data and variants_data[key]['data']:
            if item_count + len(variants_data[key]['data']) > max_items_per_page:
                if current_page['clinical'] or current_page['unknown']:
                    pages.append(current_page)
                current_page = {'clinical': {}, 'unknown': {}}
                item_count = 0
            
            current_page['unknown'][key] = variants_data[key]
            item_count += len(variants_data[key]['data'])
    
    # 마지막 페이지 추가
    if current_page['clinical'] or current_page['unknown']:
        pages.append(current_page)
    
    return pages