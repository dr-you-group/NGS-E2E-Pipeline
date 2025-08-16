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
        return {
            "검체 정보": self.clinical_dict["병리번호"],
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
    def get_SNV(self, data_type: str) -> Tuple[str, List]:
        SNV_Data = self.SNV[self.SNV['Clinical_significance'] == data_type]
        SNV_Highlight = ', '.join(SNV_Data[SNV_Data["highlight"] != '']['highlight'].tolist())
        SNV_Row = ['Gene', 'Consequence', 'AA Change', 'VAF', 'HGVSc', 'HGVSp']
        # VAF 값 소수 2번째 자리까지 반올림 처리
        SNV_Data_processed = SNV_Data[SNV_Row].copy()
        SNV_Data_processed['VAF'] = SNV_Data_processed['VAF'].apply(
            lambda x: f"{float(x):.2f}" if pd.notna(x) else x
        )
        return SNV_Highlight, [SNV_Row] + SNV_Data_processed.values.tolist()
    

    # Fusion Gene
    def get_Fusion(self, data_type:str) -> Tuple[str, List]:
        Fusion_Data = self.Fusion[self.Fusion['Clinical_significance'] == data_type]
        Fusion_Highlight = ', '.join(Fusion_Data[Fusion_Data["highlight"] != '']['highlight'].tolist())
        Fusion_Row = ['Gene fusion', 'Breakpoint 1', 'Breakpoint 2', 'Fusion supporting reads']
        return Fusion_Highlight, [Fusion_Row] + Fusion_Data[Fusion_Row].values.tolist()
    

    # Copy number variation
    def get_CNV(self, data_type:str) -> Tuple[str, List]:
        CNV_Data = self.CNV[self.CNV['Clinical_significance'] == data_type]
        CNV_Highlight = ', '.join(CNV_Data[CNV_Data["highlight"] != '']['highlight'].tolist())
        CNV_Row = ['Gene', 'Location', 'Fold Change', 'Estimated copy number']
        return CNV_Highlight, [CNV_Row] + CNV_Data[CNV_Row].values.tolist()


    # Large rearrangements in BRCA1/2
    def get_LR_BRCA(self, data_type: str) -> Tuple[str, List]:
        LR_BRCA_Data = self.LR_BRCA[self.LR_BRCA['Clinical_significance'] == data_type]
        LR_BRCA_Data_Highlight = ', '.join(LR_BRCA_Data[LR_BRCA_Data["highlight"] != '']['highlight'].tolist())
        LR_BRCA_Row = ['Gene', 'Location', 'Affected exon', 'Fold Change', 'Estimated copy number']
        return LR_BRCA_Data_Highlight, [LR_BRCA_Row] + LR_BRCA_Data[LR_BRCA_Row].values.tolist()


    # Splice Variant
    def get_Splice(self, data_type: str) -> Tuple[str, List]:
        Splice_Data = self.Splice[self.Splice['Clinical_significance'] == data_type]
        Splice_Highlight = ', '.join(Splice_Data[Splice_Data["highlight"] != '']['highlight'].tolist())
        Splice_Row = ['Gene', 'Affected exon', 'Breakpoint 1', 'Breakpoint 2', 'Splice supporting reads']
        return Splice_Highlight, [Splice_Row] + Splice_Data[Splice_Row].values.tolist()
    

    # Other BioMarkers
    def get_Biomarkers(self) -> Dict:
        return {
            'TMB': self.IO['Value'][0] + ' /Megabase',
            'MSI': self.IO['Value'][7] + ' %',
        }

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