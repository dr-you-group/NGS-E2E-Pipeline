from utils import NGS_EXCEL2DB

ngs = NGS_EXCEL2DB('sample_data/Master_table_SS2430925.xlsx')

print(ngs.get_Clinical_Info())

print(ngs.get_SNV('VCS'))
print(ngs.get_Fusion('VCS'))
print(ngs.get_CNV('VCS'))
print(ngs.get_LR_BRCA('VCS'))
print(ngs.get_Splice('VCS'))

print(ngs.get_Biomarkers())

print(ngs.get_SNV('VUS'))
print(ngs.get_Fusion('VUS'))
print(ngs.get_CNV('VUS'))
print(ngs.get_LR_BRCA('VUS'))
print(ngs.get_Splice('VUS'))


print(ngs.get_Failed_Gene())
print(ngs.get_Comments())
print(ngs.get_Diagnostic_Info())
print(ngs.get_Filter_History())
print(ngs.get_DRNA_Qubit_Density())
print(ngs.get_QC())
print(ngs.get_Analysis_Program())
print(ngs.get_Diagnosis_User_Registration())