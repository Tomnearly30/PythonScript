#conding=utf8  
import os
import xlrd
import xlwt

###############表格写入功能#################################################
def WriteinExcel(n,file):
    ################Sample ID 在'RawMolecules_report.txt'第3行#############
    SampleID = file[2][18:25]
    worksheet.write(n, 0, SampleID)
    ################Time在'RawMolecules_report.txt'第6行############
    Time = file[5][18:].split('\n')[0]
    worksheet.write(n, 1, Time)
    ################Total DNA (>= 20 kbp)(Gbp）在'RawMolecules_report.txt'第10行############
    DataSize = file[9][40:55]
    worksheet.write(n, 2, DataSize)
    ################N50 (>= 20 kbp)）在'RawMolecules_report.txt'第11行############
    N50_20kbp = file[10][40:50]
    worksheet.write(n, 3, N50_20kbp)
    ################Total DNA (>= 150 kbp)(Gbp）在'RawMolecules_report.txt'第10行############
    DataSize2 = file[11][40:55]
    worksheet.write(n, 4, DataSize2)
    ################N50 (>= 150 kbp)）在'RawMolecules_report.txt'第11行############
    N50_150kbp = file[12][40:50]
    worksheet.write(n, 5, N50_150kbp)
    ################LD在'RawMolecules_report.txt'第22行############
    LD = file[21][40:45]
    worksheet.write(n, 6, LD)
    ################MapRate在'RawMolecules_report.txt'第27行############
    MapRate = file[26][31:35]
    worksheet.write(n, 7, MapRate)
    ################Effective coverage在'RawMolecules_report.txt'第28行############
    Effective_Cover = file[27][31:36]
    worksheet.write(n, 8, Effective_Cover)
    ################Reference在'RawMolecules_report.txt'第23行############
    Reference = file[22][40:66]
    worksheet.write(n, 9, Reference)
    ################Positive label variance在'RawMolecules_report.txt'第30行############
    PLV = file[29][31:35]
    worksheet.write(n, 10, PLV)
    ################Negative label variance在'RawMolecules_report.txt'第31行############
    NLV = file[30][31:35]
    worksheet.write(n, 11, NLV)
    ################JobID在'RawMolecules_report.txt'第5行############
    JobID = file[4][18:23]
    worksheet.write(n, 12, JobID)



###############################################
#################设置表头########################
###############################################

workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
worksheet.write(0, 0, 'Sample Information')
worksheet.write(0, 1, 'Imaging Time')
worksheet.write(0, 2, 'Total DNA (>= 20 kbp)')
worksheet.write(0, 3, 'N50_20kbp')
worksheet.write(0, 4, 'Total DNA (>= 150 kbp)')
worksheet.write(0, 5, 'N50_150kbp')
worksheet.write(0, 6, 'LD(>= 150 kbp)(/100 kbp)')
worksheet.write(0, 7, 'Bionano Map Rate (%)')
worksheet.write(0, 8, 'Effective coverage')
worksheet.write(0, 9, 'Reference')
worksheet.write(0, 10, 'Positive label variance')
worksheet.write(0, 11, 'Negative label variance')
worksheet.write(0, 12, 'Job ID')

##################################################################

n = 1
for dirpath,dirnames, filenames in os.walk('./'):
    if str(dirpath) != str("./"):
       if os.path.isfile(dirpath+"/"+"RawMolecules_report.txt") == True:
          f = open(dirpath+"/"+"RawMolecules_report.txt")
          data1 = f.readlines()
          WriteinExcel(n, data1)
          n = n + 1


workbook.save('Bionano_Record.xls')
####################################################################

	