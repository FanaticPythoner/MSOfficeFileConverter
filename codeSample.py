from MSOfficeFileConverter import WordDocument, ExcelDocument

# Creating the WordDocument object.
document = WordDocument('SampleWord.docx')

# Exporting to PDF.
document.toPdf('ExportFolder','OutputSampleWord.pdf')

# Exporting to Html. Notice the new folder created in ExportFolder.
document.toHtml('ExportFolder', 'OutputSampleWord.html')


# ======================================================================================================


# Creating the ExcelDocument object.
document = ExcelDocument('SampleExcel.xlsx')

# Exporting to PDF. Notice that even if we dont specify the file extension, the export still work.
document.toPdf('ExportFolder','SampleExcel.pdf')

# Exporting to XLS. Notice that if we dont give any arguments, the tool export the file in the same
# directory as the original Excel file and with the same name as the original Excel file.
document.toXls()

# Exporting to CSV. Notice the new folder created in ExportFolder with all the spreadsheets exported, 
# instead of only the selected one being exported as default in Excel.
document.toCsv('ExportFolder','SampleExcel.csv')
