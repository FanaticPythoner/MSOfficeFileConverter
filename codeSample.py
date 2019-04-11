from MSOfficeFileConverter import WordDocument, ExcelDocument

#Creating the WordDocument object
document = WordDocument('Example\Path\To\file.docx')
#Exporting to PDF
document.toPdf('Example\Export\Path','OutputFileName')


#Creating the ExcelDocument object
document = ExcelDocument('Example\Path\To\file.xlsx')
#Exporting to PDF
document.toPdf('Example\Export\Path','OutputFileName')
