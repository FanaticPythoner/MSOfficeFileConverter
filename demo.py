from MSOfficeFileConverter import WordDocument

#Creating the WordDocument object

document = WordDocument('Example\Path\To\file.docx')

#Exporting to PDF

document.toPdf('Example\Export\Path','OutputFileName')