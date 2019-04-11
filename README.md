# MSOfficeFileConverter
In 3 lines of code, this tool allow you to convert/export a Microsoft Office file to a specified format, without any other dependency. My code use the build-in function of Microsoft Office to export any Word/Excel/PowerPoint document. The only thing you need is the Microsoft Office program you want to convert from (EX. You need Word if you want to convert a .docx to .pdf).

That being said, if the "save as" function of a Microsoft Office software work, MSOfficeFileConverter will work. That is the exact reason why I made this library : Alternatives to export, for example, a Word document to PDF are generally buggy, bad, or damn expensive. MSOfficeFileConverter is free, extremely simple, and just work.

### Language: ### 

- Tested in Python 3.6, should work in all Python 3 version.

### Limitations: ###

- For now, MSOfficeFileConverter only work on Windows.
               
- For now, only work for Word and Excel. I am planning to implement PowerPoint in the following days.

### Table of Contents: ###

- Classes
  
  - [*WordDocument*](https://github.com/FanaticPythoner/MSOfficeFileConverter#worddocument-class)
  
  - [*ExcelDocument*](https://github.com/FanaticPythoner/MSOfficeFileConverter#exceldocument-class)

# Installation

- Download the repository

- Install the requirements in the requirements.txt file (pip install -r requirements.txt)

- Use the sample code below [*Usage / Code Sample*](https://github.com/FanaticPythoner/MSOfficeFileConverter#usage--code-sample-) in the [*WordDocument*](https://github.com/FanaticPythoner/MSOfficeFileConverter#worddocument-class) class below as an example. Enjoy.


# WordDocument Class

### Description : ###
Open a Word document from a specified file path, then offer methods to convert it to whatever format you want.

Currently support the export in the following formats:

- Docx
    
- Docx (Strict Open XML Document)
    
- Docm
    
- Doc
    
- Dotm
    
- Dot
    
- Pdf
    
- Xps
    
- Mht
    
- Mthml
    
- Html
    
- Html (Filtered)
    
- Htm
    
- Rtf
    
- Txt
    
- Xml
    
- Xml (Macro Enabled)
    
- Xml (2003)
    
- Odt

 
### Usage / Code sample : ###
*#### This example create a WordDocument object then convert it to PDF. Notice that we didn't specify the PDF file extension in the second parameter, and the file name output is still 'OutputFileName.pdf'. The result would be the same if the second parameter was 'OutputFileName.pdf'. Takes 3 lines of code. ####*
```python
from MSOfficeFileConverter import WordDocument
document = WordDocument('Example\\Path\\To\\file.docx')
document.toPdf('Example\\Export\\Path','OutputFileName')
```

*#### This example create a WordDocument object then convert it to PDF. Notice that we didn't specify anything in the toPdf() method. Doing that will force the WordDocument instance to export the file in the same directory as the original Word file (here 'Example\\Path\\To') and with the same name as the original Word file, except with the destination extension (here 'OutputFileName.pdf'). Takes 3 lines of code. ####*
```python
from MSOfficeFileConverter import WordDocument
document = WordDocument('Example\\Path\\To\\file.docx')
document.toPdf()
```

# ExcelDocument Class

### Description : ###
Open an Excel document from a specified file path, then offer methods to convert it to whatever format you want.

Currently support the export in the following formats:

- Xlsx

- Xlsm

- Xlsb

- Xls

- Xml (Data)

- Mht

- Mhtml

- Xltx

- Xltm

- Xlt

- Txt (Windows)

- Txt (Macintosh)

- Txt (Unicode)

- Txt (MSDOS)

- Csv (UTF-8)

- Csv (Windows)

- Csv (Macintosh)

- Csv (Unicode)

- Csv (MSDOS)

- Xml (Spreadsheet 2003)

- Xls (Excel 5.0/95 Workbook)

- Prn

- Dif

- Slk

- Xlam

- Xla

- Pdf

- Xps

- Xlsx (Strict Open XML Spreadsheet)

- Ods

 
### Usage / Code sample : ###
*#### This example create an ExcelDocument object then convert it to PDF. Takes 3 lines of code. ####*
```python
from MSOfficeFileConverter import ExcelDocument
document = ExcelDocument('Example\\Path\\To\\file.xlsx')
document.toPdf('Example\\Export\\Path','OutputFileName.pdf')
```

*#### This example create an ExcelDocument object then convert in batch all sheets to CSV. All the CSV files are stored in a new folder named CSV_Files_file.xlsx: The output folder path then changes to 'Example\\Export\\Path\\CSV_Files_file.xlsx'. Takes 3 lines of code. ####*
```python
from MSOfficeFileConverter import ExcelDocument
document = ExcelDocument('Example\\Path\\To\\file.xlsx')
document.toCsv('Example\\Export\\Path','OutputFileName')
```
