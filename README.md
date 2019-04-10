# MSOfficeFileConverter
Allow you to convert/export a Microsoft Office file to a specified format, without any other dependency. My code use the build-in function of Microsoft Office to export any Word/Excel/PowerPoint document. The only thing you need is the Microsoft Office program you want to convert from (EX. You need Word if you want to convert a .docx to .pdf).

That being said, if the "save as" function of a Microsoft Office software work, MSOfficeFileConverter will work. That is the exact reason why I made this library : Alternatives to export, for example, a Word document to PDF are generally buggy, bad, or damn expensive. MSOfficeFileConverter is free, extremely simple, and just work.

### Language: ### 

- Tested in Python 3.6, should work in all Python 3 version.

### Limitations: ###

- For now, MSOfficeFileConverter only work on Windows.
               
- Only work for Word. I am planning to implement Excel and PowerPoint in the following days.


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

#*Creating the WordDocument object*

document = WordDocument('Example\\Path\\To\\file.docx')

#*Exporting to PDF*

document.toPdf('Example\\Export\\Path','OutputFileName')
