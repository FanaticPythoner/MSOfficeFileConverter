# MSOfficeFileConverter
Allow you to convert/export a Microsoft Office file to a specified format, without any other dependency. You only need the Microsoft Office program you want to convert from (EX. You need Word if you want to convert a .docx to .pdf).

*Language:* Tested in Python 3.6, should work in all Python 3 version.

*Limitations:* - For now, the version in development only support Windows.
             - Only work for Word. I am planning in implementing Excel and PowerPoint in the following days.



# WordDocument Class

*Description :*
Open as Word document from a specified file path, then offer methods to convert it to whatever format you want.

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

 
 
*Usage :*

#Creating the WordDocument object

document = WordDocument('Example\\Path\\To\\file.docx')

#Exporting to PDF

document.toPdf('Example\\Export\\Path','ExampleFileName')
