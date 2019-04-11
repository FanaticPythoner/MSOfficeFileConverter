# Autor : FanaticPythoner.
# Please read the "LICENSE" file before doing anything.

from winreg import *
import itertools
from contextlib import contextmanager
from win32api import GetFileVersionInfo, LOWORD, HIWORD
import os
import win32com.client
import sys


@contextmanager
def suppress(*exceptions):
    try:
        yield
    except exceptions:
        pass

allSupportedMSProgram = ['Excel', 'PowerPoint', 'Word']
allSupportedMSProgramExe = ['excel.exe','powerpnt.exe','winword.exe']

def _createRegKeys():
    """
    Create DWORD values to registry allowing the module\nto create macros into Microsoft Office documents.\n\n*Do not use it, there is an underscore for a reason.
    """
    def subkeys(path, hkey=HKEY_LOCAL_MACHINE, flags=0):
        """
        Get all subkeys of a registry key as string.
        """
        with suppress(WindowsError), OpenKey(hkey, path, 0, KEY_READ|flags) as k:
            for i in itertools.count():
                yield EnumKey(k, i)

    def get_version_number(filename):
        """Get the version of a file"""
        try:
            info = GetFileVersionInfo (filename, "\\")
            ms = info['FileVersionMS']
            ls = info['FileVersionLS']
            return HIWORD (ms), LOWORD (ms), HIWORD (ls), LOWORD (ls)
        except:
            return 0,0,0,0

    defaultPath = 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths'
    subKeys = subkeys(defaultPath)
    for key in subKeys:
        if key.lower() in allSupportedMSProgramExe:
            subKey = OpenKey(HKEY_LOCAL_MACHINE, defaultPath + '\\' + key, 0, KEY_READ)
            filePath = QueryValueEx(subKey, 'Path')
            version = get_version_number(filePath[0] + '\\' + key)
            version = str(version[0]) + '.' + str(version[1])
            writePath = "Software\\Microsoft\\Office\\" + version
            officeProductName = allSupportedMSProgram[allSupportedMSProgramExe.index(key.lower())]
            writeKey = OpenKey(HKEY_CURRENT_USER, writePath + '\\' + officeProductName + '\\Security', 0, KEY_ALL_ACCESS)
            SetValueEx(writeKey, "AccessVBOM", 0, REG_DWORD, 1)
            SetValueEx(writeKey, "VBAWarnings", 0, REG_DWORD, 1)


class WordDocument:
    """Open a Word document from a specified file path, then offer methods to convert it to whatever format you want.

Usage:
    #Creating the WordDocument object
    document = WordDocument('Example\\Path\\To\\file.docx')
    #Exporting to PDF
    document.toPdf('Example\\Export\\Path','ExampleFileName')

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
    - Odt"""
    def __init__(self, documentPath):
        if not os.path.isfile(documentPath):
            raise Exception('The specified file path does not exist.')
        _createRegKeys()
        self.documentPath = documentPath
        splittedPath = documentPath.split('\\')
        self.fileName = '.'.join(splittedPath[-1].split('.')[:-1])
        self.defaultExportPath = '\\'.join(splittedPath[:-1])

    def _export(self, exportFilePath, enumNum):
        """Internal magic function.\n\n*Do not use it, there is an underscore for a reason."""
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        document = word.Documents.Open(self.documentPath)
        document.SaveAs(exportFilePath, enumNum)
        word.Documents(1).Close(SaveChanges=False)
        word.Application.Quit()
        del word

    def _validateArgs(self, exportFolder, exportFileName):
        """Validate the args of an export function.\n\n*Do not use it, there is an underscore for a reason."""
        if exportFolder is None:
            exportFolder = self.defaultExportPath
        elif not os.path.isdir(exportFolder):
            raise Exception('The specified output directory does not exist.')

        exportFolder = os.path.abspath(exportFolder)

        if exportFileName is None:
            exportFileName = self.fileName

        fileExtension = '.' + sys._getframe(1).f_code.co_name[2:].split('_')[0].lower()

        if not exportFileName.endswith(fileExtension):
            exportFileName = exportFileName + fileExtension

        elif len(os.path.normpath(exportFolder + '\\' + exportFileName)) != len(exportFolder + '\\' + exportFileName):
            raise Exception('The specified file name or the specified export folder contain invalid characters.')
        return exportFolder, exportFileName

    def toDocx(self, exportFolder=None, exportFileName=None):
        """Export to Word Document.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath, 16)

    def toDocm(self, exportFolder=None, exportFileName=None):
        """Export to Word Macro-Enabled Document.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,13)

    def toDoc(self, exportFolder=None, exportFileName=None):
        """Export to Word 1997-2003 Document.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,0)

    def toDotm(self, exportFolder=None, exportFileName=None):
        """Export to Word Macro-Enabled Template.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,15)

    def toDot(self, exportFolder=None, exportFileName=None):
        """Export to Word 1997-2003 Template.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,1)

    def toPdf(self, exportFolder=None, exportFileName=None):
        """Export to PDF.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,17)

    def toXps(self, exportFolder=None, exportFileName=None):
        """Export to XPS Document.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,18)

    def toMht(self, exportFolder=None, exportFileName=None):
        """Export to Single File Web Page (*.mht).
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,9)

    def toMhtml(self, exportFolder=None, exportFileName=None):
        """Export to Single File Web Page (*.mhtml).
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,9)

    def toHtml(self, exportFolder=None, exportFileName=None):
        """Export to Web Page (*.html).
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\HtmlFiles_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,8)

    def toHtm(self, exportFolder=None, exportFileName=None):
        """Export to Web Page (*.htm).
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\HtmFiles_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,8)

    def toHtml_Filtered(self, exportFolder=None, exportFileName=None):
        """Export to Web Page, Filtered (*.htm; *.html).
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\HtmlFilteredFiles_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,10)

    def toRtf(self, exportFolder=None, exportFileName=None):
        """Export to Rich Text Format.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,6)

    def toTxt(self, exportFolder=None, exportFileName=None):
        """Export to Plain Text.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,2)

    def toXml(self, exportFolder=None, exportFileName=None):
        """Export to Word XML Document.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,12)

    def toXml_MacroEnabled(self, exportFolder=None, exportFileName=None):
        """Export to Word XML Document with macro enabled.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,13)

    def toXml_2003(self, exportFolder=None, exportFileName=None):
        """Export to Word 2003 XML Document.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,11)

    def toDocx_ReadOnly(self, exportFolder=None, exportFileName=None):
        """Export to Strict Open XML Document.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,24)

    def toOdt(self, exportFolder=None, exportFileName=None):
        """Export to OpenDocument Text.
        - If you do not specify an export folder, the document will be created in the same directory as the original Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._export(finalExportPath,23)


class ExcelDocument:
    """Open a Excel document from a specified file path, then offer methods to convert it to whatever format you want.

Usage:
    #Creating the ExcelDocument object
    document = ExcelDocument('Example\\Path\\To\\file.xlsx')
    #Exporting to PDF
    document.toPdf('Example\\Export\\Path','ExampleFileName')

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
    """

    def __init__(self, documentPath):
        if not os.path.isfile(documentPath):
            raise Exception('The specified file path does not exist.')
        _createRegKeys()
        self.documentPath = documentPath
        splittedPath = documentPath.split('\\')
        self.fileName = '.'.join(splittedPath[-1].split('.')[:-1])
        self.defaultExportPath = '\\'.join(splittedPath[:-1])
        self.defaultDocumentExtension = splittedPath[-1].split('.')[-1]

    def _exportAll(self, exportFilePath, enumNum):
        """Internal magic function.\n\n*Do not use it, there is an underscore for a reason."""
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(self.documentPath)
        workbook.SaveAs(exportFilePath, enumNum)
        workbook.Close(SaveChanges=False)
        excel.Application.Quit()
        del excel

    def _exportFixedFormat(self, exportFilePath, enumNum):
        """Internal magic function.\n\n*Do not use it, there is an underscore for a reason."""
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(self.documentPath)
        workbook.ExportAsFixedFormat(enumNum, exportFilePath)
        workbook.Close(SaveChanges=False)
        excel.Application.Quit()
        del excel

    def _exportAllSheets(self, exportFilePath, enumNum):
        """Internal magic function.\n\n*Do not use it, there is an underscore for a reason."""
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(self.documentPath)

        splittedFilePath = exportFilePath.split('\\')
        fileNameSplitted= splittedFilePath[-1].split('.')
        folderName = '\\'.join(splittedFilePath[:-1])
        extension = fileNameSplitted[-1]
        fileName = '.'.join(fileNameSplitted[:-1])
        for index, item in enumerate(workbook.Worksheets):
            newExportFilePath = folderName + '\\' + str(index + 1) + '.' + extension
            item.SaveAs(newExportFilePath, enumNum)

        workbook.Close(SaveChanges=False)
        excel.Application.Quit()
        del excel

    def _validateArgs(self, exportFolder, exportFileName):
        """Validate the args of an export function.\n\n*Do not use it, there is an underscore for a reason."""
        if exportFolder is None:
            exportFolder = self.defaultExportPath
        elif not os.path.isdir(exportFolder):
            raise Exception('The specified output directory does not exist.')

        exportFolder = os.path.abspath(exportFolder)

        if exportFileName is None:
            exportFileName = self.fileName

        fileExtension = '.' + sys._getframe(1).f_code.co_name[2:].split('_')[0].lower()

        if not exportFileName.endswith(fileExtension):
            exportFileName = exportFileName + fileExtension

        elif len(os.path.normpath(exportFolder + '\\' + exportFileName)) != len(exportFolder + '\\' + exportFileName):
            raise Exception('The specified file name or the specified export folder contain invalid characters.')
        return exportFolder, exportFileName

    def toXlsx(self, exportFolder=None, exportFileName=None):
        """Export to Excel Workbook.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 51)

    def toXlsm(self, exportFolder=None, exportFileName=None):
        """Export to Excel Macro-Enabled Workbook.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 52)

    def toXlsb(self, exportFolder=None, exportFileName=None):
        """Export to Excel Binary Workbook.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 50)

    def toXls(self, exportFolder=None, exportFileName=None):
        """Export to Excel 1997-2003 Workbook.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 56)

    def toCsv_UTF8(self, exportFolder=None, exportFileName=None):
        """Export to CSV UTF-8 (Comma delimited).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\CSV_UTF8_Files_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAllSheets(finalExportPath, 62)

    def toXml(self, exportFolder=None, exportFileName=None):
        """Export to XML Data.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 46)

    def toMht(self, exportFolder=None, exportFileName=None):
        """Export to Single File Web Page (*.mht).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 45)

    def toMhtml(self, exportFolder=None, exportFileName=None):
        """Export to Single File Web Page (*.html).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 45)

    def toXltm(self, exportFolder=None, exportFileName=None):
        """Export to Excel Macro-Enabled Template.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 53)

    def toXlt(self, exportFolder=None, exportFileName=None):
        """Export to Excel 1997-2003 Template.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 17)

    def toTxt_Windows(self, exportFolder=None, exportFileName=None):
        """Export to Text (Windows).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\TXT_Windows_Files_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAllSheets(finalExportPath, 20)

    def toTxt_Unicode(self, exportFolder=None, exportFileName=None):
        """Export to Unicode Text.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\TXT_Unicode_Files_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAllSheets(finalExportPath, 42)

    def toXls_95Workbook(self, exportFolder=None, exportFileName=None):
        """Export to Excel 5.0/95 Workbook.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 39)

    def toCsv(self, exportFolder=None, exportFileName=None):
        """Export to CSV (Comma delimited).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\CSV_Files_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAllSheets(finalExportPath, 6)

    def toCsv_Windows(self, exportFolder=None, exportFileName=None):
        """Export to CSV (Windows).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\CSV_Windows_Files_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAllSheets(finalExportPath, 23)

    def toPrn(self, exportFolder=None, exportFileName=None):
        """Export to Formatted Text (Space delimited).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\PRN_Files_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAllSheets(finalExportPath, 36)

    def toTxt_Macintosh(self, exportFolder=None, exportFileName=None):
        """Export to Text (Macintosh).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\TXT_Macintosh_Files_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAllSheets(finalExportPath, 19)

    def toTxt_MSDOS(self, exportFolder=None, exportFileName=None):
        """Export to Text (MS-DOS).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\TXT_MSDOS_Files_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAllSheets(finalExportPath, 21)

    def toCsv_Macintosh(self, exportFolder=None, exportFileName=None):
        """Export to CSV (Macintosh).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\CSV_Macintosh_Files_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAllSheets(finalExportPath, 22)

    def toCsv_MSDOS(self, exportFolder=None, exportFileName=None):
        """Export to CSV (MS-DOS).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\CSV_MSDOS_Files_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAllSheets(finalExportPath, 24)

    def toDif(self, exportFolder=None, exportFileName=None):
        """Export to DIF (Data Interchange Format).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 9)

    def toSlk(self, exportFolder=None, exportFileName=None):
        """Export to SYLK (Symbolic Link).
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\SLK_Files_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAllSheets(finalExportPath, 2)

    def toXlam(self, exportFolder=None, exportFileName=None):
        """Export to Excel Add-in.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 55)

    def toXla(self, exportFolder=None, exportFileName=None):
        """Export to Excel 1997-2003 Add-in.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 18)

    def toPdf(self, exportFolder=None, exportFileName=None):
        """Export to PDF.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportFixedFormat(finalExportPath, 0)

    def toXps(self, exportFolder=None, exportFileName=None):
        """Export to XPS Document.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 1)

    def toXlsx_ReadOnly(self, exportFolder=None, exportFileName=None):
        """Export to Scrict Open XML Spreadsheet.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 61)
        
    def toOds(self, exportFolder=None, exportFileName=None):
        """Export to OpenDocument Spreadsheet.
        - If you do not specify an export folder, the document will be created in the same directory as the Excel directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        self._exportAll(finalExportPath, 60)

#TODO
#class PowerPointDocument:
#    def __init__(self, documentPath):
#        _createRegKeys()
#        self.documentPath = documentPath
