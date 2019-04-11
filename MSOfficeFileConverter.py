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
officeProgramAndVersionsFound = []

def _createRegKeys():
    """
    Create DWORD values to registry allowing the module\nto create macros into microsoft office documents.\n\n*Do not use it, there is an underscore for a reason.
    """
    def subkeys(path, hkey=HKEY_LOCAL_MACHINE, flags=0):
        """
        Get all subkeys of a registry key as string.
        """
        with suppress(WindowsError), OpenKey(hkey, path, 0, KEY_READ|flags) as k:
            for i in itertools.count():
                yield EnumKey(k, i)

    def get_version_number (filename):
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
            officeProgramAndVersionsFound.append((version, key))
            writePath = "Software\\Microsoft\\Office\\" + version
            officeProductName = allSupportedMSProgram[allSupportedMSProgramExe.index(key.lower())]
            writeKey = OpenKey(HKEY_CURRENT_USER, writePath + '\\' + officeProductName + '\\Security', 0, KEY_ALL_ACCESS)
            SetValueEx(writeKey, "AccessVBOM", 0, REG_DWORD, 1)
            SetValueEx(writeKey, "VBAWarnings", 0, REG_DWORD, 1)


class WordDocument:
    """Open as Word document from a specified file path, then offer methods to convert it to whatever format you want.

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

    def _export(self, macro, finalExportPath):
        """Internal magic function.\n\n*Do not use it, there is an underscore for a reason."""
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        document = word.Documents.Open(self.documentPath)
        documentModule = document.VBProject.VBComponents.Add(1)
        documentModule.CodeModule.AddFromString(macro)
        word.Application.Run('export')
        #word.Documents(1).Close(SaveChanges=False)
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
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatDocumentDefault',1)
        self._export(macro,finalExportPath)


    def toDocm(self, exportFolder=None, exportFileName=None):
        """Export to Word Macro-Enabled Document.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatXMLDocumentMacroEnabled',1)
        self._export(macro,finalExportPath)


    def toDoc(self, exportFolder=None, exportFileName=None):
        """Export to Word 1997-2003 Document.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatDocument',1)
        self._export(macro,finalExportPath)

    #Broken for some reason.
    #def toDotx(self, exportFolder=None, exportFileName=None):
    #    """Export to Word Template.
    #    - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
    #    - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
    #    exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
    #    finalExportPath = exportFolder + '\\' + exportFileName
    #    macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace(', _ \nFileFormat:= ','',1)
    #    self._export(macro,finalExportPath)


    def toDotm(self, exportFolder=None, exportFileName=None):
        """Export to Word Macro-Enabled Template.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatXMLTemplateMacroEnabled',1)
        self._export(macro,finalExportPath)


    def toDot(self, exportFolder=None, exportFileName=None):
        """Export to Word 1997-2003 Template.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatTemplate97',1)
        self._export(macro,finalExportPath)


    def toPdf(self, exportFolder=None, exportFileName=None):
        """Export to PDF.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\specialFormatMacro.txt','r').readlines()).replace('OutputFileName:= ""','OutputFileName:= "' + finalExportPath + '"',1).replace('ExportFormat:= ','ExportFormat:= wdExportFormatPDF',1)
        self._export(macro,finalExportPath)

    def toXps(self, exportFolder=None, exportFileName=None):
        """Export to XPS Document.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\specialFormatMacro.txt','r').readlines()).replace('OutputFileName:= ""','OutputFileName:= "' + finalExportPath + '"',1).replace('ExportFormat:= ','ExportFormat:= wdExportFormatXPS',1)
        self._export(macro,finalExportPath)


    def toMht(self, exportFolder=None, exportFileName=None):
        """Export to Single File Web Page (*.mht).
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatWebArchive',1)
        self._export(macro,finalExportPath)

    def toMhtml(self, exportFolder=None, exportFileName=None):
        """Export to Single File Web Page (*.mhtml).
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatWebArchive',1)
        self._export(macro,finalExportPath)

    def toHtml(self, exportFolder=None, exportFileName=None):
        """Export to Web Page (*.html).
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\HtmlFiles_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatHTML',1)
        self._export(macro,finalExportPath)

    def toHtm(self, exportFolder=None, exportFileName=None):
        """Export to Web Page (*.htm).
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\HtmFiles_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatHTML',1)
        self._export(macro,finalExportPath)

    def toHtml_Filtered(self, exportFolder=None, exportFileName=None):
        """Export to Web Page, Filtered (*.htm; *.html).
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        exportFolder = exportFolder + '\\HtmlFilteredFiles_' + exportFileName
        os.makedirs(exportFolder)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatFilteredHTML',1)
        self._export(macro,finalExportPath)

    def toRtf(self, exportFolder=None, exportFileName=None):
        """Export to Rich Text Format.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatRTF',1)
        self._export(macro,finalExportPath)

    def toTxt(self, exportFolder=None, exportFileName=None):
        """Export to Plain Text.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatText',1)
        self._export(macro,finalExportPath)

    def toXml(self, exportFolder=None, exportFileName=None):
        """Export to Word XML Document.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatXMLDocument',1)
        self._export(macro,finalExportPath)

    def toXml_MacroEnabled(self, exportFolder=None, exportFileName=None):
        """Export to Word XML Document with macro enabled.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatXMLDocumentMacroEnabled',1)
        self._export(macro,finalExportPath)

    def toXml_2003(self, exportFolder=None, exportFileName=None):
        """Export to Word 2003 XML Document.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatXML',1)
        self._export(macro,finalExportPath)


    def toDocx_ReadOnly(self, exportFolder=None, exportFileName=None):
        """Export to Strict Open XML Document.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatStrictOpenXMLDocument',1)
        self._export(macro,finalExportPath)


    def toOdt(self, exportFolder=None, exportFileName=None):
        """Export to OpenDocument Text.
        - If you do not specify an export folder, the document will be created in the same directory as the Word directory.
        - If you do not specify an export file name, the document will have the same name as the original document, only the extension will change."""
        exportFolder, exportFileName = self._validateArgs(exportFolder,exportFileName)
        finalExportPath = exportFolder + '\\' + exportFileName
        macro = ''.join(open('Macros\\Word\\defaultMacro.txt','r').readlines()).replace('FileName:=""','FileName:="' + finalExportPath + '"',1).replace('FileFormat:= ','FileFormat:= wdFormatOpenDocumentText',1)
        self._export(macro,finalExportPath)

#TODO
#class ExcelDocument:
#    def __init__(self, documentPath):
#        _createRegKeys()
#        self.documentPath = documentPath

#TODO
#class PowerPointDocument:
#    def __init__(self, documentPath):
#        _createRegKeys()
#        self.documentPath = documentPath
