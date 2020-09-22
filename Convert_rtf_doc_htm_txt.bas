Attribute VB_Name = "rtf_doc_htm_txt_Convert"
Option Explicit

'Convert between .rtf .doc .htm .html .txt
''''''''''''''''''''''''''''''''''''''''''

'NOTE: Add a reference to 'Microsoft Word x.x Object Library'

'NOTE: Do not attempt to convert a passworded .doc or convert to an existing passworded .doc
'      You can only safely convert to a new passworded .doc

'Conversion formats possible with Office 2000:
'---------------------------------------------
'wdFormatDocument
'wdFormatDOSText
'wdFormatDOSTextLineBreaks
'wdFormatEncodedText
'wdFormatHTML
'wdFormatRTF
'wdFormatTemplate
'wdFormatText
'wdFormatTextLineBreaks
'wdFormatUnicodeText

'Other conversion formats possible with Office 2007 with free MS plugin:
'-----------------------------------------------------------------------
'wdFormatPDF
'etc

Public Sub ConvertFile(infile As String, outfile As String, Optional sPassword As String = "")
    Static WordObj As Word.Application
    Set WordObj = CreateObject("Word.Application")
    On Error GoTo err
    WordObj.Documents.Open (infile)
    WordObj.Visible = False 'Disable viewing the Word session and its document
    Select Case UCase(Right(outfile, 4))
        Case ".RTF"
            WordObj.ActiveDocument.SaveAs outfile, wdFormatRTF
        Case ".DOC"
            WordObj.ActiveDocument.SaveAs outfile, wdFormatDocument, , sPassword
        Case ".HTM"
            WordObj.ActiveDocument.SaveAs outfile, wdFormatHTML
        Case "HTML"
            WordObj.ActiveDocument.SaveAs outfile, wdFormatHTML
        Case ".TXT"
            WordObj.ActiveDocument.SaveAs outfile, wdFormatText
        Case Else
            MsgBox "Unrecognised Output File type!" + vbLf + vbTab + UCase(Right(outfile, 4)), vbCritical + vbOKOnly, "File Error"
    End Select
err:
    WordObj.Quit savechanges:=False 'close Word
    Set WordObj = Nothing
End Sub
'
'Usage Examples:
'
'Private Sub Command1_Click()
'    ConvertFile "c:\test.rtf", "c:\test.htm"
'End Sub
'
'Private Sub Command1_Click()
'    ConvertFile "c:\test.rtf", "c:\test.doc", "MyPassword" 'Can only set optional password for .doc, if used on other types it will be ignored
'End Sub
