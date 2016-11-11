Attribute VB_Name = "excel2pdfselection"
Sub PrintSelectionToPDF()
'SUBROUTINE: PrintSelectionToPDF
'DEVELOPER: Vikas Sharma
'DESCRIPTION: Print your currently selected range to a PDF
Dim ThisRng As Range
Dim strfile As String
Dim myfile As Variant
If Selection.Count = 1 Then
Set ThisRng = Application.InputBox("Select a range", "Get Range", Type:=8)
Else
Set ThisRng = Selection
End If
'Prompt for save location
strfile = "Selection" & "_" _
& Format(Now(), "yyyymmdd_hhmmss") _
& ".pdf"
strfile = ThisWorkbook.Path & "\" & strfile
myfile = Application.GetSaveAsFilename _
(InitialFileName:=strfile, _
FileFilter:="PDF Files (*.pdf), *.pdf", _
Title:="Select Folder and File Name to Save as PDF")
If myfile <> "False" Then 'save as PDF
ThisRng.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
myfile, Quality:=xlQualityStandard, IncludeDocProperties:=True, _
IgnorePrintAreas:=False, OpenAfterPublish:=True
Else
MsgBox "No File Selected. PDF will not be saved", vbOKOnly, "No File Selected"
End If
End Sub



