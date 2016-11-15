Attribute VB_Name = "importxmltoexcel"
Sub ImportXMLtoList()
Dim strTargetFile As String
Dim wb As Workbook

     Application.ScreenUpdating = False
     Application.DisplayAlerts = False
     strTargetFile = "C:\Users\M166363\Desktop\books.xml"
     Set wb = Workbooks.OpenXML(Filename:=strTargetFile, LoadOption:=xlXmlLoadImportToList)
     Application.DisplayAlerts = True

     wb.Sheets(1).UsedRange.Copy ThisWorkbook.Sheets("Sheet2").Range("A1")
     wb.Close False
     Application.ScreenUpdating = True
End Sub
