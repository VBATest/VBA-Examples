Attribute VB_Name = "SheetEmpty"
Sub Emptysheet()
Dim answer As Integer
answer = MsgBox("Are you sure you want to empty the sheet?", vbYesNo + vbQuestion, "Empty Sheet")
If answer = vbYes Then
Cells.ClearContents
Else
'do nothing
End If

End Sub


