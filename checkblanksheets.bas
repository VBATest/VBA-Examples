Attribute VB_Name = "checkblanksheets"
Sub checkblank()
For Each cell In Range("A1:f47")
If cell.Value = "" Then
cell.Value = "Blank"
Else
End If
Next cell
End Sub
