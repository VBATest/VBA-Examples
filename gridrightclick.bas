Attribute VB_Name = "gridrightclick"
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
RightClick
Cancel = True

End Sub

Sub Grids()
ActiveWindow.DisplayGridlines = Not ActiveWindow.DisplayGridlines
End Sub
Sub Formulas()
ActiveWindow.DisplayFormulas = Not ActiveWindow.DisplayFormulas
End Sub
Sub Preview()
ActivSheet.PrintPreview
End Sub

Sub RightClick()
Dim vArr As Variant, i As Integer
Dim OMenu As CommandBar, Item As CommandBarControl
Set OMenu = CommandBars.Add("", msoBarPopup, , True)

vArr = Array("Grids", "Formulas", "Preview")
For i = 0 To UBound(vArr)
Set OItem = OMenu.Controls.Add
OItem.Caption = vArr(i)
OItem.OnAction = vArr(i)

Next i
OMenu.ShowPopup

End Sub
