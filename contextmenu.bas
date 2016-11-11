Attribute VB_Name = "Module1"
Sub AddToCellMenu()
    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl

    ' Delete the controls first to avoid duplicates.
    Call DeleteFromCellMenu

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Add one built-in button(Save = 3) to the Cell context menu.
    ContextMenu.Controls.Add Type:=msoControlButton, ID:=3, Before:=1

    ' Add one custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, Before:=2)
        .OnAction = "'" & ThisWorkbook.name & "'!" & "ToggleCaseMacro"
        .FaceId = 610
        .Caption = "Reconcile..."
        .Tag = "My_Cell_Control_Tag"
    End With


   
End Sub

Sub DeleteFromCellMenu()
    Dim ContextMenu As CommandBar
    Dim ctrl As CommandBarControl

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Delete the custom controls with the Tag : My_Cell_Control_Tag.
    For Each ctrl In ContextMenu.Controls
        If ctrl.Tag = "My_Cell_Control_Tag" Then
            ctrl.Delete
        End If
    Next ctrl

    ' Delete the custom built-in Save button.
    On Error Resume Next
    ContextMenu.FindControl(ID:=3).Delete
    On Error GoTo 0
End Sub


Sub ToggleCaseMacro()
    Dim Ret_type As Integer
Dim strMsg As String
Dim strTitle As String

' Dialog Message
strMsg = "Click any one of the below buttons."
' Dialog's Title
strTitle = "Reconcile1"
'Display MessageBox
    Ret_type = MsgBox(strMsg, vbYesNoCancel + vbQuestion, strTitle)
' Check pressed button
Select Case Ret_type
Case 6
    Range("A1:G50").Value = "Yes"
    
Case 7
    Range("A1:G50").Value = "No"
Case 2
    MsgBox "No Data"
End Select


End Sub


Sub ToggleCaseMacro1()
    Dim Ret_type As Integer
Dim strMsg As String
Dim strTitle As String
' Dialog Message
strMsg = "Click any one of the below buttons."
' Dialog's Title
strTitle = "Reconcile2"
'Display MessageBox
    Ret_type = MsgBox(strMsg, vbYesNoCancel + vbQuestion, strTitle)
    
' Check pressed button
Select Case Ret_type
Case 6
    MsgBox "You clicked 'YES' button."
Case 7
    MsgBox "You clicked 'NO' button."
Case 2
    MsgBox "You clicked 'CANCEL' button."
End Select
End Sub

Sub checkblank()
For Each cell In Range("A1:f47")
If cell.Value = "" Then
cell.Value = "Blank"
Else
End If
Next cell
End Sub

Sub Addservice()
Dim name As String
name = InputBox("Enter Service Name :", "Add Service")

Range("A1").Value = name

  
If Len(name) = 0 Then 'Checking if Length of name is 0 characters
  MsgBox "Please enter a valid name!", vbCritical
 

Else
  MsgBox "Hello " & name & " welcome to our Network."

  

End If
End Sub

