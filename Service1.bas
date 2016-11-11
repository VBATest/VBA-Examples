Attribute VB_Name = "Service1"
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
