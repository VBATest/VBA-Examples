Attribute VB_Name = "Addservice"
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
