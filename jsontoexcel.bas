Attribute VB_Name = "jsontoexcel"
Sub getJSON()
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", "http://lexevscts2.nci.nih.gov/lexevscts2/codesystemversions?format=json"
    MyRequest.send
    ' MsgBox MyRequest.ResponseText

Dim Json As Object
Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)

MsgBox Json("CodeSystemVersionCatalogEntryDirectory")("entry")(1)("codeSystemVersionName")

End Sub
