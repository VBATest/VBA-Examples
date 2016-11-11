Attribute VB_Name = "Sendmail"
Sub SendEmail()
'Dim objHTTP As New MSXML2.XMLHTTP
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
URL = "http://localhost:8888/rest/mail/send"
objHTTP.Open "POST", URL, False
objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
objHTTP.send ("var1=value1&var2=value2&var3=value3")

End Sub
