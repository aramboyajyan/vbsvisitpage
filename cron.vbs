' Visit a user defined URL
' To be used as a replacement for cron tasks on Windows servers.
' 
' Created by: Topsitemakers
' http://www.topsitemakers.com/

Call RunIt()
Sub RunIt()

Dim RequestObj
Dim URL
Set RequestObj = CreateObject("Microsoft.XMLHTTP")

' Define the URL to be visited
URL = "http://example.com/cron.php"

' Open the request to our URL and send a request
RequestObj.open "POST", URL ,false
RequestObj.Send

' Done - cleanup
Set RequestObj = Nothing
End Sub