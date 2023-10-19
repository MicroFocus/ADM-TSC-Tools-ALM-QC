' RestGetTestAttachments.vbs
' Download attachment from a test using REST API
' arg 0: QC URL
' arg 1: QC apikey
' arg 2: QC secret password
' arg 3: Domain
' arg 4: Project
' arg 5: TestID
' arg 6: path

 
' call it like
'Usage:
'cscript RestGetTestAttachments.vbs https://serverurl/qcbin apikey secret domain project TestID c:\temp

'Example:
'cscript RestGetTestAttachments.vbs https://almserver.saas.microfocus.com/qcbin/qcbin oauth2-token secretpassword default qualitycenter_demo 61 c:\temp

Sub DownloadFile(restReq, oathtoken, url ,filename)

    restReq.open "GET", url, False
	restReq.setRequestHeader "Accept", "application/octet-stream"
    restReq.setRequestHeader "Cookie", oathtoken
    restReq.send()
    If restReq.status = "200" Then
         ' Create a binary stream to save the response content
		Set objStream = CreateObject("Adodb.Stream")
		objStream.Type = 1 ' Binary
		objStream.Open
		objStream.Write restReq.responseBody
		objStream.SaveToFile path & "\" & filename
		objStream.Close
		wscript.echo "Download file:" & path & "\" & filename
    else
        wscript.echo "Download request failed."
    End If
End Sub

'////////////////////////////////////////////////////////////////////////////////////
'//
'// --- Main ---
'///////////////////////////////////////////////////////////////////////////////////
'On Error Resume Next
Dim args
Dim QcURL
Dim apikey
Dim secret
Dim QcDomain
Dim QcProject
Dim testID
Dim path
Dim filename
Dim restReq
Dim oathtoken 
Dim reqHeader
Dim url
Dim ResponseText


Set args = WScript.Arguments
QcURL = args.Item(0)
apikey = args.Item(1)
secret = args.Item(2)
QcDomain = args.Item(3)
QcProject = args.Item(4)
testid = args.Item(5)
path = args.Item(6)

oathtoken = ""


set restReq = CreateObject("MSXML2.ServerXMLHTTP")

wscript.echo "QCURL: " & QcURL
wscript.echo "QcDomain: " &  QcDomain
wscript.echo "QcProject: " &  QcProject
wscript.echo "Test ID: " &  testid
wscript.echo "path: "  & path

url = cstr(QCURL & "/rest/oauth2/login")
reqHeader = CStr("{""clientId"":""" & apikey & """, ""secret"":""" & secret & """}")

restReq.open "POST", url, False
restReq.setRequestHeader "Content-Type", "application/json"
restReq.send reqHeader
If restReq.status = "200" Then
    oathtoken = restReq.getResponseHeader("Set-Cookie")
else
    wscript.echo  "Failed to login. Please try again. Return Status: " & restReq.status
End If

If oathtoken <> "" Then

    'Query to get list of attachments
    url = cstr(QcURL & "/rest/domains/" & QcDomain & "/projects/" & QcProject & "/tests/" & testid & "/attachments?fields=id,name")
    restReq.open "GET", url, False
	restReq.setRequestHeader "Content-Type", "application/xml"
    restReq.setRequestHeader "Cookie", oathtoken
    restReq.send()
	
    If restReq.status = "200" Then
		'Parse the 
	
		' Create XML document and load the XML string
		Dim xmlDoc 
		Dim entityNode 
		Dim entityNodes 
		Dim attachmentid 
		Dim attachmentFilename
		
		
		
		Set xmlDoc = CreateObject("MSXML2.DOMDocument")
		xmlDoc.async = False
		xmlDoc.LoadXML restReq.ResponseText
	   

		Set entityNodes = xmlDoc.SelectNodes("//Entity[@Type='attachment']")
		
		'Loop through each attachment in the xml file provided
		For Each entityNode In entityNodes
			attachmentid = entityNode.SelectSingleNode("Fields/Field[@Name='id']/Value").Text
			attachmentFilename = entityNode.SelectSingleNode("Fields/Field[@Name='name']/Value").Text
			DownloadFile restReq, oathtoken, QcURL & "/rest/domains/" & QcDomain & "/projects/" & QcProject & "/tests/" & testid & "/attachments/" & attachmentFilename, attachmentFilename
		Next
		
		Set xmlDoc = Nothing
		Set entityNodes = Nothing
		
        
    else
        wscript.echo "Failed to get list of attachments."
    End If
    
    'Logout
    url = CStr(QcURL & "/authentication-point/logout")
    restReq.open "GET", url, False
    restReq.send()
    If restReq.status = "200" Then
        Wscript.Echo "You have successfully logged out."
    else
        Wscript.Echo "Failed to logout: " & restReq.status
    End If
    
End If


set restReq = nothing
set args = nothing
