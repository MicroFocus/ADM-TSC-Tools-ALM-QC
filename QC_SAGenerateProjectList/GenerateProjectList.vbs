'Usage
'c:\windows\syswow64\cscript GenerateProjectList.vbs <serverurl> <username> <userpassword> <workingdirectory> <projectlistfile>
'Example:
'c:\windows\syswow64\cscript GenerateProjectList.vbs https://<server>/qcbin myname ##### c:\temp ProjectList.txt


'////////////////////////////////////////////////////////////////////////////////////
'//Purpose: Write to file
'///////////////////////////////////////////////////////////////////////////////////
Sub Export(path ,filename, m1)
	' NewTextEC.vbs
	' Sample VBScript to write to a file. With added error-correcting
	' Author Guy Thomas http://computerperformance.co.uk/
	 ' VBScript Write File
	' ---------------------------------------------' 
	'wscript.echo m1
	Dim objFSO, objFolder, objShell, objTextFile, objFile
	Dim strDirectory, strFile, strText
	strDirectory = path
	strFile = "\" & filename
	strText = m1

	' Create the File System Object
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	' Check that the strDirectory folder exists
	If objFSO.FolderExists(strDirectory) Then
	   Set objFolder = objFSO.GetFolder(strDirectory)
	Else
	   Set objFolder = objFSO.CreateFolder(strDirectory)
	   WScript.Echo "Just created " & strDirectory
	End If

	If objFSO.FileExists(strDirectory & strFile) Then
	   Set objFolder = objFSO.GetFolder(strDirectory)
	Else
	   Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
	   Wscript.Echo "Just created " & strDirectory & strFile
	End If 

	set objFile = nothing
	set objFolder = nothing
	' OpenTextFile Method needs a Const value
	' ForAppending = 8 ForReading = 1, ForWriting = 2
	Const ForAppending = 8

	Set objTextFile = objFSO.OpenTextFile _
	(strDirectory & strFile, ForAppending, True)

	' Writes strText every time you run this VBScript
	objTextFile.WriteLine(strText)
	objTextFile.Close
	' End of VBScript to write to a file with error-correcting Code
End Sub

'//////////////////////////////////////////////////////////////
'Purpose: This function checks if required folde
'//////////////////////////////////////////////////////////////
Public Sub FoldExists(Folder)
Dim MyFileSystem

Set MyFileSystem = CreateObject("Scripting.FileSystemObject")

    If MyFileSystem.FolderExists(Folder) = False Then
       MyFileSystem.CreateFolder (Folder)
    End If
Set MyFileSystem = Nothing
End Sub

'//////////////////////////////////////////////////////////////
'Purpose: This function check to see if the file exists
'//////////////////////////////////////////////////////////////
Public Function FExists(OrigFile)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
FExists = fs.fileexists(OrigFile)
End Function

'//////////////////////////////////////////////////////////////
'Purpose: Parse xml to retrieve data defined in query
'//////////////////////////////////////////////////////////////
Function parseXML(xmlstring,xmlquery)
On Error Resume Next

	'Instantiate the XML parser.
	Dim Nodelist ' As IXMLDOMNodeList
	Dim MyXMLDom

	Set MyXMLDom = CreateObject("Microsoft.XMLDOM")
	If Err.number <> 0 Then
		wscript.echo  "Failed to load Microsoft.XMLDOM. Object not found."
	End If
	'Load the xml file
    MyXMLDom.Loadxml(xmlstring)
	If Err.number  <> 0 Then
		wscript.echo "Faild to load xml string."
	End If

	Set Nodelist = MyXMLDom.selectNodes(xmlquery)
	If Err.number <> 0 Then
		wscript.echo "Path:" & xmlquery & " is incorrect. Node not found."
	End If

	'wscript.echo "Count:" & Nodelist.length
	set parseXML = Nodelist	
	
	set Nodelist = nothing
	set MyXMLDom = nothing
End Function


'//////////////////////////////////////////////////////////////
'Purpose: Get list of projects
'//////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////
Private Sub generateProjectList(sa,QCWorkdir,QCProjectFile)
On Error resume next
	Dim response
	Dim Nodelist
	Dim DomainNodeList
	Dim ProjectNodeList
	
    response = sa.GetAllDomains
	set Nodelist = parseXML(response, "/GetAllDomains/TDXItem")
	

    If Err.Number <> 0 Then
            wscript.echo "Failed to get project list." & Err.Description
            Exit Sub
    Else
        wscript.echo "Writing project list to " & QCWorkdir & "\" & QCProjectFile & "."
		For t = 0 To Nodelist.length - 1
			Set DomainNodeList = Nodelist.Item(t).selectNodes("DOMAIN_NAME")
			Set ProjectNodeList = Nodelist.Item(t).selectNodes("PROJECTS_LIST/TDXItem/PROJECT_NAME")
			'Set ProjectNodeStatusList = Nodelist.Item(t).selectNodes("PROJECTS_LIST/TDXItem/PR_IS_ACTIVE")
			For x = 0 To DomainNodeList.length - 1
				For y = 0 To ProjectNodeList.length - 1					
						'if ProjectNodeStatusList.Item(y).Text = "Y" then 'Enable if you want only active projects
							'domain.project
							Export QCWorkdir , QCProjectFile, DomainNodeList.Item(x).Text & "." & ProjectNodeList.Item(y).Text
						'end if
				Next
			Next
		next
    End If
	set NodeList = nothing
	set DomainNodeList = nothing
	set ProjectNodeList = nothing
End Sub


'/////////////////////////////////////////////////////////////////
'Main
'/////////////////////////////////////////////////////////////////
Private Function Main
Dim args
Dim QcURL
Dim QcUser
Dim QcPassword
Dim QCWorkdir
Dim QCProjectFile
Dim sa

Set args = WScript.Arguments
QcURL = args.Item(0)
QcUser = args.Item(1)
QcPassword = args.Item(2)
QCWorkdir = args.Item(3)
QCProjectFile = args.Item(4)

'Check for working directory
FoldExists (QCWorkdir)

'Connect to SA
set sa = CreateObject("SACLIENT.SaApi.9")

sa.login QcUrl, QCUser, QcPassword
If Err.Number <> 0 Then
    Wscript.Echo "Fail to connect to Site Admin." & Err.Description
	Main = 1 
    Exit function
else	
	Wscript.echo "Connected to Site Admin: "  & qcurl
End If

call generateProjectList(sa,QCWorkdir,QCProjectFile)
sa.logout
wsript.echo  "Done!"
Main = 0
Set fs = Nothing
set afs = Nothing
Set sa = Nothing
End Function

'////////////////////////////////////////////////////////////////////
'Start main
call Main
'////////////////////////////////////////////////////////////////////