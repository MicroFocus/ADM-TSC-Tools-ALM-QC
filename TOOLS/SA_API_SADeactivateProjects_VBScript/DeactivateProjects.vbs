'Usage
'c:\windows\syswow64\cscript DeactivateProject.vbs <serverurl> <username> <userpassword> <workingdirectory> <projectlistfile>
'Example:
'c:\windows\syswow64\cscript DeactivateProject.vbs https://<server>/qcbin myname ##### c:\temp ProjectList.txt


'////////////////////////////////////////////////////////////////////////////////////
'//
'// ---'Write to log file---
'///////////////////////////////////////////////////////////////////////////////////
Sub updateMessage(path,filename, m1)
	' NewTextEC.vbs
	' Sample VBScript to write to a file. With added error-correcting
	' Author Guy Thomas http://computerperformance.co.uk/
	 ' VBScript Write File
	' ---------------------------------------------' 
	wscript.echo m1
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
'Purpose: Deactivate Projects
'//////////////////////////////////////////////////////////////
Private Sub DeactivateProject(sa,SourceDomain, ProjectName,QCWorkdir)
On Error resume next
	Dim response

        response = sa.DeactivateProject(SourceDomain, ProjectName)
        If Err.Number <> 0 Then
            updateMessage QCWorkdir,"DeactivateLog.txt","Failed to deactivate project " & SourceDomain & "." & ProjectName & "." & Err.Description
            Exit Sub
        Else
            updateMessage QCWorkdir,"DeactivateLog.txt","Project " & SourceDomain & "." & ProjectName & " has been deactivated."
        End If
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
Dim ProjArray
Dim t
Dim projectname
Dim domainname
Dim strLine
Dim fs
Dim afs

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

'Read the list of projects to deactivate
Set fs = CreateObject("Scripting.FileSystemObject")
 
if FExists(QCWorkDir & "\" & QCProjectFile) then
	set afs = fs.OpenTextFile(QCWorkDir & "\" & QCProjectFile,1)
	If Err.Number <> 0 Then
		Wscript.Echo "Failed to read project list." & Err.Description
		Main = 1 
    Exit function
End If
else
	Wscript.Echo "Project list does not exist." 
	Main = 1
    Exit function
end if


updateMessage QCWorkDir ,"DeactivateLog.txt", "Preparing to deactivate projects."
do while not afs.AtEndOfStream
	 domainname = ""
	 projectname = ""
	 strLine = ""
     strLine = afs.ReadLine()
	 ProjArray = Split(strLine, ".")
	 domainname =  Trim(ProjArray(0))
	 projectname = Trim(ProjArray(1))
	 if not (domainname = "" or  projectname = "") then
		updatemessage QCWorkdir,"DeactivateLog.txt", "Deactivating " & domainname & "." & projectname
		Call DeactivateProject(sa, domainname, projectname,QCWorkdir)
	 end if

loop
afs.Close
sa.logout
Main = 0
Set fs = Nothing
set afs = Nothing
Set sa = Nothing
End Function

'////////////////////////////////////////////////////////////////////
'Start main
call Main
'////////////////////////////////////////////////////////////////////