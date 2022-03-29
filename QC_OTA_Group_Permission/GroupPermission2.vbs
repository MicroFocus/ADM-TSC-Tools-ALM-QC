' GroupPermission2.vbs
' retrieve user group permission and export to csv
' arg 0: QC URL
' arg 1: QC user
' arg 2: QC password
' arg 3: path

 
' call it like
' cscript GroupPermission2.vbs https://<almhost>/qcbin/ <user> <password> <path>
'Example:
'c:\windows\syswow64\cscript GroupPermission2.vbs https://almserver.saas.microfocus.com/qcbin user password c:\temp


'////////////////////////////////////////////////////////////////////////////////////
'//
'// ---'Output to console---
'///////////////////////////////////////////////////////////////////////////////////
Public Sub Output (Content)
	WScript.StdOut.WriteLine Content
End Sub

'////////////////////////////////////////////////////////////////////////////////////
'//
'// ---'Get Group Permission--
'///////////////////////////////////////////////////////////////////////////////////
Sub GetPermission(tdconn, path,filename)
On Error resume next
Dim cust 
Dim custus 
Dim custu 
Dim myuser 
Dim groupname 
Dim custgroups 
Dim groupid
Dim groupList

err.clear
Set cust = tdconn.Customization
if err.number <> 0 then
   exit sub
end if

cust.Load
Set custus = cust.Users
Set custgroups = cust.UsersGroups

For Each custu In custus.Users
    groupList = ""
    For Each groupname In custgroups.Groups
        'Determine which groups the user is in
        If custu.InGroup(groupname.Name) Then
            If groupList = "" Then
                groupList = groupname.Name
            Else
                groupList = groupname.Name & ";" & groupList
            End If
        End If
    Next

   m_out path, filename, """" & custu.Name &  ""","""  & custu.FullName &  ""","""   & custu.Email &  ""","""  & custu.Phone &  ""","""   & tdconn.domainname &  ""","""   & tdconn.projectname &  ""","""   & groupList & """"

 Next

Set cust = Nothing
Set custus = Nothing
Set custu = Nothing
Set custgroups = Nothing

End Sub

'////////////////////////////////////////////////////////////////////////////////////
'//
'// ---'Write to log file---
'///////////////////////////////////////////////////////////////////////////////////
Sub M_out(path ,filename, m1)
	' NewTextEC.vbs
	' Sample VBScript to write to a file. With added error-correcting
	' Author Guy Thomas http://computerperformance.co.uk/
	 ' VBScript Write File
	' ---------------------------------------------' 

	Dim objFSO, objFolder, objShell, objTextFile, objFile
	Dim strDirectory, strFile, strText
	strDirectory = path
	strFile = filename
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



'////////////////////////////////////////////////////////////////////////////////////
'//
'// ---'Timestamp for log
'// 
'///////////////////////////////////////////////////////////////////////////////////
Function LZ(ByVal Number)
  If Number < 10 Then
    LZ = "0" & CStr(Number)
  Else
    LZ = CStr(Number)
  End If
End Function

Function getTimeStamp
  Dim CurrTime
  CurrTime = Now()

  getTimeStamp = CStr(Year(CurrTime)) & "-" _
    & LZ(Month(CurrTime)) & "-" _
    & LZ(Day(CurrTime)) & " " _
    & LZ(Hour(CurrTime)) & "_" _
    & LZ(Minute(CurrTime)) & "_" _
    & LZ(Second(CurrTime))
End Function




'////////////////////////////////////////////////////////////////////////////////////
'//
'// --- Main ---
'///////////////////////////////////////////////////////////////////////////////////
'On Error Resume Next
Dim args
Dim tdConnection
Dim QcURL
Dim QcUser
Dim QcPassword
Dim QcDomain
Dim QcProject
Dim i
Dim path
Dim domainItem
Dim projn
Dim filename

Set args = WScript.Arguments
QcURL = args.Item(0)
QcUser = args.Item(1)
QcPassword = args.Item(2)
path = args.Item(3)
filename = "\GroupPermission_" & getTimeStamp & ".csv"
Set tdConnection = CreateObject("TDApiOle80.TDConnection.1")
If tdConnection Is Nothing Then
	Output "Could not create TD Connection object"
Else
	tdConnection.InitConnectionEx QcURL
	AttachmentNameList = ""
	If tdConnection.Connected = False Then
		Output "Could not initialize QC connection"
	Else
		tdConnection.Login QCUser, QcPassword
		
		If tdConnection.LoggedIn = False Then
			Output "Could not log into ALM."
		Else
			m_out path,filename, "User Id" &  ","  & "Full Name" &  ","  & "Email" &  ","  & "Phone" &  ","  & "Domain" &  ","  & "Project" &  ","  & "Role(s)"

			For Each domainItem In tdConnection.VisibleDomains
			    For Each projn In tdConnection.VisibleProjects(domainItem)
			
				tdConnection.Connect domainItem, projn
		
				If tdConnection.ProjectConnected = False Then
					Output "Invalid domain/project"
				Else
					Output "Connecting to project " & domainItem & "." & projn
					GetPermission tdConnection,path,filename
				 	tdConnection.Disconnect
				End If	
			    Next
  		         Next

		End If
		tdConnection.Logout
		tdConnection.ReleaseConnection
		output "ALM Disconnected."
		output "Result is written to " & path & filename
	End If
End If	

set tdConnection = nothing
