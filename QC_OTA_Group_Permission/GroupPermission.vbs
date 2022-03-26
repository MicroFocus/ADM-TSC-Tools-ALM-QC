' GroupPermission_Excel.vbs
' Retrieve group permission
' arg 0: QC URL
' arg 1: QC user
' arg 2: QC password
' arg 3: path

 
' call it like
' cscript GroupPermission.vbs https://<almhost>/qcbin/ <user> <password> <path>
'Example:
'c:\windows\syswow64\cscript GroupPermission.vbs https://almserver.saas.microfocus.com/qcbin user password c:\temp

'////////////////////////////////////////////////////////////////////////////////////
'Global Variable
Dim Row
Row = 0
'////////////////////////////////////////////////////////////////////////////////////


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
Sub GetPermission(tdconn, Sheet)
On Error resume next

'ALM 
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

	Sheet.Cells(Row, 1).Value = custu.Name 
	Sheet.Cells(Row, 2).Value = custu.FullName               		
	Sheet.Cells(Row, 3).Value = custu.Email
	Sheet.Cells(Row, 4).Value = custu.Phone
	Sheet.Cells(Row, 5).Value = tdconn.domainname
	Sheet.Cells(Row, 6).Value = tdconn.projectname
	Sheet.Cells(Row, 7).Value = groupList
	Row = Row + 1
 Next

Set cust = Nothing
Set custus = Nothing
Set custu = Nothing
Set custgroups = Nothing



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

'Excel
Dim Excel, Sheet


Set args = WScript.Arguments
QcURL = args.Item(0)
QcUser = args.Item(1)
QcPassword = args.Item(2)
path = args.Item(3)
filename = "\GroupPermission_" & getTimeStamp & ".xlsx"
Set tdConnection = CreateObject("TDApiOle80.TDConnection.1")
If tdConnection Is Nothing Then
	Output "Could not create TD Connection object"
Else



	Set Excel = CreateObject("Excel.Application") 'Open Excel
	Excel.WorkBooks.Add() 'Add a new workbook
	
	'Get the first worksheet.
	Set Sheet = Excel.ActiveSheet

	'sheet name as Group Permission
	Sheet.Name = "Group Permission"

	'Specify the Excel Sheet Properties  
	With Sheet.Range("A1:H1")
	.Font.Name = "Arial"
	.Font.FontStyle = "Bold"
	.Font.Size = 10
	.Font.Bold = True
	.Interior.ColorIndex = 15 'Light Grey
	End With

	'Excel Sheet Column Header Values
	Sheet.Cells(1, 1) = "User ID"
    	Sheet.Cells(1, 2) = "Full Name"
    	Sheet.Cells(1, 3) = "Email"
    	Sheet.Cells(1, 4) = "Phone"
    	Sheet.Cells(1, 5) = "Domain"
    	Sheet.Cells(1, 6) = "Project"
    	Sheet.Cells(1, 7) = "Role(s)"
    
	tdConnection.InitConnectionEx QcURL
	AttachmentNameList = ""
	If tdConnection.Connected = False Then
		Output "Could not initialize QC connection"
	Else
		tdConnection.Login QCUser, QcPassword
		
		If tdConnection.LoggedIn = False Then
			Output "Could not log into ALM."
		Else

			'Starts saving info from the 2nd Row
			Row = 2
			For Each domainItem In tdConnection.VisibleDomains
			    For Each projn In tdConnection.VisibleProjects(domainItem)
			
				tdConnection.Connect domainItem, projn
		
				If tdConnection.ProjectConnected = False Then
					Output "Invalid domain/project"
				Else
					Output "Connecting to project " & domainItem & "." & projn
					GetPermission tdConnection,Sheet
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

'Save the newly created workbook and close Excel.
Excel.ActiveWorkbook.SaveAs(path & filename)
Excel.Quit