VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Rename Users"
   ClientHeight    =   10755
   ClientLeft      =   135
   ClientTop       =   420
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   10755
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox RenameUser_list 
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   8880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Server_textbox 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Text            =   "http://<servername:port>/qcbin/"
      Top             =   120
      Width           =   8895
   End
   Begin VB.Frame Customize 
      Caption         =   "Cutomize"
      Height          =   3975
      Left            =   240
      TabIndex        =   10
      Top             =   4680
      Width           =   10575
      Begin VB.CheckBox Check26 
         Caption         =   "Rename Users"
         Height          =   495
         Left            =   1080
         TabIndex        =   17
         ToolTipText     =   "Rename Users"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Excel_textbox 
         Height          =   495
         Left            =   1080
         TabIndex        =   12
         Text            =   "C:\temp\RenameUsers\UsersList.xls"
         Top             =   360
         Width           =   9255
      End
      Begin VB.TextBox Log_textbox 
         Height          =   405
         Left            =   1080
         TabIndex        =   14
         Text            =   "C:\temp\RenameUsers"
         Top             =   1080
         Width           =   9255
      End
      Begin VB.CommandButton Update_button 
         Caption         =   "Update"
         Height          =   495
         Left            =   1080
         TabIndex        =   15
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Excel_label 
         Caption         =   "Path:"
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Log_label 
         Caption         =   "Log:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.CommandButton Connect_button 
      Caption         =   "Connect"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame SelectTarget 
      Caption         =   "Select Target Projects"
      Height          =   3735
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   10575
      Begin VB.CheckBox AllProject_checkbox 
         Caption         =   "All Projects"
         Height          =   495
         Left            =   3120
         TabIndex        =   8
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton Project_button 
         Caption         =   "Select Project(s)"
         Height          =   495
         Left            =   4440
         TabIndex        =   9
         Top             =   3120
         Width           =   1455
      End
      Begin VB.ListBox Project_list 
         Height          =   2595
         Left            =   3120
         MultiSelect     =   1  'Simple
         TabIndex        =   7
         Top             =   360
         Width           =   7335
      End
      Begin VB.CheckBox AllDomain_checkbox 
         Caption         =   "All Domains"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ListBox Domain_list 
         Height          =   2595
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton Domain_button 
         Caption         =   "Select Domain(s)"
         Height          =   495
         Left            =   1440
         TabIndex        =   5
         Top             =   3120
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   9600
      Width           =   10695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public td As Object 'As New TDAPIOLELib.TDConnection
Dim isconnect As Boolean
'Start
Dim AllDomains As Boolean
Dim AllProjects As Boolean
Public cuser As String
Public cpassword As String
Public cdomain As String
Public cproject As String
Public cLogPath As String
Public cServerURL As String
Public comcl As Boolean
Public ReadOnlyGroup As String
'End
Dim cProjectsFile As String
Dim cExcelFile As String
Dim cCommandLine As Boolean
Dim filename As String
Dim formReset As Boolean
Dim cCommandList As String

Public stamp
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Remove button was clicked.
Private Sub Update_button_Click()
On Error GoTo Exp
Dim mdomain
Dim projn
Dim t
Dim dom
Dim loopcount
loopcount = 0

'Read the excel sheet
If cExcelFile = "" Then
    MsgBox "Excel path cannot be empty."
    Exit Sub
End If
If FExists(cExcelFile) Then
  'file exist do nothing
Else
   MsgBox "Excel import file does not exist."
    Exit Sub
End If
ReadExcel
updateMessage "Reading project selection."
'All domains selected
    If AllDomains Then
       'All project selected
       If AllProjects Then
            For Each mdomain In td.VisibleDomains
                For Each projn In td.VisibleProjects(mdomain)
                    wait (loopcount)
                    loopcount = loopcount + 1
                    updateMessage ("Connecting to TD Project [" + mdomain + "." + projn + "]")
                    If td.ProjectConnected Then
                       td.Disconnect
                    End If
                    On Error Resume Next
                    td.Connect mdomain, projn
                    If Err.Number <> 0 Then
                        updateMessage "[" & mdomain & ":" & projn & "] Err: Failed to connect."
                        Exit Sub
                    Else
                        PropagateChanges
                    End If
                Next
            Next
        Else 'not all projects
            For t = 0 To Project_list.ListCount - 1
                If Project_list.Selected(t) Then
                    projn = Mid(Project_list.List(t), InStr(1, Project_list.List(t), ".") + 1)
                    dom = Mid(Project_list.List(t), 1, InStr(1, Project_list.List(t), ".") - 1)
                    wait (loopcount)
                    loopcount = loopcount + 1
                    updateMessage ("Connecting to QC Project [" + dom + "." + projn + "].")
                    If td.ProjectConnected Then
                        td.Disconnect
                    End If
                    On Error Resume Next
                    td.Connect dom, projn
                    If Err.Number <> 0 Then
                        updateMessage "[" & dom & ":" & projn & "] Err: Failed to connect."
                        Exit Sub
                    Else
                       PropagateChanges
                    End If
                End If
            Next
        End If 'End all domains
    Else 'All Domain not seleected
        'All Project selected
        If AllProjects Then
            For t = 0 To Domain_list.ListCount - 1
                If Domain_list.Selected(t) Then
                    For Each projn In td.VisibleProjects(Domain_list.List(t))
                    wait (loopcount)
                    loopcount = loopcount + 1
                    updateMessage ("Connecting to QC Project [" + Domain_list.List(t) + "." + projn + "].")
                    If td.ProjectConnected Then
                        td.Disconnect
                    End If
                    On Error Resume Next
                    td.Connect Domain_list.List(t), projn
                    If Err.Number <> 0 Then
                        updateMessage "[" & Domain_list.List(t) & ":" & projn & "] Err: Failed to connect."
                        Exit Sub
                    Else
                    PropagateChanges
                    
                    End If
                    Next
                End If
            Next
        Else 'Not all projects selected
            For t = 0 To Project_list.ListCount - 1
                If Project_list.Selected(t) Then
                    projn = Mid(Project_list.List(t), InStr(1, Project_list.List(t), ".") + 1)
                    dom = Mid(Project_list.List(t), 1, InStr(1, Project_list.List(t), ".") - 1)
                    wait (loopcount)
                    loopcount = loopcount + 1
                    updateMessage ("Connecting to QC Project [" + dom + "." + projn + "].")
                    If td.ProjectConnected Then
                        td.Disconnect
                    End If
                    On Error Resume Next
                    td.Connect dom, projn
                    If Err.Number <> 0 Then
                        updateMessage "[" & dom & ":" & projn & "] Err: Failed to connect."
                        Exit Sub
                    Else
                        PropagateChanges
                    End If
                End If
            Next
        End If
   End If
    
If td.Connected Then
        td.Disconnect
End If
updateMessage ("Done!")
td.Connect cdomain, cproject
MsgBox ("Done!")
Exit Sub
Exp:
MsgBox (Err.Description)
updateMessage Err.Description
End Sub
Private Sub wait(loopcount As Integer)
If loopcount < 50 Then
    'do nothing
ElseIf loopcount < 80 Then
    Sleep 10000
ElseIf loopcount < 100 Then
    Sleep 15000
Else
    Sleep 20000
End If
Label2.Refresh
End Sub
'Read the excel sheet and insert it into the users_list object.
Private Sub ReadExcel()
On Error GoTo Exp
'Exit if Favorite radiobutton is enabled
'If Check11.Value = 1 Then
'    Exit Sub
'End If
updateMessage "Reading Excel file."
Dim Row As Integer
Dim FirstColumn
Dim UserToRemove
Dim myExcel

Set myExcel = CreateObject("Excel.Application")

'Close Excel file
myExcel.Workbooks.Close

'Open Excel file for reading
myExcel.Workbooks.Open cExcelFile, , True

'Bring up Excel app
myExcel.Visible = False

'Select the worksheet
myExcel.Worksheets("Sheet1").Activate

'reset list
RenameUser_list.Clear

'////Reading excel file
Row = 0
Do
    Row = Row + 1
    FirstColumn = myExcel.activesheet.cells(Row, 1).Value
    If FirstColumn = "" Or IsNull(FirstColumn) Then
        updateMessage "Entry is blank. Stop reading excel file."
    Else
        If FirstColumn = "Rename User" Then
            RenameUser_list.AddItem myExcel.activesheet.cells(Row, 2)
        End If
    End If
    
Loop While Not FirstColumn = ""


'Close Excel file
myExcel.Workbooks.Close
myExcel.Application.quit
Set myExcel = Nothing
Exit Sub
Exp:
MsgBox (Err.Description)
updateMessage Err.Description
End Sub
Private Sub PropagateChanges()
If Check26.Value = 1 Then
    Call RenameUser
End If

End Sub
Private Sub RenameUser()
'Code from Bilal
On Error GoTo Exp

Dim SecondColumn() As String
Dim txtUserName As String
Dim txtNewName As String
Dim UserIndex As Integer
Dim glist 'As TDAPIOLELib.List
Dim CustUsersGroups 'As CustomizationUsersGroups
Dim CustGroup 'As CustomizationUsersGroup
Dim CustUsers 'As CustomizationUsers
Dim CustUser 'As CustomizationUser
Dim CustUserNew 'As CustomizationUser
Dim selCmdB 'As TDAPIOLELib.Command
Dim selCmd 'As TDAPIOLELib.Command
Dim rst 'As TDAPIOLELib.Recordset
Dim rstB 'As TDAPIOLELib.Recordset
Dim Grp 'As CustomizationUsersGroup
Dim rcset 'As TDAPIOLELib.Recordset
Dim updateCmd 'As TDAPIOLELib.Command
Dim cmd 'As TDAPIOLELib.Command
Dim msg As String
Dim count As Integer
Dim selSQL As String
Dim selSQLB As String
Dim userDoesntExist As String
Dim naffectedRows As Integer
Dim tableName As String
Dim colName As String
Dim updateSQL As String
Dim i As Integer



updateMessage "Renaming " & RenameUser_list.ListCount & " users."
For UserIndex = 0 To RenameUser_list.ListCount - 1
    SecondColumn = Split(RenameUser_list.List(UserIndex), ":")
    txtUserName = Trim(LCase(SecondColumn(0)))
    txtNewName = Trim(LCase(SecondColumn(1)))


    updateMessage "Renaming " & txtUserName & " to " & txtNewName & " (" & UserIndex + 1 & " of " & RenameUser_list.ListCount & ")."
    
    
    ' -\-\-\-\- Verify that the old user exists in project -\-\-\-\-
    Set selCmd = td.Command
    
        selSQL = "SELECT COUNT(*) FROM USERS WHERE LOWER(US_USERNAME)='" & txtUserName & "'"
        selSQLB = "SELECT COUNT(*) FROM USERS WHERE LOWER(US_USERNAME)='" & txtNewName & "'"
    
    selCmd.CommandText = selSQL
    Set rst = selCmd.Execute
    userDoesntExist = rst.FieldValue(0)
    'Check if old user exist in project
    If (userDoesntExist = "0") Then
        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] Error: Old username (" & txtUserName & ") does not exist in project."
        E_out "[" & td.DomainName & ":" & td.ProjectName & "]Error: Failed to replace user " & txtUserName & " with " & txtNewName

    Else
        ' -\-\-\-\- Verify that the new user does not exist in project -\-\-\-\-
        Set selCmdB = td.Command
        selCmdB.CommandText = selSQLB
        Set rstB = selCmdB.Execute
        userDoesntExist = rstB.FieldValue(0)
        'check if new exist in project
        If (userDoesntExist <> "0") Then
            updateMessage "[" & td.DomainName & ":" & td.ProjectName & "]Error: New username (" & txtNewName & ") already exists in project."
            E_out "[" & td.DomainName & ":" & td.ProjectName & "]Error: Failed to replace user " & txtUserName & " with " & txtNewName
            Exit For
        Else
            On Error Resume Next
             ' *************** Get user group(s) ***************
            'Get customization data into the local cache
            'td.Customization.Load
            
            'Get the groups that the old user belong to
            Set CustUsers = td.Customization.Users
            Set CustUser = CustUsers.User(txtUserName)
            Set glist = CustUser.GroupsList
            
            ' *************** Add new user to the same group as the old user **************
            'Check to see if the new user exist
            If Not td.Customization.Users.UserExistsInSite(txtNewName) Then
              updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] User " & txtNewName & " does not exist in Site Admin." & Err.Description
              E_out "[" & td.DomainName & ":" & td.ProjectName & "]Error: Failed to replace user " & txtUserName & " with " & txtNewName
              Exit For
            End If
            
            'Add user to the project with the CustomizationUsers Object
            ' referenced through TDConnection.Customization.Users
            Set CustUserNew = td.Customization.Users.AddUser(txtNewName)
            If Err.Number <> 0 Then
             updateMessage "[" & td.DomainName & ":" & td.ProjectName & "]Error: User " & txtNewName & " already exist in project." & Err.Description
              E_out "[" & td.DomainName & ":" & td.ProjectName & "]Error: Failed to replace user " & txtUserName & " with " & txtNewName
              Exit For
            End If
            
            If Err.Number <> 0 Then
             updateMessage "[" & td.DomainName & ":" & td.ProjectName & "]Error: Failed to commit update." & Err.Description
              E_out "[" & td.DomainName & ":" & td.ProjectName & "]Error: Failed to replace user " & txtUserName & " with " & txtNewName
              Exit For
            End If
            
            'addToUserGroup is a group name, for example "QATester"

            For Each Grp In glist
                'Add the user
                'Grp.AddUser CStr(txtNewName)
                CustUserNew.AddToGroup (Grp.Name)
                If Err.Number <> 0 Then
                    updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] Failed to added " & txtNewName & " to group " & Grp.Name & "." & Err.Description
                End If
                            'Commit the changes to the project
            Next Grp
            td.Customization.Commit
                       
            updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] Updating project data to replace " & txtUserName & " with " & txtNewName & "."


            ' ************************** Update PROJECT DB *********************
            ' -\-\-\- Updating users table. -\-\-\-\-
            
            Set updateCmd = td.Command
            
            naffectedRows = -1
            
            ' -/-/-/- Updating project tables in which the user name is written (UserCombo) -/-/-/-/-
            Set cmd = td.Command
            cmd.CommandText = "select SF_TABLE_NAME, SF_COLUMN_NAME FROM SYSTEM_FIELD WHERE SF_EDIT_STYLE='UserCombo'"
            Set rcset = cmd.Execute
        
             ' ************************** Updating Regular tables *********************
            On Error Resume Next
            While (rcset.EOR = False)
                naffectedRows = -1
                tableName = rcset.FieldValue(0)
                colName = rcset.FieldValue(1)
                updateSQL = "UPDATE " + tableName + " SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                updateCmd.CommandText = updateSQL
                updateCmd.Execute
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
                'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                
                If tableName = "ASSET_RELATIONS" Then
                    
                    updateSQL = "UPDATE HIST_ASSET_RELATIONS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_ASSET_RELATIONS :" & Err.Description
                    Else
                        'naffectedRows = updateCmd.AffectedRows
                        'msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_ASSET_RELATIONS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_ASSET_RELATIONS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
               If tableName = "ASSET_REPOSITORY_ITEMS" Then
                    
                    updateSQL = "UPDATE HIST_ASSET_REPOSITORY_ITEMS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_ASSET_REPOSITORY_ITEMS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    updateSQL = "UPDATE VC_ASSET_REPOSITORY_ITEMS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_ASSET_REPOSITORY_ITEMS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
               If tableName = "BP_ITER_PARAM" Then
                    
                    updateSQL = "UPDATE HIST_BP_ITER_PARAM SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_BP_ITER_PARAM :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_BP_ITER_PARAM SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_BP_ITER_PARAM :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
               If tableName = "BP_ITERATION" Then
                    
                    updateSQL = "UPDATE HIST_BP_ITERATION SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_BP_ITERATION :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_BP_ITERATION SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_BP_ITERATION :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
               If tableName = "BP_PARAM" Then
                    
                    updateSQL = "UPDATE HIST_BP_PARAM SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_BP_PARAM :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                     
                    updateSQL = "UPDATE VC_BP_PARAM SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_BP_PARAM :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
                If tableName = "BPM_ELEMENTS" Then
                    
                    updateSQL = "UPDATE HIST_BPM_ELEMENTS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_BPM_ELEMENTS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                     
                    updateSQL = "UPDATE VC_BPM_ELEMENTS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_BPM_ELEMENTS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
              
                If tableName = "BPM_MODELS" Then
                    
                    updateSQL = "UPDATE HIST_BPM_MODELS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_BPM_MODELS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                     
                    updateSQL = "UPDATE VC_BPM_MODELS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_BPM_MODELS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
              
              
                If tableName = "BPM_PATHS" Then
                    
                    updateSQL = "UPDATE HIST_BPM_PATHS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_BPM_PATHS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                     
                    updateSQL = "UPDATE VC_BPM_PATHS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_BPM_PATHS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
              
               If tableName = "BPTEST_TO_COMPONENTS" Then
                    
                    updateSQL = "UPDATE HIST_BPTEST_TO_COMPONENTS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_BPTEST_TO_COMPONENTS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_BPTEST_TO_COMPONENTS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_BPTEST_TO_COMPONENTS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
               If tableName = "COMPONENT" Then
                    
                    updateSQL = "UPDATE HIST_COMPONENT SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_COMPONENT :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_COMPONENT SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_COMPONENT :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
              
                If tableName = "COMPONENT_STEP" Then
                    
                    updateSQL = "UPDATE HIST_COMPONENT_STEP SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_COMPONENT_STEP :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_COMPONENT_STEP SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_COMPONENT_STEP :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
              
                If tableName = "COMPONENT_STEP_PARAMS" Then
                    
                    updateSQL = "UPDATE HIST_COMPONENT_STEP_PARAMS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_COMPONENT_STEP_PARAMS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_COMPONENT_STEP_PARAMS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_COMPONENT_STEP_PARAMS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
              
                If tableName = "CROS_REF" Then
                    
                    updateSQL = "UPDATE HIST_CROS_REF SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_CROS_REF :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_CROS_REF SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_CROS_REF :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
                If tableName = "DESSTEPS" Then
                    
                    updateSQL = "UPDATE HIST_DESSTEPS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_DESSTEPS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_DESSTEPS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_DESSTEPS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
                 If tableName = "FACETS" Then
                    
                    updateSQL = "UPDATE HIST_FACETS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_FACETS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_FACETS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_FACETS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
                If tableName = "FRAMEWORK_PARAM" Then
                    
                    updateSQL = "UPDATE HIST_FRAMEWORK_PARAM SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_FRAMEWORK_PARAM :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_FRAMEWORK_PARAM SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_FRAMEWORK_PARAM :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
                If tableName = "REQ" Then
                    
                    updateSQL = "UPDATE HIST_REQ SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_REQ :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_REQ SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_REQ :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
              
                If tableName = "RESOURCES" Then
                    
                    updateSQL = "UPDATE HIST_RESOURCES SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_RESOURCES :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_RESOURCES SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_RESOURCES :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
                If tableName = "RUNTIME_PARAM" Then
                    
                    updateSQL = "UPDATE HIST_RUNTIME_PARAM SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_RUNTIME_PARAM :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_RUNTIME_PARAM SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_RUNTIME_PARAM :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
                If tableName = "STEP_PARAMS" Then
                    
                    updateSQL = "UPDATE HIST_STEP_PARAMS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_STEP_PARAMS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_STEP_PARAMS set " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_STEP_PARAMS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
                If tableName = "TEST" Then
                    
                    updateSQL = "UPDATE HIST_TEST SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_TEST :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_TEST SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_TEST :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
                 If tableName = "TEST_CONFIGS" Then
                    
                    updateSQL = "UPDATE HIST_TEST_CONFIGS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_TEST_CONFIGS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_TEST_CONFIGS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_TEST_CONFIGS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
                 If tableName = "TEST_CRITERIA" Then
                    
                    updateSQL = "UPDATE HIST_TEST_CRITERIA SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_TEST_CRITERIA :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_TEST_CRITERIA SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_TEST_CRITERIA :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
               If tableName = "TEST_PARAMS" Then
                    
                    updateSQL = "UPDATE HIST_TEST_PARAMS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_TEST_PARAMS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                      '  updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_TEST_PARAMS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_TEST_PARAMS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                       ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               
               If tableName = "USER_ASSETS" Then
                    
                    updateSQL = "UPDATE HIST_USER_ASSETS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] HIST_USER_ASSETS :" & Err.Description
                    Else
                        naffectedRows = updateCmd.AffectedRows
                        msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
                    
                    updateSQL = "UPDATE VC_USER_ASSETS SET " + colName + " = '" + txtNewName + "' WHERE lower(" + colName + ") = '" + txtUserName + "'"
                    updateCmd.CommandText = updateSQL
                    Err.Clear
                    updateCmd.Execute
                    If Err.Number <> 0 Then
                        updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_USER_ASSETS :" & Err.Description
                    Else
                        'naffectedRows = updateCmd.AffectedRows
                        'msg = updateSQL & " - Affected " & naffectedRows & " lines."
                        'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
                    End If
               End If
               rcset.Next
            Wend

            
                    
            'VC_COMPONENT_MULTIVALUE
            updateSQL = "UPDATE VC_COMPONENT_MULTIVALUE SET COMV_USER_NAME = '" + txtNewName + "' WHERE lower(COMV_USER_NAME) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_COMPONENT_MULTIVALUE :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
               ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If
            
            'VC_REQ_MULTIVALUE
            updateSQL = "UPDATE VC_REQ_MULTIVALUE SET RQMV_USER_NAME = '" + txtNewName + "' WHERE lower(RQMV_USER_NAME) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_REQ_MULTIVALUE :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
                'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If

            'VC_RESOURCES_MULTIVALUE
            updateSQL = "UPDATE VC_RESOURCES_MULTIVALUE SET RSCMV_USER_NAME = '" + txtNewName + "' WHERE lower(RSCMV_USER_NAME) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_RESOURCES_MULTIVALUE :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
               ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If

            
            'VC_TEST_MULTIVALUE
            updateSQL = "UPDATE VC_TEST_MULTIVALUE SET TSMV_USER_NAME = '" + txtNewName + "' WHERE lower(TSMV_USER_NAME ) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_TEST_MULTIVALUE :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
                'updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If
                        
            
             'VC_DELETED_ASSETS_INFO
            'updateSQL = "UPDATE VC_DELETED_ASSETS_INFO SET DAI_DELETED_BY = '" + txtNewName + "' WHERE lower(DAI_DELETED_BY) = '" + txtUserName + "'"
            'updateCmd.CommandText = updateSQL
            'Err.Clear
            'updateCmd.Execute
            'If Err.Number <> 0 Then
            '    updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] VC_DELETED_ASSETS_INFO :" & Err.Description
            'Else
            '    naffectedRows = updateCmd.AffectedRows
            '    msg = updateSQL & " - Affected " & naffectedRows & " lines."
            '   ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            'End If
           
             'DELETED_ASSETS_INFO
            'updateSQL = "UPDATE DELETED_ASSETS_INFO SET DAI_DELETED_BY = '" + txtNewName + "' WHERE lower(DAI_DELETED_BY) = '" + txtUserName + "'"
            'updateCmd.CommandText = updateSQL
            'Err.Clear
            'updateCmd.Execute
            'If Err.Number <> 0 Then
            '    updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] DELETED_ASSETS_INFO :" & Err.Description
            'Else
            '    naffectedRows = updateCmd.AffectedRows
            '    msg = updateSQL & " - Affected " & naffectedRows & " lines."
            '   ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            'End If
            
            ' ************************** Updating Common_Settings table *********************
            updateSQL = "UPDATE COMMON_SETTINGS SET CSET_OWNER = '" + txtNewName + "' WHERE lower(CSET_OWNER) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] COMMON_SETTINGS :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
               ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If
            
            
              ' ************************** Updating Favorites table *********************
            updateSQL = "UPDATE FAVORITES SET FAV_OWNER = '" + txtNewName + "' WHERE lower(FAV_OWNER) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] FAVORITES :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
               ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If
             
            updateSQL = "UPDATE FAVORITE_FOLDERS SET FF_OWNER = '" + txtNewName + "' WHERE lower(FF_OWNER) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] FAVORITE_FOLDERS :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
               ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If
             
             ' ************************** Updating Analysis tables *********************
            updateSQL = "UPDATE ANALYSIS_ITEM_FOLDERS SET AIF_OWNER = '" + txtNewName + "' WHERE lower(AIF_OWNER) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] ANALYSIS_ITEM_FOLDERS :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
               ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If
             
            updateSQL = "UPDATE ANALYSIS_ITEMS SET AI_OWNER = '" + txtNewName + "' WHERE lower(AI_OWNER) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] ANALYSIS_ITEMS :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
               ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If
            
            updateSQL = "UPDATE ANALYSIS_ITEMS SET AI_MODIFIED_BY = '" + txtNewName + "' WHERE lower(AI_MODIFIED_BY) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] ANALYSIS_ITEMS :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
               ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If

             ' ************************** Updating Dashboard tables *********************
            updateSQL = "UPDATE DASHBOARD_FOLDERS SET DF_OWNER = '" + txtNewName + "' WHERE lower(DF_OWNER) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] DASHBOARD_FOLDERS :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
               ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If
             
          
            
            updateSQL = "UPDATE DASHBOARD_PAGES SET DP_OWNER = '" + txtNewName + "' WHERE lower(DP_OWNER) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] DASHBOARD_PAGES :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
               ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If
            
            
            updateSQL = "UPDATE DASHBOARD_PAGES SET DP_MODIFIED_BY = '" + txtNewName + "' WHERE lower(DP_MODIFIED_BY) = '" + txtUserName + "'"
            updateCmd.CommandText = updateSQL
            Err.Clear
            updateCmd.Execute
            If Err.Number <> 0 Then
                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] DASHBOARD_PAGES :" & Err.Description
            Else
                naffectedRows = updateCmd.AffectedRows
                msg = updateSQL & " - Affected " & naffectedRows & " lines."
               ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
            End If
'            ' ************************** Updating Audti Log table *********************
'            updateSQL = "UPDATE AUDIT_LOG SET AU_USER = '" + txtNewName + "' WHERE lower(AU_USER) = '" + txtUserName + "'"
'            updateCmd.CommandText = updateSQL
'            Err.Clear
'            updateCmd.Execute
'            If Err.Number <> 0 Then
'                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] AUDIT_LOG :" & Err.Description
'            Else
'                naffectedRows = updateCmd.AffectedRows
'                msg = updateSQL & " - Affected " & naffectedRows & " lines."
'                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
'            End If
            
'            ' ************************** Updating Audti Properties table *********************
'            updateSQL = "UPDATE AUDIT_PROPERTIES SET AP_OLD_VALUE = '" + txtNewName + "' WHERE lower(AP_OLD_VALUE) = '" + txtUserName + "'"
'            updateCmd.CommandText = updateSQL
'            Err.Clear
'            updateCmd.Execute
'            If Err.Number <> 0 Then
'                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] AUDIT_PROPERTIES :" & Err.Description
'            Else
'                naffectedRows = updateCmd.AffectedRows
'                msg = updateSQL & " - Affected " & naffectedRows & " lines."
'                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
'            End If

'            updateSQL = "UPDATE AUDIT_PROPERTIES SET  AP_NEW_VALUE = '" + txtNewName + "' WHERE lower( AP_NEW_VALUE) = '" + txtUserName + "'"
'            updateCmd.CommandText = updateSQL
'            Err.Clear
'            updateCmd.Execute
'            If Err.Number <> 0 Then
'                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] AUDIT_PROPERTIES :" & Err.Description
'            Else
'                naffectedRows = updateCmd.AffectedRows
'                msg = updateSQL & " - Affected " & naffectedRows & " lines."
'                updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] " & msg
'            End If
            Err.Clear
            'remove old user
            td.Customization.Users.RemoveUser txtUserName
            If Err.Number <> 0 Then
              updateMessage "[" & td.DomainName & ":" & td.ProjectName & "]Error: Failed to commit update." & Err.Description
            Else
            ' updateMessage "[" & td.DomainName & ":" & td.ProjectName & "] Removed " & txtUserName & " from project." & Err.Description
            End If
            
            'Commit the changes to the project
            td.Customization.Commit
            If Err.Number <> 0 Then
             updateMessage "[" & td.DomainName & ":" & td.ProjectName & "]Error: Failed to commit update." & Err.Description
              Exit For
            End If

            C_out "[" & td.DomainName & ":" & td.ProjectName & "] Completed update to replace user " & txtUserName & " with " & txtNewName

        End If
     
    End If 'end of check if user exist
    
Next ' end of reading users list to rename

Set rst = Nothing
Set selCmd = Nothing
Set selCmdB = Nothing
Set rstB = Nothing
Set CustUsers = Nothing
Set CustUser = Nothing
Set glist = Nothing
Set CustUsersGroups = Nothing
Set CustGroup = Nothing
Set updateCmd = Nothing
Set rcset = Nothing
Set cmd = Nothing

Exit Sub
Exp:
updateMessage Err.Description
End Sub

'Select project from list
Private Sub Project_button_Click()
On Error GoTo Exp

If AllProject_checkbox.Value = 1 Then
    AllProjects = True
End If
If AllProjects = True Then
Else
    If Project_list.SelCount < 1 Then
        MsgBox ("Please select 1 or more domains !")
        Exit Sub
    End If
End If
Customize.Visible = True
SelectTarget.Enabled = False
Project_button.Enabled = False
Project_list.Enabled = False
AllProject_checkbox.Enabled = False
Exit Sub
Exp:
MsgBox (Err.Description)
updateMessage Err.Description
End Sub
'Select domain from list
Private Sub Domain_button_Click()
On Error GoTo Exp
Dim domainItem
Dim projn
Dim t
If AllDomain_checkbox.Value = 1 Then
    AllDomains = True
Else
    AllDomains = False
End If
If AllDomains = True Then
  Project_list.Clear
  For Each domainItem In td.VisibleDomains
    For Each projn In td.VisibleProjects(domainItem)
        Project_list.AddItem domainItem + "." + projn
    Next
  Next
Else
    If Domain_list.SelCount < 1 Then
    MsgBox ("Please select 1 or more domains !")
    Exit Sub
    End If
    Project_list.Clear
    For t = 0 To Domain_list.ListCount - 1
        If Domain_list.Selected(t) Then
            For Each projn In td.VisibleProjects(Domain_list.List(t))
                Project_list.AddItem Domain_list.List(t) + "." + projn
            Next
        End If
    Next
End If

'Disabled the domain selection
Domain_button.Enabled = False
Domain_list.Enabled = False
AllDomain_checkbox.Enabled = False

'Enable the project selection
Project_list.Visible = True
Project_button.Visible = True
AllProject_checkbox.Visible = True
'updateMessage ("Ready.")
Exit Sub
Exp:
MsgBox (Err.Description)
updateMessage Err.Description
End Sub
'Initialize Server
Private Sub Connect_button_Click()
On Error GoTo Exp
Dim mdomain
formReset = False
comcl = False
updateMessage ("Initializing QC Server connection ...")
' Initialize connection to TDServer
Set td = CreateObject("TDApiOle80.TDConnection.1")
td.InitConnectionEx (Server_textbox.Text)
isconnect = True
SourceConnection.Show (1)

If Not comcl Then
    MsgBox "Failed to connect to source project."
    Exit Sub
End If
For Each mdomain In td.VisibleDomains
    Domain_list.AddItem mdomain
Next

'server connection field disabled
Connect_button.Enabled = False
Server_textbox.Enabled = False

'Make the Select Target Projects frame visible
SelectTarget.Visible = True
'updateMessage ("Ready")
'domain connection
AllDomain_checkbox.Visible = True
Domain_list.Visible = True
Domain_button.Visible = True
'project connection
Project_button.Visible = False
Project_list.Visible = False
AllProject_checkbox.Visible = False

'customize frame
Customize.Visible = False
Exit Sub
Exp:
MsgBox (Err.Description)
updateMessage Err.Description
End Sub
'Reset the form

'Unload the form and disconnect from TD
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Exp
updateMessage "Closing Quality Center connection and exiting RenameUsers tool."
'updateMessage "Server end time is " & CStr(td.ServerTime) & "."
If td.Connected Then
    If td.ProjectConnected Then
        td.Disconnect
        td.Logout
    End If
End If
td.ReleaseConnection
Exp:
updateMessage Err.Description
End Sub
'Load the form
Private Sub Form_Load()
filename = Format(Now, "mm-dd-yy hh_mm_ss AM/PM")
SelectTarget.Visible = False
Customize.Visible = False
cLogPath = Log_textbox.Text
FoldExists (cLogPath)
cExcelFile = Excel_textbox.Text
'####################################Command Line######################
Dim a_strArgs() As String
Dim i As Integer
Dim helpInfo As String

helpInfo = "Please entered the required parameters:" & Chr(13) & Chr(10) & _
        "AutomateChange.exe /s:ServerURL /d:Domain /p:Project /l:List /c:CommandFile /i:Log /f:ProjectFile /e:Excelfile" & Chr(13) & Chr(10) & _
        "Where:" & Chr(13) & Chr(10) & _
        "   /s QC server URL" & Chr(13) & Chr(10) & _
        "   /u User name" & Chr(13) & Chr(10) & _
        "   /w User password" & Chr(13) & Chr(10) & _
        "   /d Domain Name" & Chr(13) & Chr(10) & _
        "   /p Project Name" & Chr(13) & Chr(10) & _
        "   /l List (list separate by commas)" & Chr(13) & Chr(10) & _
        "   /i Log path" & Chr(13) & Chr(10) & _
        "   /e Excel path" & Chr(13) & Chr(10) & _
        "   /f Project file"
If VBA.Command() = "" Then
    'Run GUI version
    Call RenameFile
Else
    'Run Command Line version
    a_strArgs = Split(VBA.Command$, " ")
    If UBound(a_strArgs) <> 8 Then
        MsgBox helpInfo
        Unload Me
        Exit Sub
    End If
    For i = LBound(a_strArgs) To UBound(a_strArgs)
        Select Case LCase(Left(a_strArgs(i), 3))
        Case "/s:"
            'QC URL is specified
            If Mid(a_strArgs(i), 4) = "" Then
                MsgBox "Invalid URL."
                Exit Sub
            Else
                cServerURL = Mid(a_strArgs(i), 4)
            End If
        Case "/u:"
            'User is specified
            If Mid(a_strArgs(i), 4) = "" Then
                MsgBox "Invalid User name."
                Exit Sub
            Else
                cuser = Mid(a_strArgs(i), 4)
            End If
        Case "/w:"
            'Password is specified
            If Mid(a_strArgs(i), 4) = "" Then
                MsgBox "Invalid password."
                Exit Sub
            Else
                cpassword = Mid(a_strArgs(i), 4)
            End If
        Case "/d:"
            'Domain is specified
            If Mid(a_strArgs(i), 4) = "" Then
                MsgBox "Invalid Domain."
                Exit Sub
            Else
                cdomain = Mid(a_strArgs(i), 4)
            End If
        Case "/p:"
            'Project is specified
            If Mid(a_strArgs(i), 4) = "" Then
                MsgBox "Invalid password."
                Exit Sub
            Else
                cproject = Mid(a_strArgs(i), 4)
            End If
        Case "/c:"
            'Command list is specified
            If Mid(a_strArgs(i), 4) = "" Then
                MsgBox "File not specified."
                Exit Sub
            Else
                cCommandList = Mid(a_strArgs(i), 4)
            End If
         Case "/i:"
            'Log directory is specified
            If Mid(a_strArgs(i), 4) = "" Then
            Else
                cLogPath = Mid(a_strArgs(i), 4)
            End If
       Case "/l:"
            'List of action is specified
            If Mid(a_strArgs(i), 4) = "" Then
                MsgBox "Invalid action list."
                Exit Sub
            Else
                cCommandList = Mid(a_strArgs(i), 4)
            End If
         Case "/f:"
            'Project File is specified
            If Mid(a_strArgs(i), 4) = "" Then
                MsgBox "Invalid file path."
                Exit Sub
            Else
                cProjectsFile = Mid(a_strArgs(i), 4)
            End If
         Case "/e:"
            'Excel File is specified
            If Mid(a_strArgs(i), 4) = "" Then
                MsgBox "Invalid excel file path."
                Exit Sub
            Else
                cExcelFile = Mid(a_strArgs(i), 4)
            End If
       Case Else
            MsgBox "Invalid argument: " & a_strArgs(i)
      End Select
    
   Next i
   cCommandLine = True
   Call StartCheck
   Unload SourceConnection
   Unload Me
End If

End Sub
Public Sub StartCheck()
Call RenameFile
Call ConnectToSource
Call LoadProjects
Call ReadExcel
Dim t As Integer
Dim projn As String
Dim dom As String
For t = 0 To Project_list.ListCount - 1
    projn = Mid(Project_list.List(t), InStr(1, Project_list.List(t), ".") + 1)
    dom = Mid(Project_list.List(t), 1, InStr(1, Project_list.List(t), ".") - 1)
    updateMessage ("Connecting to QC Project [" + dom + "." + projn + "].")
    If td.ProjectConnected Then
        td.Disconnect
    End If
    td.Connect dom, projn
    If Err.Number <> 0 Then
        updateMessage "[" & dom & ":" & projn & "] Err:" & Err.Description
        Exit Sub
    Else
        Call PropagateChangesCommandLine
    End If
Next
End Sub
Sub RenameFile()
'Rename old log files
Dim OldFileName As String
Dim NewFileName As String
Dim OldFileNameArchieve As String
If FExists(cLogPath & "\AutomateChange.txt") Then
OldFileName = cLogPath & "\AutomateChange.txt"
NewFileName = cLogPath & "\AutomateChange_" & filename & ".txt"
Name OldFileName As NewFileName
End If

If FExists(cLogPath & "\Error.txt") Then
OldFileName = cLogPath & "\Error.txt"
NewFileName = cLogPath & "\Error_" & filename & ".txt"
Name OldFileName As NewFileName
End If

If FExists(cLogPath & "\Completed.txt") Then
OldFileName = cLogPath & "\Completed.txt"
NewFileName = cLogPath & "\Completed_" & filename & ".txt"
Name OldFileName As NewFileName
End If

End Sub
Sub ConnectToSource()
On Error GoTo Exp
    Dim cust 'As Customization
    ' Initialize connection to TDServer
    Set td = CreateObject("TDApiOle80.TDConnection.1")
    td.InitConnectionEx (cServerURL)
    isconnect = True
    
    'Authenticate the user
    If td.Connected = True Then
        td.Login cuser, cpassword
        If Err.Number <> 0 Then
            updateMessage "Invalid: Failed to authenticate."
            Unload Form1
        End If
    Else
        updateMessage "Invalid: Not connected to server."
        Unload Form1
    End If
    'Connecting to the Template project
    On Error Resume Next
    td.Connect cdomain, cproject
    If Err.Number <> 0 Then
            MsgBox "Cannot connect to project"
            updateMessage "Cannot connect to source project."
    Else
        comcl = True
        stamp = td.ServerTime
        updateMessage "Server start time is " & CStr(Form1.td.ServerTime) & "."
    End If
Exp:
updateMessage Err.Description
End Sub
Public Sub LoadProjects()
On Error GoTo Exp
Dim Row As Integer
Dim FirstColumn
Dim UserToRemove
Dim myExcel

Set myExcel = CreateObject("Excel.Application")

'Close Excel file
myExcel.Workbooks.Close

'Open Excel file for reading
myExcel.Workbooks.Open cProjectsFile, , True

'Select the worksheet
myExcel.Worksheets("Sheet1").Activate

'reset list
Project_list.Clear

'////Reading excel file
Row = 0
Do
    Row = Row + 1
    FirstColumn = myExcel.activesheet.cells(Row, 1).Value
    If FirstColumn = "" Or IsNull(FirstColumn) Then
        updateMessage "Project count:" & Row - 1
    Else
            Project_list.AddItem myExcel.activesheet.cells(Row, 1)
    End If
Loop While Not FirstColumn = ""

'Close Excel file
myExcel.Workbooks.Close
Set myExcel = Nothing
Exit Sub
Exp:
updateMessage Err.Description
End Sub
Private Sub PropagateChangesCommandLine()
Dim CommandList

CommandList = Split(cCommandList, ",")
Dim i As Integer
Dim CommandName As String
For i = 0 To UBound(CommandList)
    CommandName = CStr(CommandList(i))
    Select Case CommandName
    Case "RenameUser"
        Call RenameUser
    End Select
Next
End Sub
Function GetItem(ifrom As String, ito As String, original As String) As String
Dim temp
temp = Mid(original, InStr(1, original, ifrom) + Len(ifrom))
GetItem = Mid(temp, 1, InStr(1, temp, ito) - 1)
End Function
'Write to updatemessage file
Public Function M_out(m1 As String)
FoldExists (cLogPath)
Open cLogPath & "\AutomateChange.txt" For Append As #5
Write #5, m1
Close #5
End Function

'Write to updatemessage file
Public Function E_out(m1 As String)
FoldExists (cLogPath)
Open cLogPath & "\Error.txt" For Append As #5
Write #5, Now & " :" & m1
Close #5
End Function

'Write to updatemessage file
Public Function C_out(m1 As String)
FoldExists (cLogPath)
Open cLogPath & "\Completed.txt" For Append As #5
Write #5, Now & " :" & m1
Close #5
End Function

Function updateMessage(m1 As String)
Label2.Caption = m1
Label2.Refresh
M_out (m1)
End Function
'This function checks if required folders exists and if not, will create the folder
Public Sub FoldExists(Folder As String)
Dim MyFileSystem

Set MyFileSystem = CreateObject("Scripting.FileSystemObject")

    If MyFileSystem.FolderExists(Folder) = False Then
      CreateFolder (Folder & "\")
    End If
Set MyFileSystem = Nothing
End Sub
Public Function FExists(OrigFile As String)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
FExists = fs.FileExists(OrigFile)
End Function

Private Sub Log_textbox_Change()
cLogPath = Log_textbox.Text
End Sub
Private Sub Excel_textbox_Change()
cExcelFile = Excel_textbox.Text
End Sub
'Code from the internet
Public Function CreateFolder(destDir As String) As Boolean
   Dim i As Long
   Dim prevDir As String
   On Error Resume Next
   For i = Len(destDir) To 1 Step -1
       If Mid(destDir, i, 1) = "\" Then
           prevDir = Left(destDir, i - 1)
           Exit For
       End If
   Next i
   If prevDir = "" Then CreateFolder = False: Exit Function
   If Not Len(Dir(prevDir & "\", vbDirectory)) > 0 Then
       If Not CreateFolder(prevDir) Then CreateFolder = False: Exit Function
   End If
   On Error GoTo errDirMake
   MkDir destDir
   CreateFolder = True
   Exit Function
errDirMake:
   CreateFolder = False
End Function

