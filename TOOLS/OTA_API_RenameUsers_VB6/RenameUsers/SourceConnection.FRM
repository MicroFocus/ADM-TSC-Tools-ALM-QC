VERSION 5.00
Begin VB.Form SourceConnection 
   Caption         =   "Authenticate"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form2"
   ScaleHeight     =   2730
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Password_text 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton Authenticate_button 
      Caption         =   "Authenticate"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.ComboBox Project_combobox 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton ConnectToProject_button 
      Caption         =   "Connect"
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox User_text 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Domain:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label TDConnection 
      Caption         =   "Password:"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "User name:"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Project:"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "SourceConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Authenticate_button_Click()
    Dim domain_name
    
    If User_text.Text = "" Then
        MsgBox "Invalid: Username is empty."
        Exit Sub
    End If
    If Password_text.Text = "" Then
        MsgBox "Invalid: Password is empty."
        Exit Sub
    End If
    
    'Authenticate the user
    If Form1.td.Connected = True Then
        On Error Resume Next
        Form1.td.Login CStr(User_text.Text), CStr(Password_text.Text)
        If Err.Number <> 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        If Form1.td.LoggedIn = False Then
            MsgBox "Invalid: Failed to authenticate."
            Exit Sub
        Else
            Authenticate_button.Enabled = False
            For Each domain_name In Form1.td.VisibleDomains
                Combo1.AddItem domain_name
            Next
        End If
    End If
End Sub
Private Sub Combo1_Click()
Project_combobox.Clear
For Each Project_name In Form1.td.VisibleProjects(Combo1.Text)
    Project_combobox.AddItem Project_name
Next
End Sub
'Handle Connection to server event
Private Sub ConnectToProject_button_Click()
    'Storing the Template project connection information
    Form1.cuser = User_text.Text
    Form1.cpassword = Password_text.Text
    Form1.cdomain = Combo1.Text
    Form1.cproject = Project_combobox.Text
    
    'Connecting to the Template project
    On Error Resume Next
    Form1.td.Connect Combo1.Text, Project_combobox.Text
    If Err.Number <> 0 Then
            MsgBox "Cannot connect to project"
            Form1.updateMessage "Cannot connect to source project."
    Else
        'Retrieving the group name
        Form1.comcl = True
        Form1.stamp = Form1.td.ServerTime
        Form1.updateMessage "Connecting to source " & Combo1.Text & ":" & Project_combobox.Text & "."
        Form1.updateMessage "Server start time is " & CStr(Form1.td.ServerTime) & "."
        Unload Me
    End If
End Sub

