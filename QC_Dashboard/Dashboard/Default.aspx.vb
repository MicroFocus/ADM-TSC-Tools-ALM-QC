Public Class _Default
    Inherits System.Web.UI.Page
    'session("tdc") As Object 'As New TDAPIOLELib TDConnection
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            EnableServerInit(True)
            EnableUserPassEntry(False)
            EnableDomainSelect(False)
            EnableProjectSelect(False)
            EnableConnect(False)
            Status.Text = ""
            Session("connected") = False
        End If
    End Sub

    Protected Sub InitializeButton_Click(sender As Object, e As EventArgs) Handles InitializeButton.Click
        On Error Resume Next
        'add comment

        If ServerURLTextbox.Text <> "" Then 'test for blank entry here
            If InStr(ServerURLTextbox.Text, "http://") > 0 Or InStr(ServerURLTextbox.Text, "https://") > 0 Then
                Err.Clear()
                Session("tdc") = CreateObject("TDApiOle80.TDConnection.1")
                Session("tdc").InitConnectionEx(ServerURLTextbox.Text)
                If Err.Number <> 0 Then
                    MsgBox(Err.Description)
                    ServerURLTextbox.Focus()
                Else
                    EnableServerInit(False)
                    EnableUserPassEntry(True)
                    EnableDomainSelect(False)
                    EnableProjectSelect(False)
                    EnableConnect(False)
                End If
            Else
                MsgBox("Please enter a valid ALM URL.")
            End If

        Else
            MsgBox("ALM URL cannot be empty.")
        End If
    End Sub
    Protected Sub AuthenticateButton_Click(sender As Object, e As EventArgs) Handles AuthenticateButton.Click
        On Error Resume Next
        If UserNameTextbox.Text <> "" And PasswordTextbox.Text <> "" Then
            Err.Clear()
            Session("serverurl") = ServerURLTextbox.Text
            Session("username") = UserNameTextbox.Text
            Session("password") = PasswordTextbox.Text
            Session("tdc").Login(UserNameTextbox.Text, PasswordTextbox.Text)

            'verify
            If Err.Number <> 0 Then
                Status.Text = Err.Description
            Else
                Dim mdomain As Object
                For Each mdomain In Session("tdc").VisibleDomains
                    DomainList.Items.Add(mdomain)
                Next mdomain
                EnableServerInit(False)
                EnableUserPassEntry(False)
                EnableDomainSelect(True)
                EnableProjectSelect(False)
                EnableConnect(False)

            End If

        Else
            If UserNameTextbox.Text = "" Then
                Status.Text = "User name cannot be empty."
            End If
            If PasswordTextbox.Text = "" Then
                Status.Text = "Password cannot be empty."
            End If
        End If
    End Sub
    Protected Sub SelectDomainButton_Click(sender As Object, e As EventArgs) Handles SelectDomainButton.Click

        If Not DomainList.SelectedItem.Selected Then
            MsgBox("Please select a domain!")
            Exit Sub
        Else
            For Each proj In Session("tdc").VisibleProjects(DomainList.SelectedItem.Text)
                ProjectList.Items.Add(DomainList.SelectedItem.Text & "." + proj)
            Next proj
            EnableServerInit(False)
            EnableUserPassEntry(False)
            EnableDomainSelect(False)
            EnableProjectSelect(True)
            EnableConnect(False)
            'Disable domain list item
            For Each item As ListItem In DomainList.Items
                item.Attributes.Add("disabled", "disabled")
            Next
            Session("domain") = DomainList.SelectedItem.Text
        End If
    End Sub
    Protected Sub SelectProjectButton_Click(sender As Object, e As EventArgs) Handles SelectProjectButton.Click
        If Not ProjectList.SelectedItem.Selected Then
            MsgBox("Please select a project!")
            Exit Sub
        Else
            EnableServerInit(False)
            EnableUserPassEntry(False)
            EnableDomainSelect(False)
            EnableProjectSelect(False)
            EnableConnect(True)
            'Disable domain list item
            For Each domainitem As ListItem In DomainList.Items
                domainitem.Attributes.Add("disabled", "disabled")
            Next
            'Disable project list item
            For Each projectitem As ListItem In ProjectList.Items
                projectitem.Attributes.Add("disabled", "disabled")
            Next
            Session("project") = ProjectList.SelectedItem.Text
        End If

        'Disconnect after information has been retrieved
        If Not Session("tdc") Is Nothing Then
            If Session("tdc").LoggedIn = True Then
                Session("tdc").Logout
            End If
            If Session("tdc").Connected = True Then
                Session("tdc").Disconnect
            End If
        End If

    End Sub
    Protected Sub ConnectButton_Click(sender As Object, e As EventArgs) Handles ConnectButton.Click
        Session("connected") = True

        Server.Transfer("Home.aspx")
    End Sub
    Private Sub EnableServerInit(ByRef toEnable As Boolean)
        InitializeButton.Enabled = toEnable
        ServerURLTextbox.Enabled = toEnable
    End Sub

    Private Sub EnableUserPassEntry(ByRef toEnable As Boolean)
        UserNameTextbox.Enabled = toEnable
        PasswordTextbox.Enabled = toEnable
        AuthenticateButton.Enabled = toEnable
    End Sub

    Private Sub EnableDomainSelect(ByRef toEnable As Boolean)

        DomainList.Enabled = toEnable
        SelectDomainButton.Enabled = toEnable

    End Sub

    Private Sub EnableProjectSelect(ByRef toEnable As Boolean)

        ProjectList.Enabled = toEnable
        SelectProjectButton.Enabled = toEnable
    End Sub
    Private Sub EnableConnect(ByRef toEnable As Boolean)
        ConnectButton.Enabled = toEnable
    End Sub

    Private Sub _Default_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        If Not Page.IsPostBack Then
            If Not Session("tdc") Is Nothing Then
                If Session("tdc").LoggedIn = True Then
                    Session("tdc").Logout
                End If
                If Session("tdc").Connected = True Then
                    Session("tdc").Disconnect
                End If
            End If
        End If
    End Sub
End Class