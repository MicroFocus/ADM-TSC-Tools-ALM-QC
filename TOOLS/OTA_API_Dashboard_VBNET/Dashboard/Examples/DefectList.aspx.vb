Public Class DefectList
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("connected") Is Nothing Or Session("connected") = False Then
            Server.Transfer("~/Default.aspx")
            'done
        End If

    End Sub
    Protected Sub RunButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RunButton.Click
        On Error Resume Next
        Dim bugobj
        Dim BugFact
        Dim mydatatable As DataTable
        Dim myNewRow As DataRow

        mydatatable = New DataTable("Defects")

        mydatatable.Columns.Add("Defect ID", System.Type.GetType("System.String"))
        mydatatable.Columns.Add("Status", System.Type.GetType("System.String"))
        mydatatable.Columns.Add("Detected By", System.Type.GetType("System.String"))
        mydatatable.Columns.Add("Summary", System.Type.GetType("System.String"))


        'Get bugfactory
        Session("tdc").InitConnectionEx(Session("serverurl"))
        Session("tdc").Login(Session("username"), Session("password"))
        Session("tdc").Connect(Session("domain"), Session("Project"))

        BugFact = Session("tdc").BugFactory

        'Get a bug object.
        For Each bugobj In BugFact.newlist("")

            myNewRow = mydatatable.NewRow()
            myNewRow("Defect ID") = bugobj.Field("BG_BUG_ID")
            myNewRow("Status") = bugobj.Field("BG_STATUS")
            myNewRow("Detected By") = bugobj.Field("BG_DETECTED_BY")
            myNewRow("Summary") = bugobj.Field("BG_SUMMARY")
            mydatatable.Rows.Add(myNewRow)
        Next


        If Err.Number <> 0 Then
            LogText.Text = Err.Description
        Else
            LogText.Text = "Defect list generated successfully."
            If Session("tdc").LoggedIn = True Then
                Session("tdc").Logout
            End If
            If Session("tdc").Connected = True Then
                Session("tdc").Disconnect
            End If
        End If

        GridView1.DataSource = mydatatable
        GridView1.DataBind()
    End Sub
End Class