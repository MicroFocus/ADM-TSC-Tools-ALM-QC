Public Class Home
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("connected") Is Nothing Or Session("connected") = False Then
            Server.Transfer("~/Default.aspx")
        End If

        Dim mydatatable As DataTable
        Dim myNewRow As DataRow


        mydatatable = New DataTable("Server Information")

        mydatatable.Columns.Add("Server Info", System.Type.GetType("System.String"))
        mydatatable.Columns.Add("Value", System.Type.GetType("System.String"))

        myNewRow = mydatatable.NewRow()
        myNewRow("Server Info") = "Server URL"
        myNewRow("Value") = Session("serverurl")
        mydatatable.Rows.Add(myNewRow)

        myNewRow = mydatatable.NewRow()
        myNewRow("Server Info") = "Project"
        myNewRow("Value") = Session("project")
        mydatatable.Rows.Add(myNewRow)

        myNewRow = mydatatable.NewRow()
        myNewRow("Server Info") = "User"
        myNewRow("Value") = Session("username")
        mydatatable.Rows.Add(myNewRow)

        GridView1.DataSource = mydatatable
        GridView1.DataBind()

    End Sub

    Protected Sub LogoutButton_Click(sender As Object, e As EventArgs) Handles LogoutButton.Click
        Session.Abandon()
        Server.Transfer("Default.aspx")
    End Sub
End Class