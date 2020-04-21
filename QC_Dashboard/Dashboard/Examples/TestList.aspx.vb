Public Class TestList
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("connected") Is Nothing Or Session("connected") = False Then
            Server.Transfer("~/Default.aspx")
        End If

    End Sub
    Protected Sub RunButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RunButton.Click
        On Error Resume Next
        Dim testobj
        Dim testFact
        Dim mydatatable As DataTable
        Dim myNewRow As DataRow
        Dim testFilter

        mydatatable = New DataTable("Tests")

        mydatatable.Columns.Add("Test ID", System.Type.GetType("System.String"))
        mydatatable.Columns.Add("Execution Status", System.Type.GetType("System.String"))
        mydatatable.Columns.Add("Created By", System.Type.GetType("System.String"))
        mydatatable.Columns.Add("Test Name", System.Type.GetType("System.String"))
        mydatatable.Columns.Add("Test Type", System.Type.GetType("System.String"))

        'Get bugfactory
        Session("tdc").InitConnectionEx(Session("serverurl"))
        Session("tdc").Login(Session("username"), Session("password"))
        Session("tdc").Connect(Session("domain"), Session("Project"))

        TestFact = Session("tdc").TestFactory
        testFilter = testFact.Filter
        testFilter.order("TS_TEST_ID") = 1
        'Get a bug object.

        For Each testobj In testFact.newlist(testFilter.text)

            myNewRow = mydatatable.NewRow()
            myNewRow("Test ID") = testobj.Field("TS_TEST_ID")
            myNewRow("Execution Status") = testobj.Field("TS_EXEC_STATUS")
            myNewRow("Created By") = testobj.Field("TS_RESPONSIBLE")
            myNewRow("Test Name") = testobj.Field("TS_NAME")
            myNewRow("Test Type") = testobj.Field("TS_TYPE")
            mydatatable.Rows.Add(myNewRow)
        Next


        If Err.Number <> 0 Then
            LogText.Text = Err.Description
        Else
            LogText.Text = "Tests list generated successfully."
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