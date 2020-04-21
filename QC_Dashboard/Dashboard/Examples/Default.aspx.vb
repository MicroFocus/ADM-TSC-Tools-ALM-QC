Public Class _Default1
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("connected") Is Nothing Or Session("connected") = False Then
            Server.Transfer("~/Default.aspx")
        End If
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Server.Transfer("~/Examples/" & RadioButtonList1.SelectedValue)
    End Sub
End Class