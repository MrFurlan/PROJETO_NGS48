Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Logout
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Session.Abandon()
            Response.Redirect(IIf(Request.QueryString("url") Is Nothing, "~/Login.aspx", String.Format("~/Login.aspx?url={0}", Request.QueryString("url"))))
        End If
    End Sub

End Class