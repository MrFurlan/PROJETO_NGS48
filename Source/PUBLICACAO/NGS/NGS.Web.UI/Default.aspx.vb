Imports NGS.Lib.Uteis
Imports System.Web.Configuration
Imports System.Net

Public Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (IsPostBack) Then
            Return
        End If
        Redirect()
    End Sub

    Private Sub Redirect()
        If Page.User IsNot Nothing AndAlso Page.User.Identity.IsAuthenticated AndAlso Session("ssNomeUsuario") IsNot Nothing Then
            Response.Redirect("~/Index.aspx")
        Else
            Response.Redirect("~/Login.aspx")
        End If
    End Sub

End Class