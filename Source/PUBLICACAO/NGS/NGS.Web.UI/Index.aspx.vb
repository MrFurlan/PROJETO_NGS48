Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Index
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            pageLoad()
        End If
    End Sub

    Private Sub pageLoad()
        Me.setMenu(eModulo.Gestao)
        If ((Session("ssNomeUsuario") Is Nothing) OrElse (IsNothing(Request.Cookies("conexao")) AndAlso IsNothing(Session("Conexao")))) Then
            Response.Redirect("~/Logout.aspx")
            Exit Sub
        Else
            Session("Conexao") = FuncoesStrings.DecodificarDe64Bits(Request.Cookies("conexao").Value)
        End If
    End Sub

End Class