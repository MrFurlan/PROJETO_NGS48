Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Financeiro
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Not Funcoes.VerificaPermissao("Financeiro", "ACESSAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para acessar esse módulo!", IIf(Not String.IsNullOrEmpty(Session("ssMenuDeAcesso")), Session("ssMenuDeAcesso") & ".aspx", "~/Index.aspx"), eTitulo.Info)
                Exit Sub
            End If
        End If
    End Sub

End Class
