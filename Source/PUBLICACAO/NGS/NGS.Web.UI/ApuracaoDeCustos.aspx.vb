Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ApuracaoDeCustos
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Custos)
        If Not IsPostBack And IsConnect Then
            If Not Funcoes.VerificaPermissao("ApuracaoDeCustos", "ACESSAR") Then
                Dim url As String = IIf(Not String.IsNullOrEmpty(Session("ssMenuDeAcesso")), String.Format("~/{0}.aspx", Session("ssMenuDeAcesso")), "~/Index.aspx")
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", url)
                Exit Sub
            End If
        End If
    End Sub

End Class