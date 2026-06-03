Imports NGS.Lib.Negocio

Public Class Impdados
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack AndAlso IsConnect Then
            '    If Not Funcoes.VerificaPermissao("Impdados", "ACESSAR") Then
            'MsgBox(Me.Page, "Usuário sem permissão para acerssar essa página", "~/Gestao.aspx")
            'Exit Sub
            '    End If
        End If
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Impdados")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class