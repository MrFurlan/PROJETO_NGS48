Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Requisicoes
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            'If Funcoes.VerificaPermissao("Requisicoes", "ACESSAR") Then

            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "")
            '    Exit Sub
            'End If
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click

    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click

    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click

    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click

    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click

    End Sub

    Protected Sub grdRequisicao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub grdConsulta_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub lnkAdicionar_Click(sender As Object, e As EventArgs) Handles lnkAdicionar.Click
        ucProdutoNFG.InicializarUC(-1, False, False)
        Popup.ConsultaProdutoNFG(Me.Page, "objProdutoNotaNF" & HID.Value)
    End Sub

End Class