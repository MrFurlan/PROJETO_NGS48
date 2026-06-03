Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading
Imports Elmah.ContentSyndication

Public Class ucEnviarXMLEmissao
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Protected Sub btnEnviar_Click(sender As Object, e As EventArgs) Handles btnEnviar.Click
        If rdNotaFiscal.Checked Then
            'HttpContext.Current.Response.Redirect("NotaFiscalXItens.aspx")
            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('NotaFiscalXItens.aspx');", True)
        ElseIf rdNotaGeral.Checked Then
            'HttpContext.Current.Response.Redirect("NotasFiscaisGerais.aspx")
            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('NotasFiscaisGerais.aspx');", True)
        ElseIf rdConhecimentoTransporte.Checked Then
            'HttpContext.Current.Response.Redirect("NotasFiscaisGerais.aspx")
            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('ConhecimentoDeTransporte.aspx');", True)
        End If

        Popup.CloseDialog(Me.Page, "divEnviarXMLEmissao")

    End Sub

    Protected Sub bntCarregar_Click(sender As Object, e As EventArgs)
        If rdNotaFiscal.Checked Then
            HttpContext.Current.Response.Redirect("NotaFiscalXItens.aspx")
        ElseIf rdNotaGeral.Checked Then
            HttpContext.Current.Response.Redirect("NotasFiscaisGerais.aspx")
        ElseIf rdConhecimentoTransporte.Checked Then
            HttpContext.Current.Response.Redirect("ConhecimentoDeTransporte.aspx")
        End If
    End Sub

    Public Overrides Sub limpar()
        rdNotaGeral.Checked = False
        rdNotaFiscal.Checked = False
        rdConhecimentoTransporte.Checked = False
        rdConhecimentoTransporte.Enabled = False

        Dim ModeloNFe As String = Mid(Session("chaveXMLautomacao"), 21, 2)

        If ModeloNFe.Equals("57") AndAlso (Left(Session("ssEmpresa"), 8) = "53267147" OrElse
                                            Left(Session("ssEmpresa"), 8) = "62747840" OrElse
                                            Left(Session("ssEmpresa"), 8) = "62780383" OrElse
                                            Left(Session("ssEmpresa"), 8) = "63358210") Then
            rdConhecimentoTransporte.Enabled = True
            rdConhecimentoTransporte.Checked = True
            rdNotaGeral.Enabled = False
            rdNotaFiscal.Enabled = False
        End If

    End Sub

    Protected Sub lnkFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divEnviarXMLEmissao")
    End Sub

End Class