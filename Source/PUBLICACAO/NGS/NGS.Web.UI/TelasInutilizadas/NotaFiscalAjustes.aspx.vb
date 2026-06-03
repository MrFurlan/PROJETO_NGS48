Imports System.Data
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class NotaFiscalAjustes
    Inherits BasePage

    Dim Sql As String
    Dim Sqla As String

    Dim SqlArray As New ArrayList
    Dim Empresa() As String
    Dim Cliente() As String
    Dim SequenciaLote As Integer = 0
    Dim ds As New DataSet

    Dim Condicao As String = ""
    Dim Codigo As String = ""
    Dim Endereco As Integer = 0

    Dim Mensagem As String = ""

    Dim xPop As String
    Dim yPop As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack AndAlso IsConnect Then
            txtPesoFiscal.Text = HttpContext.Current.Session("txtPesoFiscal")
            txtUnitario.Text = HttpContext.Current.Session("txtUnitarioOficial")
            txtValorTotal.Text = HttpContext.Current.Session("txtValorTotal")

            txtBaseIcms.Text = Replace(HttpContext.Current.Session("txtBaseDeCalculo"), ".", "")
            txtValorIcms.Text = Replace(HttpContext.Current.Session("txtValorDeIcms"), ".", "")
        End If
    End Sub

    Protected Sub cmdLimpar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Limpar()
    End Sub

    Sub Limpar()
        txtPesoFiscal.Text = HttpContext.Current.Session("txtPesoFiscal")
        txtUnitario.Text = HttpContext.Current.Session("txtUnitarioOficial")
        txtValorTotal.Text = HttpContext.Current.Session("txtValorTotal")

        txtBaseIcms.Text = HttpContext.Current.Session("txtBaseDeCalculo")
        txtValorIcms.Text = HttpContext.Current.Session("txtValorDeIcms")
    End Sub

    Protected Sub cmdSair_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        HttpContext.Current.Session("txtPesoFiscal") = Replace(txtPesoFiscal.Text, ".", "")
        HttpContext.Current.Session("txtUnitarioOficial") = Replace(txtUnitario.Text, ".", "")
        HttpContext.Current.Session("txtValorTotal") = Replace(txtValorTotal.Text, ".", "")
        HttpContext.Current.Session("txtBaseDeCalculo") = Replace(txtBaseIcms.Text, ".", "")
        HttpContext.Current.Session("txtValorDeIcms") = Replace(txtValorIcms.Text, ".", "")
        HttpContext.Current.Session("ssRevisao") = "S"

        Dim jscript As String
        jscript = "<script language='javascript' type='text/javascript'>"
        jscript += ";self.close();"
        jscript += "</script>"
        ScriptManager.RegisterStartupScript(Me, Me.UpdatePanel1.GetType(), Guid.NewGuid().ToString(), jscript, False)
    End Sub

    Sub Calcular()
        txtValorTotal.Text = (CInt(txtPesoFiscal.Text) * CDec(txtUnitario.Text)) / CInt(HttpContext.Current.Session("txtBaseDeCalculoDoProduto"))
        txtValorTotal.Text = FormatNumber(txtValorTotal.Text, 2)
    End Sub

    Protected Sub cmdCalcular_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Calcular()
    End Sub

End Class