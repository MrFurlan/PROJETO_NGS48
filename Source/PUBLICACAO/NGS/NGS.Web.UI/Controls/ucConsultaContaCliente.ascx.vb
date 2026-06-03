Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaContaCliente
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Protected Overrides Sub Selecionar()
        Dim str As String = gdvConta.SelectedRow.Cells(1).Text & " - " & gdvConta.SelectedRow.Cells(2).Text
        DirectCast(Me.Page, IBasePage).Carregar(str)
        Popup.CloseDialog(Me.Page, "divConsultaContaClientes")
    End Sub

    Public Sub BindGridView(ByVal TipoConta As String)
        Dim ds As New DataSet
        Dim sql As String = ""

        sql = "SELECT Conta_Id, Titulo FROM PlanoDeContas "
        sql &= "WHERE "
        sql &= "Empresa_Id = '99999999999999' AND "
        sql &= "EndEmpresa_Id = 0"

        If TipoConta = "CLIENTE" Then
            sql &= " AND LEN(Conta_Id) = 7 and Cliente = 'S'"
        Else
            sql &= " AND LEN(Conta_Id) = 9"
        End If

        ds = Banco.ConsultaDataSet(sql, "Consulta")

        gdvConta.DataSource = ds
        gdvConta.DataBind()
    End Sub

    Protected Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaContaClientes")
    End Sub

    Protected Sub gdvConta_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gdvConta.SelectedIndexChanged
        Selecionar()
    End Sub

End Class