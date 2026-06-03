Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaOrdemDeProducao
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            ddl.Carregar(ddlGrupoProdutoConsultaProducao, CarregarDDL.Tabela.GrupoProdutoXConsumo)
            Me.Limpar()
        End If
    End Sub

    Protected Sub ddlGrupoProdutoConsultaProducao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProdutoConsultaProducao.SelectedIndexChanged
        Try
            CarregarProdutoConsultaProducao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarProdutoConsultaProducao()
        ddl.Carregar(ddlProdutosConsultaProducao, CarregarDDL.Tabela.ProdutoProducao, " P.Grupo = '" & ddlGrupoProdutoConsultaProducao.SelectedValue & "'")
    End Sub

    Protected Sub lnkConsultarOrdem_Click(sender As Object, e As EventArgs) Handles lnkConsultarOrdem.Click
        Try
            BuscarOrdem()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub BuscarOrdem()

        Dim strSQL As String = "SELECT op.Ordem_Id AS Codigo, OPXP.Lote, op.Movimento, OPXP.Produto_Id AS Produto, prd.Nome AS NomeProduto, (OPXP.Quantidade + OPXP.QuantidadeDeAjuste) AS Quantidade " & vbCrLf &
                                "FROM OrdemDeProducao op " & vbCrLf &
                                "INNER JOIN OrdemDeProducaoXProduto OPXP " & vbCrLf &
                                "   ON op.Empresa_Id		        = OPXP.Empresa_Id " & vbCrLf &
                                "   AND op.EndEmpresa_Id	        = OPXP.EndEmpresa_Id " & vbCrLf &
                                "   AND op.Ordem_Id			        = OPXP.Ordem_Id " & vbCrLf &
                                "	INNER JOIN Produtos AS prd" & vbCrLf &
                                "			on prd.Produto_Id = OPXP.Produto_Id" & vbCrLf &
                                "WHERE op.Empresa_Id    = '" & Session("ssEmpresa") & "'" & vbCrLf &
                                "  AND op.EndEmpresa_Id = " & Session("ssEndEmpresa") & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtSequencia.Text) AndAlso CInt(txtSequencia.Text) > 0 Then
            strSQL &= "  AND op.Ordem_Id      = " & txtSequencia.Text
        End If

        If Not String.IsNullOrWhiteSpace(txtLote.Text) Then
            strSQL &= "  AND OPXP.Lote          = '" & txtSequencia.Text & "'" & vbCrLf
        End If

        If ddlProdutosConsultaProducao.SelectedIndex > 0 Then
            strSQL &= "  AND OPXP.Produto_Id       = '" & ddlProdutosConsultaProducao.SelectedValue & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            strSQL &= "  AND op.Movimento BETWEEN '" & Format(txtDataInicial.Text, ("yyyy-MM-dd")) & "' AND '" & Format(txtDataInicial.Text, ("yyyy-MM-dd")) & "'" & vbCrLf
        End If

        strSQL &= " ORDER BY op.Movimento"

        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(strSQL, "ÖrdemDeProducao")

        gridOrdemDeProducao.DataSource = ds
        gridOrdemDeProducao.DataBind()
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        txtSequencia.Text = String.Empty
        txtLote.Text = String.Empty
        txtDataInicial.Text = String.Empty
        txtDataFinal.Text = String.Empty

        ddlGrupoProdutoConsultaProducao.SelectedIndex = 0
        ddlProdutosConsultaProducao.Items.Clear()

        gridOrdemDeProducao.DataSource = Nothing
        gridOrdemDeProducao.DataBind()
        Session.Remove("_MainUserControl")
    End Sub

    Protected Sub gridOrdemDeProducao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Selecionar()
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Dim objOP As New [Lib].Negocio.OrdemParaProducao(Session("ssEmpresa"), Session("ssEndEmpresa"), CInt(gridOrdemDeProducao.SelectedRow.Cells(1).Text()))
            Session(Session("ssTipoRetorno")) = objOP
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objOP)
                Else
                    CType(Me.Page, IBasePage).Carregar(objOP)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaOrdemDeProducao")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Fechar()
        Try
            Dim objOP As New [Lib].Negocio.OrdemParaProducao()
            CType(Me.Page, IBasePage).Carregar(objOP)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkFecharOrdem_Click(sender As Object, e As EventArgs) Handles lnkFecharOrdem.Click
        Try
            Fechar()
            Popup.CloseDialog(Me.Page, "divConsultaOrdemDeProducao")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class