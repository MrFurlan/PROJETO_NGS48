Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucLaudoManual
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            ddlGrupoProdutoProducao.Items.Clear()
            ddlEspecificacao.Items.Clear()
            CarregarGrupoDeProdutos()
            ddl.Carregar(ddlEspecificacao, CarregarDDL.Tabela.EspecificacaoDoProduto, "Ativo = 1", True)
            Me.Limpar()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Private Function ValidarCampos() As Boolean
        If ddlProdutos.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Produto nào foi selecinado.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtNLote.Text) Then
            MsgBox(Me.Page, "Número do Lote não informado.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtData.Text) Then
            MsgBox(Me.Page, "Data deve ser Informada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtVencimento.Text) Then
            MsgBox(Me.Page, "Vencimento deve ser Informado.")
            Return False
        Else
            CType(Session("objLoteFornecedor" & HID.Value), DataTable).Clear()
            Dim drItem As DataRow = CType(Session("objLoteFornecedor" & HID.Value), DataTable).NewRow()

            drItem("Produto") = ddlProdutos.SelectedValue
            drItem("Lote") = txtNLote.Text
            drItem("Movimento") = CDate(txtData.Text).ToString("dd/MM/yyyy")
            drItem("Validade") = CDate(txtVencimento.Text).ToString("dd/MM/yyyy")
            drItem("Quantidade") = 0
            drItem("Consumo") = 0

            CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Add(drItem)

            For Each gridRow In gridEspecificacaoDoProduto.Rows
                Dim referencias As String = CType(gridEspecificacaoDoProduto.Rows(gridRow.RowIndex).FindControl("txtReferencia"), TextBox).Text
                Dim resultado As String = CType(gridEspecificacaoDoProduto.Rows(gridRow.RowIndex).FindControl("txtResultado"), TextBox).Text

                CType(Session("objEspecificacao" & HID.Value), DataTable).Rows(gridRow.RowIndex).Item(2) = referencias
                CType(Session("objEspecificacao" & HID.Value), DataTable).Rows(gridRow.RowIndex).Item(3) = resultado
            Next

            Return True
        End If
    End Function

    Private Sub CarregarGrupoDeProdutos()
        ddl.Carregar(ddlGrupoProdutoProducao, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub CarregarProdutos()
        ddl.Carregar(ddlProdutos, CarregarDDL.Tabela.ProdutoPorGrupo, ddlGrupoProdutoProducao.SelectedValue, True)
    End Sub

    Private Sub CarregarEspecificacoes()
        ddlEspecificacao.Enabled = True
        txtFaixaInicial.Enabled = True
        txtFaixaFinal.Enabled = True

        LimparSession()

        Dim prd As Produto = New Produto(ddlProdutos.SelectedValue)

        For Each eP In prd.ProdutoXEspecificacao.Where(Function(s) s.Ativo = True)
            Dim drEP As DataRow = CType(Session("objEspecificacao" & HID.Value), DataTable).NewRow()

            drEP("Codigo") = eP.CodigoEspecificacao
            drEP("Descricao") = eP.EspecificacaoDoProduto.Descricao
            drEP("Referencia") = "Min. " & eP.FaixaInicial & " - Max. " & eP.FaixaFinal
            drEP("ResultadoTxt") = ""

            CType(Session("objEspecificacao" & HID.Value), DataTable).Rows.Add(drEP)
        Next

        If CType(Session("objEspecificacao" & HID.Value), DataTable) IsNot Nothing AndAlso CType(Session("objEspecificacao" & HID.Value), DataTable).Rows.Count > 0 Then
            gridEspecificacaoDoProduto.DataSource = CType(Session("objEspecificacao" & HID.Value), DataTable)
        Else
            gridEspecificacaoDoProduto.DataSource = Nothing
        End If

        gridEspecificacaoDoProduto.DataBind()
    End Sub

    'Public Overrides Sub Carregar(obj As IBaseEntity)
    '    If Session("objProdutoXLDM" & HID.Value) IsNot Nothing Then
    '        Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

    '        ddlGrupoProdutoProducao.SelectedValue = objProduto.CodigoGrupo
    '        ddlProdutos.SelectedValue = objProduto.Codigo

    '        CarregarEspecificacoes()
    '    End If
    'End Sub

    Private Sub Confirmar()
        CType(Me.Page, OrdemDeProducao).ImprimirLaudoManual()
        Popup.CloseDialog(Me.Page, "divLaudoManual")
    End Sub

    Private Sub LimparSession()
        Session.Remove("objEspecificacao" & HID.Value)

        Dim dtEspecificacao As New DataTable("Item")
        dtEspecificacao.Columns.Add("Codigo", Type.GetType("System.String"))
        dtEspecificacao.Columns.Add("Descricao", Type.GetType("System.String"))
        dtEspecificacao.Columns.Add("Referencia", Type.GetType("System.String"))
        dtEspecificacao.Columns.Add("ResultadoTxt", Type.GetType("System.String"))
        Session("objEspecificacao" & HID.Value) = dtEspecificacao
    End Sub

    Public Overrides Sub Limpar()
        ddlGrupoProdutoProducao.SelectedIndex = 0
        ddlProdutos.Items.Clear()
        ddlEspecificacao.SelectedIndex = 0

        gridEspecificacaoDoProduto.DataSource = Nothing
        gridEspecificacaoDoProduto.DataBind()

        LimparSession()

        txtNLote.Text = String.Empty
        txtData.Text = String.Empty
        txtVencimento.Text = String.Empty

        txtFaixaFinal.Text = String.Empty
        txtFaixaInicial.Text = String.Empty
    End Sub

    Protected Sub lnkBuscaProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProduto.Click
        Try
            Dim ucConsultaProduto = CType(Me.Page.FindControlRecursive("ucConsultaProduto"), ucConsultaProduto)

            If ucConsultaProduto IsNot Nothing Then
                ucConsultaProduto.Limpar()
                ucConsultaProduto.SetarHID(HID.Value)
                ucConsultaProduto.MainUserControl = Me
                Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
                Popup.ConsultaDeProduto(Me.Page, "objProdutoXLDM" & HID.Value, txtNome.ClientID, True)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoProduto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProdutoProducao.SelectedIndexChanged
        Try
            CarregarProdutos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlProdutos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProdutos.SelectedIndexChanged
        CarregarEspecificacoes()
    End Sub

    Protected Sub lnkNovaEspecificacao_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovaEspecificacao.Click
        Try
            Dim drEP As DataRow = CType(Session("objEspecificacao" & HID.Value), DataTable).NewRow()

            drEP("Codigo") = ddlEspecificacao.SelectedValue
            drEP("Descricao") = ddlEspecificacao.SelectedItem
            drEP("Referencia") = "Min. " & txtFaixaInicial.Text & " - Max. " & txtFaixaFinal.Text
            drEP("ResultadoTxt") = ""

            CType(Session("objEspecificacao" & HID.Value), DataTable).Rows.Add(drEP)

            If CType(Session("objEspecificacao" & HID.Value), DataTable) IsNot Nothing AndAlso CType(Session("objEspecificacao" & HID.Value), DataTable).Rows.Count > 0 Then
                gridEspecificacaoDoProduto.DataSource = CType(Session("objEspecificacao" & HID.Value), DataTable)
            Else
                gridEspecificacaoDoProduto.DataSource = Nothing
            End If

            gridEspecificacaoDoProduto.DataBind()

            ddlEspecificacao.SelectedIndex = 0
            txtFaixaInicial.Text = ""
            txtFaixaFinal.Text = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRemoverEspecificacao_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgRemover As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgRemover.NamingContainer, GridViewRow)

            For Each item In CType(Session("objEspecificacao" & HID.Value), DataTable).Rows
                If item("Codigo") = row.Cells(0).Text Then
                    CType(Session("objEspecificacao" & HID.Value), DataTable).Rows.Remove(item)
                    Exit For
                End If
            Next

            If CType(Session("objEspecificacao" & HID.Value), DataTable) IsNot Nothing AndAlso CType(Session("objEspecificacao" & HID.Value), DataTable).Rows.Count > 0 Then
                gridEspecificacaoDoProduto.DataSource = CType(Session("objEspecificacao" & HID.Value), DataTable)
            Else
                gridEspecificacaoDoProduto.DataSource = Nothing
            End If

            gridEspecificacaoDoProduto.DataBind()


        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConfirmar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConfirmar.Click
        Try
            If ValidarCampos() Then Confirmar()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnklimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divLaudoManual")
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "LaudoManual")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class