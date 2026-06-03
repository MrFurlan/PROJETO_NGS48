Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Diagnostics.Eventing.Reader

Public Class ucConsultaLote
    Inherits BaseUserControl

    Private objProduto As [Lib].Negocio.Produto

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub CarregarLote(ByRef codigoEmpresa As String, ByRef endCodigoEmresa As Integer, ByRef pItem As NotaFiscalXItem)
        Try
            Session("objItemProdutoLote" & HID.Value) = pItem

            txtTotalProduto.Text = pItem.QuantidadeFiscal.ToString

            Dim pLote = New OrdemParaProducaoXConsumo()
            Dim ds As New DataSet
            ds = pLote.buscarLoteDeFornecedor(codigoEmpresa, endCodigoEmresa, "'" & CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).CodigoProduto & "'")

            For Each dr In ds.Tables(0).Rows

                If CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Count = 0 Then
                    Dim drItem As DataRow = CType(Session("objLoteFornecedor" & HID.Value), DataTable).NewRow()
                    drItem("Produto") = dr("Produto")
                    drItem("Lote") = dr("Lote")
                    drItem("Fabricado") = CDate(dr("Fabricado")).ToString("dd/MM/yyyy")
                    drItem("Validade") = CDate(dr("Validade")).ToString("dd/MM/yyyy")
                    drItem("Quantidade") = dr("Quantidade")
                    drItem("Consumo") = 0
                    CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Add(drItem)
                Else
                    Dim temLote As Boolean = False

                    For Each row In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                        If CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).CodigoProduto = row("Produto") AndAlso dr("Lote") = row("Lote") AndAlso dr("Fabricado") = row("Fabricado") AndAlso dr("Validade") = row("Validade") Then
                            temLote = True
                        End If
                    Next

                    If Not temLote Then
                        Dim drItem As DataRow = CType(Session("objLoteFornecedor" & HID.Value), DataTable).NewRow()
                        drItem("Produto") = dr("Produto")
                        drItem("Lote") = dr("Lote")
                        drItem("Fabricado") = CDate(dr("Fabricado")).ToString("dd/MM/yyyy")
                        drItem("Validade") = CDate(dr("Validade")).ToString("dd/MM/yyyy")
                        drItem("Quantidade") = dr("Quantidade")
                        drItem("Consumo") = 0
                        CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Add(drItem)

                    End If
                End If
            Next

            Dim dv As DataView = New DataView(CType(Session("objLoteFornecedor" & HID.Value), DataTable))

            Dim itemProdutoLote As NotaFiscalXItem = CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem)

            If Not itemProdutoLote Is Nothing Then
                dv.RowFilter = "Produto = " & CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).CodigoProduto
            End If

            'gridLoteDeFornecedor.DataSource = CType(Session("objLoteFornecedor" & HID.Value), DataTable)
            gridLoteDeFornecedor.DataSource = dv.ToTable()

            gridLoteDeFornecedor.DataBind()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Public Sub MostrarLote(ByRef codigoEmpresa As String, ByRef endCodigoEmresa As Integer, ByRef pItem As NotaFiscalXItem)
        Try
            lnkNovo.Parent.Visible = False
            lnkLimpar.Parent.Visible = False

            Session("objItemProdutoLote" & HID.Value) = pItem

            Dim dv As DataView = New DataView(CType(Session("objLoteFornecedor" & HID.Value), DataTable))

            dv.RowFilter = "Produto = " & pItem.CodigoProduto

            gridLoteDeFornecedor.DataSource = dv.ToTable()

            gridLoteDeFornecedor.DataBind()

            Dim i As Integer = 0
            While i < gridLoteDeFornecedor.Rows.Count

                If pItem.IUD = "U" Then
                    CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Enabled = False
                Else
                    CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Enabled = True
                End If

                i += 1
            End While

            Session("objQuantidadeFiscalLote" & HID.Value) = pItem.QuantidadeFiscal

            txtTotalProduto.Text = pItem.QuantidadeFiscal.ToString & " - " & pItem.Produto.Nome

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Public Sub MostrarLotesParaOrdemProducao(ByRef codigoEmpresa As String, ByRef endCodigoEmresa As Integer, ByRef pItem As NotaFiscalXItem)
        Try

            lnkNovo.Parent.Visible = False
            lnkLimpar.Parent.Visible = False

            Session("objItemProdutoLote" & HID.Value) = pItem

            Dim pLote = New OrdemParaProducaoXConsumo()
            Dim ds As New DataSet
            ds = pLote.buscarLoteDeFornecedor(codigoEmpresa, endCodigoEmresa, "'" & CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).CodigoProduto & "'")

            For Each dr In ds.Tables(0).Rows

                If CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Count = 0 Then
                    Dim drItem As DataRow = CType(Session("objLoteFornecedor" & HID.Value), DataTable).NewRow()
                    drItem("Produto") = dr("Produto")
                    drItem("Lote") = dr("Lote")
                    drItem("Fabricado") = CDate(dr("Fabricado")).ToString("dd/MM/yyyy")
                    drItem("Validade") = CDate(dr("Validade")).ToString("dd/MM/yyyy")
                    drItem("Quantidade") = dr("Quantidade")
                    CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Add(drItem)
                Else
                    Dim temLote As Boolean = False

                    For Each row In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                        If CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).CodigoProduto = row("Produto") AndAlso dr("Lote") = row("Lote") AndAlso dr("Fabricado") = row("Fabricado") AndAlso dr("Validade") = row("Validade") Then
                            temLote = True
                        End If
                    Next

                    If Not temLote Then
                        Dim drItem As DataRow = CType(Session("objLoteFornecedor" & HID.Value), DataTable).NewRow()
                        drItem("Produto") = dr("Produto")
                        drItem("Lote") = dr("Lote")
                        drItem("Fabricado") = CDate(dr("Fabricado")).ToString("dd/MM/yyyy")
                        drItem("Validade") = CDate(dr("Validade")).ToString("dd/MM/yyyy")
                        drItem("Quantidade") = dr("Quantidade")
                        CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Add(drItem)

                    End If
                End If
            Next

            Dim dv As DataView = New DataView(CType(Session("objLoteFornecedor" & HID.Value), DataTable))

            Dim itemProdutoLote As NotaFiscalXItem = CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem)

            If Not itemProdutoLote Is Nothing Then
                dv.RowFilter = "Produto = " & CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).CodigoProduto
            End If

            'Dim dtOrdemProducao As DataTable = dv.ToTable()

            'If dtOrdemProducao.Columns.Contains("Consumo") Then
            '    ' Remover a coluna do DataTable
            '    dtOrdemProducao.Columns.Remove("Consumo")
            'End If

            'gridLoteDeFornecedor.DataSource = CType(Session("objLoteFornecedor" & HID.Value), DataTable)
            gridLoteDeFornecedor.DataSource = dv.ToTable()
            gridLoteDeFornecedor.DataBind()

            Dim i As Integer = 0
            While i < gridLoteDeFornecedor.Rows.Count

                CType(gridLoteDeFornecedor.Rows(i).FindControl("imgSelecionaLote"), ImageButton).Visible = True
                CType(gridLoteDeFornecedor.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/select.jpg"
                CType(gridLoteDeFornecedor.Rows(i).FindControl("imgSelecionaLote"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"
                CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Visible = False

                i += 1

            End While

            If gridLoteDeFornecedor.HeaderRow IsNot Nothing Then
                For iCell As Integer = 0 To gridLoteDeFornecedor.HeaderRow.Cells.Count - 1
                    If gridLoteDeFornecedor.HeaderRow.Cells(iCell).Text = "Consumo" Then
                        gridLoteDeFornecedor.HeaderRow.Cells(iCell).Text = ""
                    End If
                Next
            End If

            txtTotalProduto.Text = pItem.QuantidadeFiscal.ToString & " - " & pItem.Produto.Nome

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Public Overrides Sub Limpar()
        Session.Remove("objItemProdutoLote" & HID.Value)

        txtTotalProduto.Text = String.Empty

        gridLoteDeFornecedor.DataSource = Nothing
        gridLoteDeFornecedor.DataBind()
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            Dim totalLote As Decimal

            Dim j As Integer = 0
            Dim i As Integer = 0

            Dim columnLote As Integer = -1 ' Inicializa a variável com um valor inválido
            Dim columnValidade As Integer = -1 ' Inicializa a variável com um valor inválido
            Dim columnSaldo As Integer = -1 ' Inicializa a variável com um valor inválido

            ' Itera pelas colunas da grid
            For iRow As Integer = 0 To gridLoteDeFornecedor.Columns.Count - 1
                If gridLoteDeFornecedor.Columns(iRow).HeaderText = "Lote" Then
                    columnLote = iRow
                ElseIf gridLoteDeFornecedor.Columns(iRow).HeaderText = "Validade" Then
                    columnValidade = iRow
                ElseIf gridLoteDeFornecedor.Columns(iRow).HeaderText = "Saldo" Then
                    columnSaldo = iRow
                End If
            Next

            While i < gridLoteDeFornecedor.Rows.Count
                For Each drItemLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                    If drItemLote("Produto") = CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).CodigoProduto AndAlso
                        drItemLote("Lote") = gridLoteDeFornecedor.Rows(i).Cells(columnLote).Text AndAlso
                        CDate(drItemLote("Validade")) = CDate(gridLoteDeFornecedor.Rows(i).Cells(columnValidade).Text) AndAlso
                        CDec(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text) > 0 Then

                        If CDec(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text) > 0 Then
                            If CDec(gridLoteDeFornecedor.Rows(i).Cells(columnSaldo).Text) < CDec(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text) Then
                                MsgBox(Me.Page, "Quantidade do saldo do Lote " + gridLoteDeFornecedor.Rows(i).Cells(0).Text + " informado é menor que a quantidade informada.")
                                Exit Sub
                            End If

                            drItemLote("Consumo") = CDec(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text)

                            totalLote += drItemLote("Consumo")

                            j += 1
                        End If
                    End If
                Next

                i += 1
            End While

            If totalLote <> CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).QuantidadeFiscal Then
                MsgBox(Me.Page, "Total de Lote informado é diferente do total do Produto da Nota Fiscal.")
                Exit Sub
            End If

            Session(HttpContext.Current.Session("ssTipoRetorno")) = CType(Session("objLoteFornecedor" & HID.Value), DataTable)

            If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                CType(uc, IBaseUserControl).Carregar(CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem))
            Else
                If TypeOf Me.Page Is ControleDeProducao Then
                    If j > 1 Then
                        MsgBox(Me.Page, "Para o Controle de Produção só pode estar marcado 1 lote por lançamento.")
                        Exit Sub
                    End If
                End If

                CType(Me.Page, IBasePage).Carregar(CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem))
            End If

            Popup.CloseDialog(Me.Page, "divConsultaDeLote")

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click

        If Not CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).IUD = "U" Then
            Dim totalLote As Decimal

            Dim j As Integer = 0
            Dim i As Integer = 0

            Dim columnLote As Integer = -1 ' Inicializa a variável com um valor inválido
            Dim columnValidade As Integer = -1 ' Inicializa a variável com um valor inválido
            Dim columnSaldo As Integer = -1 ' Inicializa a variável com um valor inválido

            ' Itera pelas colunas da grid
            For iRow As Integer = 0 To gridLoteDeFornecedor.Columns.Count - 1
                If gridLoteDeFornecedor.Columns(iRow).HeaderText = "Lote" Then
                    columnLote = iRow
                ElseIf gridLoteDeFornecedor.Columns(iRow).HeaderText = "Validade" Then
                    columnValidade = iRow
                ElseIf gridLoteDeFornecedor.Columns(iRow).HeaderText = "Saldo" Then
                    columnSaldo = iRow
                End If
            Next

            While i < gridLoteDeFornecedor.Rows.Count
                For Each drItemLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                    If drItemLote("Produto") = CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).CodigoProduto AndAlso
                        drItemLote("Lote") = gridLoteDeFornecedor.Rows(i).Cells(columnLote).Text AndAlso
                        CDate(drItemLote("Validade")) = CDate(gridLoteDeFornecedor.Rows(i).Cells(columnValidade).Text) Then

                        If CDec(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text) > 0 Then
                            If CDec(gridLoteDeFornecedor.Rows(i).Cells(columnSaldo).Text) < CDec(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text) Then
                                MsgBox(Me.Page, "Quantidade do saldo do Lote " + gridLoteDeFornecedor.Rows(i).Cells(0).Text + " informado é menor que a quantidade informada.")
                                Exit Sub
                            End If

                            drItemLote("Consumo") = CDec(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text)

                            totalLote += drItemLote("Consumo")

                            j += 1
                        Else
                            drItemLote("Consumo") = 0
                        End If
                    End If
                Next

                i += 1
            End While

            If totalLote <> CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).QuantidadeFiscal Then
                MsgBox(Me.Page, "Total de Lote informado é diferente do total do Produto do Cupom Fiscal.")
                Exit Sub
            End If
        End If

        Popup.CloseDialog(Me.Page, "divConsultaDeLote")
    End Sub

    Protected Overrides Sub Selecionar()
        Try

            Session(Session("ssTipoRetorno")) = objProduto
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objProduto)
                Else
                    CType(Me.Page, IBasePage).Carregar(objProduto)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaDeLote")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgSelecionaLote_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Dim imgSelecionaLote As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(imgSelecionaLote.NamingContainer, GridViewRow)

        Session("objLinha" & HID.Value) = row.RowIndex

        objProduto = New [Lib].Negocio.Produto(CType(Session("objItemProdutoLote" & HID.Value), NotaFiscalXItem).CodigoProduto)

        Dim iColumn As Integer = 0
        For Each column As DataControlField In gridLoteDeFornecedor.Columns
            If TypeOf column Is BoundField Then
                Dim boundField As BoundField = CType(column, BoundField)
                If boundField.HeaderText = "Lote" Then
                    objProduto.NumeroDoLote = row.Cells(iColumn).Text
                ElseIf boundField.HeaderText = "Validade" Then
                    objProduto.ValidadeDoLote = row.Cells(iColumn).Text
                End If
            End If
            iColumn += 1
        Next

        Selecionar()

    End Sub
End Class