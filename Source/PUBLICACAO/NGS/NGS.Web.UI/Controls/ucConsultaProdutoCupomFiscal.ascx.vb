Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaProdutoCupomFiscal
    Inherits BaseUserControl

    Private Where As String = String.Empty

#Region "Variáveis Locais"
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal
    Private objProdutoNota As [Lib].Negocio.NotaFiscalXItem
#End Region

#Region "Sessão"
    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Protected Sub lnkProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkProduto.Click
        Try
            Session("Where" & HID.Value) = ""
            BuscarProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultaCodBarras_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkConsultaCodBarras.Click
        SessaoRecuperaNotaFiscal()

        Try
            Dim objProduto As New [Lib].Negocio.Produto(LeitorCodBarras())

            If objNotaFiscal.Itens.Any(Function(s) s.CodigoProduto = objProduto.Codigo) Then
                MsgBox(Me.Page, "Produto já adicionado ao Cupom Fiscal", eTitulo.Erro)
                Exit Sub
            End If

            objNotaFiscal.CarregandoNota = True
            objNotaFiscal.CodigoOperacao = 21
            objNotaFiscal.CodigoSubOperacao = 35

            objProdutoNota = New NotaFiscalXItem(objNotaFiscal)

            objProdutoNota.CodigoProduto = objProduto.Codigo
            objProdutoNota.CodigoOperacao = 21
            objProdutoNota.CodigoSubOperacao = 35

            objProdutoNota.PesoQuantidade = objProduto.PesoQuantidade

            objProdutoNota.NotaFiscal.Cliente = New Cliente()
            objProdutoNota.NotaFiscal.Cliente.CodigoEstado = "MS"
            objProdutoNota.NotaFiscal.Cliente.InscricaoEstadual = "ISENTO"

            objProdutoNota.QuantidadeFiscal = 0
            objProdutoNota.QuantidadeFisica = 0
            objProdutoNota.Unitario = 0
            objProdutoNota.ValorTotal = 0

            Dim SaldoProdutoEstoque As New [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem(objProdutoNota.CodigoProduto)
            SaldoProdutoEstoque.CarregarResumoSaldoEmEstoque(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.SubOperacao)

            If SaldoProdutoEstoque.SaldoFiscal = 0 Then
                MsgBox(Me.Page, "Produto sem estoque disponível", eTitulo.Info)
            Else
                objNotaFiscal.Itens.Add(objProdutoNota)

                SessaoSalvaNotaFiscal()

                CType(Me.Page, CupomFiscal).AtualizarItensNoGrid()
                Limpar()

                Popup.CloseDialog(Me.Page, "divConsultaProdutoCupomFiscal")
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function LeitorCodBarras()
        Where = Session("Where" & HID.Value)
        Where = IIf(Where Is Nothing OrElse Where.Length = 0, " Grupo = 30102", Where)
        Dim ds As New DataSet
        Dim Sql As String
        Sql = "SELECT Produto_Id " &
              "  FROM Produtos   " &
              " WHERE Situacao = 1 AND CodBarras = '" & txtCodBarras.Text & "'"

        ds = Banco.ConsultaDataSet(Sql, "Produtos")

        Dim dr As DataRow = ds.Tables(0).Rows(0)

        Return dr("Produto_Id")
    End Function

    Public Sub BuscarProduto()
        Where = Session("Where" & HID.Value)
        Where = IIf(Where Is Nothing OrElse Where.Length = 0, " Grupo = 30102", Where)
        Dim ds As New DataSet
        Dim Sql As String
        Sql = "SELECT top(50) Produto_Id, " &
              " case len(DescricaoMapa) when 0 " &
              "     then Nome " &
              " else Nome + ' - ' + DescricaoMapa " &
              " end AS Nome " &
              "  FROM Produtos " &
              " WHERE Situacao = 1 AND Nome LIKE '%" & txtNome.Text & "%' "

        If Where.Trim.Length > 0 Then
            Sql &= "And " & Where
        End If
        Sql &= " ORDER BY Nome"
        ds = Banco.ConsultaDataSet(Sql, "Produtos")
        gridProduto.DataSource = ds
        gridProduto.DataBind()
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub


    Public Sub InicializarUC()
        Limpar()
    End Sub

    Public Overrides Sub Limpar()
        gridProduto.DataSource = New List(Of Object)
        gridProduto.DataBind()
        txtNome.Text = String.Empty
        txtNome.Focus()
        txtCodBarras.Text = String.Empty
    End Sub

    Protected Overrides Sub Selecionar()
        SessaoRecuperaNotaFiscal()

        Try
            Dim objProduto As New [Lib].Negocio.Produto(gridProduto.SelectedRow.Cells(1).Text())

            If objNotaFiscal.Itens.Any(Function(s) s.CodigoProduto = objProduto.Codigo) Then
                MsgBox(Me.Page, "Produto já adicionado ao Cupom Fiscal", eTitulo.Erro)
                Exit Sub
            End If

            objNotaFiscal.CarregandoNota = True
            objNotaFiscal.CodigoOperacao = 21
            objNotaFiscal.CodigoSubOperacao = 35

            objProdutoNota = New NotaFiscalXItem(objNotaFiscal)

            objProdutoNota.CodigoProduto = objProduto.Codigo
            objProdutoNota.CodigoOperacao = 21
            objProdutoNota.CodigoSubOperacao = 35

            objProdutoNota.PesoQuantidade = objProduto.PesoQuantidade

            objProdutoNota.NotaFiscal.Cliente = New Cliente()
            objProdutoNota.NotaFiscal.Cliente.CodigoEstado = "MS"
            objProdutoNota.NotaFiscal.Cliente.InscricaoEstadual = "ISENTO"

            objProdutoNota.QuantidadeFiscal = 0
            objProdutoNota.QuantidadeFisica = 0
            objProdutoNota.Unitario = 0
            objProdutoNota.ValorTotal = 0

            Dim SaldoProdutoEstoque As New [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem(objProdutoNota.CodigoProduto)
            SaldoProdutoEstoque.CarregarResumoSaldoEmEstoque(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.SubOperacao)

            If SaldoProdutoEstoque.SaldoFiscal = 0 Then
                MsgBox(Me.Page, "Produto sem estoque disponível", eTitulo.Info)
            Else
                objNotaFiscal.Itens.Add(objProdutoNota)

                SessaoSalvaNotaFiscal()

                CType(Me.Page, CupomFiscal).AtualizarItensNoGrid()
                Limpar()

                Popup.CloseDialog(Me.Page, "divConsultaProdutoCupomFiscal")
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Fechar()
        Try
            Dim objProduto As New [Lib].Negocio.Produto()
            CType(Me.Page, IBasePage).Carregar(objProduto)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub gridProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Selecionar()
    End Sub

    Protected Sub txtNome_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        BuscarProduto()
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaProdutoCupomFiscal")
    End Sub

End Class