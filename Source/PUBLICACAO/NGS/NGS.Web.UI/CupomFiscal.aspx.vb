Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO

Public Class CupomFiscal
    Inherits BasePage

#Region "Retorno UC"
    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objCliente" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            Dim objCliente As [Lib].Negocio.Cliente = Session("objCliente" & HID.Value)

            If objCliente.Codigo.Length = 11 Then
                hdfCodigoCliente.Value = objCliente.CodigoEndereco
                txtCelular.Text = objCliente.Telefone
                txtCPF.Text = objCliente.Codigo
                txtNovoCliente.Text = objCliente.Nome
            Else
                MsgBox(Me.Page, "Só pode ser utilizado Cliente Pessoa Física - CPF!", eTitulo.Info)
            End If

            Session.Remove("objCliente" & HID.Value)
        End If

    End Sub

    Public Sub CarregarNFCe(ByVal obj As [Lib].Negocio.IBaseEntity)

        LimparCampos()

        Dim objNFConsulta As [Lib].Negocio.NotaFiscal = obj
        objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNFConsulta)

        ddlUnidadeNegocio.SelectedValue = objNotaFiscal.CodigoUnidadeDeNegocio
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue, True)
        ddlEmpresa.SelectedValue = objNotaFiscal.CodigoEmpresa & "-" & objNotaFiscal.EnderecoEmpresa
        hdfCodigoCliente.Value = objNotaFiscal.CodigoCliente & "-" & objNotaFiscal.EnderecoCliente

        If objNotaFiscal.Cliente IsNot Nothing AndAlso objNotaFiscal.CodigoCliente <> objNotaFiscal.CodigoEmpresa Then
            txtNovoCliente.Text = objNotaFiscal.Cliente.Nome
            txtCPF.Text = objNotaFiscal.Cliente.Codigo
            txtCelular.Text = objNotaFiscal.Cliente.Telefone
        End If

        ddlSituacao.SelectedValue = objNotaFiscal.CodigoSituacao
        ddlSituacao.Enabled = False

        txtDataNota.Text = objNotaFiscal.DataNota.ToShortDateString()

        txtAvisoCupom.Text = "Cupom: " & objNotaFiscal.Codigo

        CarregarItensNaGrid(objNotaFiscal)

        gridEncargosGerais.DataSource = From nfEnc In objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").SelectMany(Function(s) s.Encargos)
                                        Group By nfEnc.Codigo Into ValorOficial = Sum(nfEnc.Valor)
                                        Order By IIf(Codigo = "PRODUTO", 1, IIf(Codigo = "LIQUIDO", 3, 2))
                                        Select Codigo, ValorOficial
        gridEncargosGerais.DataBind()

        lnkAdicionarItem.Parent.Parent.Visible = False

        If objNotaFiscal.Itens IsNot Nothing AndAlso objNotaFiscal.Itens.Count > 0 Then

            If objNotaFiscal.CodigoOperacao > 0 AndAlso objNotaFiscal.SubOperacao.Financeiro Then
                TabVencimentosold.Visible = True
            Else
                TabVencimentosold.Visible = False
            End If

            objNotaFiscal.AtualizaTotais()

            If objNotaFiscal.SubOperacao IsNot Nothing AndAlso objNotaFiscal.SubOperacao.Financeiro Then

                Dim TotalPago As Decimal = objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoProvisao = 1).Sum(Function(s) s.ValorDoDocumento)
                If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso
                    objNotaFiscal.VencimentosNota.Count > 0 Then

                    grdCondicoes.DataSource = objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D")
                    grdCondicoes.DataBind()

                    For x As Integer = 0 To objNotaFiscal.VencimentosNota.Count - 1
                        If objNotaFiscal.VencimentosNota(x).CodigoProvisao = 1 Then
                            grdCondicoes.Rows(x).Cells(0).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(1).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(2).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(3).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(4).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(5).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(6).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(7).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(8).ForeColor = Drawing.Color.Red
                        End If
                    Next
                End If

                txtTotalNota.Text = objNotaFiscal.TotalNota.ToString("N2")
            End If
        End If

        'Toolbar
        lnkNovo.Parent.Visible = False
        lnkConsultar.Parent.Visible = False

        If objNotaFiscal.CodigoSituacao = 1 Then
            lnkCancelar.Parent.Visible = True
        ElseIf objNotaFiscal.CodigoSituacao = 4 Then
            lnkReenviar.Parent.Visible = True
        End If

        If Session("ssNomeUsuario") = "FURLAN" OrElse Session("ssNomeUsuario") = "KAYNÃ" OrElse Session("ssNomeUsuario") = "FELIPE" Then lnkVisualizar.Parent.Visible = True
        If Session("ssNomeUsuario") = "FURLAN" OrElse Session("ssNomeUsuario") = "KAYNÃ" OrElse Session("ssNomeUsuario") = "FELIPE" Then btnContabil.Visible = True

        If objNotaFiscal.SubOperacao.Contabil Then
            TabContabil.Visible = True
            ddlProdutoContabilizacao.Items.Add(New ListItem("Todos os Lançamentos", "0"))
            For Each row In objNotaFiscal.Itens
                ddlProdutoContabilizacao.Items.Add(New ListItem("Produto " & row.CodigoProduto & " - " & row.Produto.Nome, row.CodigoProduto))
            Next

            objNotaFiscal.LancamentosContabeis.CalcularSaldo()
            gridRazao.DataSource = objNotaFiscal.LancamentosContabeis.OrderBy(Function(s) s.Sequencia)
            gridRazao.DataBind()

            lblDebito.Text = objNotaFiscal.LancamentosContabeis.Sum(Function(s) s.DebitoOficial).ToString("N2")
            lblCredito.Text = objNotaFiscal.LancamentosContabeis.Sum(Function(s) s.CreditoOficial).ToString("N2")
        Else
            TabContabil.Visible = False
        End If

        'Caregar Lote
        For Each item In objNotaFiscal.Itens

            item.IUD = "U"

            For Each nLote In item.Lotes

                Dim drItem As DataRow = CType(Session("objLoteFornecedor" & HID.Value), DataTable).NewRow()
                drItem("Produto") = item.CodigoProduto
                drItem("Lote") = nLote.Lote
                drItem("Fabricado") = nLote.Fabricado.ToString("dd/MM/yyyy")
                drItem("Validade") = nLote.Validade.ToString("dd/MM/yyyy")
                drItem("Quantidade") = 0
                drItem("Consumo") = nLote.Quantidade

                CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Add(drItem)
            Next
        Next

        Dim i As Integer = 0
        While i < grdProdutos.Rows.Count
            CType(grdProdutos.Rows(i).FindControl("imgExcluir"), ImageButton).Visible = False
            CType(grdProdutos.Rows(i).FindControl("txtQuantidadeFiscalGrid"), TextBox).Enabled = False
            CType(grdProdutos.Rows(i).FindControl("txtUnitarioGrid"), TextBox).Enabled = False
            i += 1
        End While

        objNotaFiscal.CarregandoNota = True

        SessaoSalvaNotaFiscal()

    End Sub

    'Carrega a NFe apartir do user control de arquivo.
    Public Overrides Sub Carregar(ByVal pString As String)

    End Sub


    Public Sub CarregarItensNaGrid(ByVal nf As NotaFiscal)
        Dim dtItens As New DataTable("Itens")

        dtItens.Columns.Add("CodigoProduto", Type.GetType("System.String"))
        dtItens.Columns.Add("Produto", Type.GetType("System.String"))
        dtItens.Columns.Add("QuantidadeFiscal", Type.GetType("System.Int64"))
        dtItens.Columns.Add("Unitario", Type.GetType("System.Decimal"))
        dtItens.Columns.Add("ValorTotal", Type.GetType("System.Decimal"))

        For Each row As [Lib].Negocio.NotaFiscalXItem In nf.Itens
            Dim drItem As DataRow = dtItens.NewRow()

            drItem("CodigoProduto") = row.CodigoProduto
            drItem("Produto") = row.Produto.Nome
            drItem("Unitario") = row.Unitario.ToString("N2")
            drItem("QuantidadeFiscal") = CInt(row.QuantidadeFisica)
            drItem("ValorTotal") = row.ValorTotal.ToString("N2")
            dtItens.Rows.Add(drItem)
        Next

        grdProdutos.DataSource = dtItens 'objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").ToList.OrderBy(Function(s) s.Sequencia)
        grdProdutos.DataBind()
    End Sub

#End Region

#Region "Variáveis Locais"
    Private objNotaFiscal As NotaFiscal
    Private Sql As String
    Private mensagemErro As String
#End Region

#Region "Inicialização"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If IsConnect AndAlso Not IsPostBack Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Expedicao.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("CupomFiscal", "ACESSAR") Then
                ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                ddl.Carregar(ddlSituacao, CarregarDDL.Tabela.Situacao, "")
                ddl.Carregar(ddlCondicaoDePagamento, CarregarDDL.Tabela.CondicaoDePagamento, "", True)
                ddl.Carregar(cmbCarteira, CarregarDDL.Tabela.CarteiraFinanceira, "Adiantamento = 'N' AND Produto_Id NOT IN(SELECT Carteira_Id from CarteirasXTributos where Carteira_Id = Produto_Id) ", True)

                Dim Parametros As New Hashtable
                Parametros.Clear()
                Parametros.Add("listarTudo", "N")

                ddl.Carregar(cmbFormas, CarregarDDL.Tabela.TipoDePagamento, "", True, Parametros)

                VerificaUnidade()

                LimparCampos()

            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Sub LimparCampos()
        lblUsuario.Text = Session("ssNomeUsuario")

        Session.Remove("objLoteFornecedor" & HID.Value)

        gridEncargosGerais.DataSource = Nothing
        gridEncargosGerais.DataBind()
        ddlSituacao.Enabled = True

        grdProdutos.DataSource = Nothing
        grdProdutos.DataBind()

        'Toolbar
        lnkNovo.Parent.Visible = False
        lnkReenviar.Parent.Visible = False
        lnkCancelar.Parent.Visible = False
        lnkVisualizar.Parent.Parent.Parent.Visible = False
        btnContabil.Visible = False
        lnkConsultar.Parent.Visible = True

        ddlUnidadeNegocio.Enabled = True
        ddlEmpresa.Enabled = True
        btnConsultaClientes.Visible = True
        cmbCarteira.Enabled = True
        LnkParcelamento.Enabled = True
        ddlCondicaoDePagamento.Enabled = True
        cmbFormas.Enabled = True
        grdCondicoes.Enabled = True
        cmdOkVencimento.Enabled = True
        txtDataVencimento.Enabled = True
        txtCodigoDeBarras.Text = String.Empty
        txtCodigoDeBarras.Enabled = False
        BtValidarCodBarras.Enabled = False
        CkbCodigoDeBarras.Checked = False
        ckPreImpresso.Checked = False
        CkbCodigoDeBarras.Enabled = False
        ckPreImpresso.Enabled = False

        chkProvisao.Checked = False
        If FinanceiroVirtual Then chkProvisao.Visible = False

        lnkAdicionarItem.Parent.Parent.Visible = True

        'Tab Contábil
        TabContabil.Visible = False
        ddlProdutoContabilizacao.Items.Clear()
        gridRazao.DataSource = Nothing
        gridRazao.DataBind()
        lblDebito.Text = String.Empty
        lblCredito.Text = String.Empty


        'Tab Vencimentos
        TabVencimentosold.Visible = False
        divTitulo.Visible = False

        lblBanco.Text = "Banco"
        lblAgencia.Text = "Agência"
        lblContaCorrente.Text = "Conta"
        divBanco.Visible = False

        VerificaCodigoAVista()

        txtDataVencimento.Text = String.Empty
        txtValorVencimento.Text = String.Empty
        grdCondicoes.DataSource = Nothing
        grdCondicoes.DataBind()

        txtTotalNota.Text = "0,00"
        txtDataNota.Text = DateTime.Today.ToString("dd/MM/yyyy")

        objNotaFiscal = New NotaFiscal()
        objNotaFiscal.IUD = "I"
        objNotaFiscal.CodigoUnidadeDeNegocio = ddlUnidadeNegocio.SelectedValue
        Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
        objNotaFiscal.CodigoEmpresa = strEmpresa(0)
        objNotaFiscal.EnderecoEmpresa = strEmpresa(1)
        objNotaFiscal.Usuario = Session("ssNomeUsuario")
        objNotaFiscal.UsuarioInclusao = Session("ssNomeUsuario")

        objNotaFiscal.EntradaSaida = eEntradaSaida.Saida

        objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.NFC_e
        objNotaFiscal.NossaEmissao = True
        objNotaFiscal.CodigoSituacao = 4
        objNotaFiscal.Eletronica = True
        objNotaFiscal.CarregandoNota = False

        hdfCodigoCliente.Value = String.Empty
        txtCPF.Text = String.Empty
        txtNovoCliente.Text = String.Empty
        txtCelular.Text = String.Empty

        ddlSituacao.SelectedValue = 1
        ddlSituacao.Enabled = False
        cmbCarteira.SelectedIndex = 0

        ddlUsuarios.Items.Clear()
        ddlUsuarios.Items.Add("Inc.- " & Session("ssNomeUsuario"))

        Session.Remove("objNotaFiscal" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
        ucProdutoCupomFiscal.SetarHID(HID.Value)
        ucConsultaProdutoCupomFiscal.SetarHID(HID.Value)
        ucMonitorCupomFiscal.SetarHID(HID.Value)
        ucConsultaLote.SetarHID(HID.Value)
        ucEmailNFe.SetarHID(HID.Value)

        Dim dtLoteFornecedor As New DataTable("ItemLoteFornecedor")
        dtLoteFornecedor.Columns.Add("Produto", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Lote", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Fabricado", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Validade", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
        dtLoteFornecedor.Columns.Add("Consumo", Type.GetType("System.Decimal"))
        Session("objLoteFornecedor" & HID.Value) = dtLoteFornecedor

        mensagemErro = String.Empty
        objNotaFiscal.NFG = False
        SessaoSalvaNotaFiscal()

        'Verifica se a empresa está habilitada para gravar arquivo
        Dim Empresa As New [Lib].Negocio.Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
        SessaoRecuperaNotaFiscal()

        BuscarNumeroCupom()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub BuscarNumeroCupom()

        Dim Sql As String = "select Sequencia from Numerador " & vbCrLf &
                            " where Empresa_Id    = '" & UsuarioServidor.CodigoEmpresa & "'" & vbCrLf &
                            "   and EndEmpresa_Id = " & UsuarioServidor.EnderecoEmpresa & vbCrLf &
                            "   and Numerador_id  = " & eTiposNumerador.NFCe

        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(Sql, "Consulta")

        If ds Is Nothing OrElse
            ds.Tables Is Nothing OrElse
            ds.Tables.Count = 0 OrElse
            ds.Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Numerador do Cupom não cadastrado.", "~/Expedicao.aspx")
        Else
            txtAvisoCupom.Text = "Cupom: " & (ds.Tables(0).Rows(0)("Sequencia") + 1)
        End If

    End Sub

    Private Sub VerificaUnidade()
        Dim Sql As String = ""
        Sql = "SELECT Top 1 isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf &
              "             isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf &
              "             isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf &
              "  from Usuarios" & vbCrLf &
              " where Usuario_Id = '" & UsuarioServidor.NomeUsuario & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidadeNegocio.SelectedValue = Dr("AcessoUnidade")
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue, True)
            ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Private Sub VerificaCodigoAVista()
        Dim Sql As String = "Select top 1 Pagamento_Id from Pagamentos where AVista = 1"

        hdCondicaoDePagamento.Value = 0

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "ConsultaCodigo").Tables(0).Rows
            hdCondicaoDePagamento.Value = Dr("Pagamento_Id")
        Next

        If hdCondicaoDePagamento.Value > 0 Then ddlCondicaoDePagamento.SelectedValue = hdCondicaoDePagamento.Value

        cmbFormas.SelectedValue = 1
    End Sub

    Public Sub AtualizarItensNoGrid()

        SessaoRecuperaNotaFiscal()

        CarregarItensNaGrid(objNotaFiscal)

        gridEncargosGerais.DataSource = From nfEnc In objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").SelectMany(Function(s) s.Encargos)
                                        Group By nfEnc.Codigo Into ValorOficial = Sum(nfEnc.Valor)
                                        Order By IIf(Codigo = "PRODUTO", 1, IIf(Codigo = "LIQUIDO", 3, 2))
                                        Select Codigo, ValorOficial
        gridEncargosGerais.DataBind()

        If objNotaFiscal.Itens IsNot Nothing AndAlso objNotaFiscal.Itens.Count > 0 Then

            If objNotaFiscal.CodigoOperacao > 0 AndAlso objNotaFiscal.SubOperacao.Financeiro Then
                TabVencimentosold.Visible = True

                If cmbCarteira.SelectedIndex = 0 Then cmbCarteira.SelectedValue = objNotaFiscal.Itens(0).Produto.CodigoCarteiraVenda

            Else
                TabVencimentosold.Visible = False
            End If

            objNotaFiscal.AtualizaTotais()

            If (objNotaFiscal.VencimentosNota.Any()) Then
                LnkParcelamento_Click(LnkParcelamento, Nothing)
            End If

            Dim i As Integer = 0
            While i < grdProdutos.Rows.Count

                For Each drItemLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                    If drItemLote("Produto") = grdProdutos.Rows(i).Cells(1).Text Then

                        If drItemLote("Consumo") > 0 Then
                            CType(grdProdutos.Rows(i).FindControl("txtQuantidadeFiscalGrid"), TextBox).Enabled = False
                            CType(grdProdutos.Rows(i).FindControl("txtUnitarioGrid"), TextBox).Enabled = False
                            Exit For
                        End If
                    End If
                Next

                i += 1
            End While

        End If

        If lnkNovo.Parent.Visible = False AndAlso grdProdutos.Rows.Count > 0 Then lnkNovo.Parent.Visible = objNotaFiscal.IUD <> "U" AndAlso objNotaFiscal.IUD <> "D"

        txtTotalNota.Text = objNotaFiscal.TotalNota.ToString("N2")

        SessaoSalvaNotaFiscal()

    End Sub
#End Region

#Region "Sessões"
    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub
#End Region

#Region "Cabeçalho"
    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidadeNegocio.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()
        If ddlUnidadeNegocio.SelectedIndex > 0 Then
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue, True)
            objNotaFiscal.CodigoUnidadeDeNegocio = ddlUnidadeNegocio.SelectedValue
        Else
            ddlEmpresa.Items.Clear()
            objNotaFiscal.CodigoUnidadeDeNegocio = String.Empty
            objNotaFiscal.CodigoEmpresa = String.Empty
            objNotaFiscal.EnderecoEmpresa = 0
        End If
        SessaoSalvaNotaFiscal()
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEmpresa.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()

        If ddlEmpresa.SelectedIndex > 0 Then
            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")

            If objNotaFiscal.CodigoOperacao > 0 AndAlso objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 AndAlso Not objNotaFiscal.CodigoEmpresa = strEmpresa(0) Then
                objNotaFiscal.CodigoEmpresa = strEmpresa(0)
                objNotaFiscal.EnderecoEmpresa = strEmpresa(1)

                Dim objPedido As New Pedido(objNotaFiscal, hdCondicaoDePagamento.Value)

                'NUMERADOR DOS PEDIDOS
                Dim SqlN As String = "exec sp_Numerador '" & objPedido.CodigoEmpresa & "'," & objPedido.EnderecoEmpresa & "," & [Lib].Negocio.eTiposNumerador.Pedido & ""
                Dim dsN As New DataSet
                dsN = Banco.ConsultaDataSet(SqlN, "Numerador")

                Dim CodigoNumerador As Integer = 0
                If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                    CodigoNumerador = dsN.Tables(0).Rows(0).Item(0)
                End If

                If Not CodigoNumerador > 0 Then
                    MsgBox(Me.Page, "Numerador de Pedidos não cadastrado!")
                    Exit Sub
                End If

                objPedido.Codigo = CodigoNumerador
                objNotaFiscal.CodigoPedido = CodigoNumerador
                objNotaFiscal.Pedido = objPedido
                objNotaFiscal.CIFFOB = objPedido.FreteCIFFOB

                For Each row In objNotaFiscal.VencimentosNota
                    If row.CodigoProvisao = 1 Then
                        MsgBox(Me.Page, "Não pode ser alterado a Empresa com Financeiro Baixado - Título " & row.Codigo)
                        LimparCampos()
                        Exit Sub
                    Else
                        If Not objNotaFiscal.SubOperacao.Financeiro Then row.IUD = "D"

                        row.CodigoUnidadeDeNegocio = objNotaFiscal.CodigoUnidadeDeNegocio
                        row.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                        row.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa

                        row.CodigoEmpresaPedido = objNotaFiscal.CodigoEmpresa
                        row.EndEmpresaPedido = objNotaFiscal.EnderecoEmpresa
                        row.CodigoPedido = objNotaFiscal.CodigoPedido

                        row.CodigoCliente = objNotaFiscal.CodigoCliente
                        row.EndCliente = objNotaFiscal.EnderecoCliente

                        row.CodigoDestinatario = objNotaFiscal.CodigoCliente
                        row.EndDestinatario = objNotaFiscal.EnderecoCliente
                        row.NomeDoDestinatario = ""

                        row.Movimento = objNotaFiscal.Movimento

                        row.Historico = "PAGTO " & Trim(objNotaFiscal.TipoDeDocumento.Historico) & " " & Trim(objNotaFiscal.Codigo) & "-" & Trim(objNotaFiscal.Serie) & " - " & Trim(objNotaFiscal.Cliente.Nome) & " - " & Trim(row.Carteira.Descricao)
                        row.Observacoes = objNotaFiscal.Observacoes
                    End If
                Next
            Else
                objNotaFiscal.CodigoEmpresa = strEmpresa(0)
                objNotaFiscal.EnderecoEmpresa = strEmpresa(1)
            End If

            'Verifica se a empresa está habilitada para gravar arquivo
            Dim Empresa As New [Lib].Negocio.Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
        Else
            objNotaFiscal.CodigoEmpresa = String.Empty
            objNotaFiscal.EnderecoEmpresa = 0
        End If

        SessaoSalvaNotaFiscal()
    End Sub

    Protected Sub btnConsultaClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConsultaClientes.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
    End Sub

    Protected Sub cmbCarteira_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCarteira.SelectedIndexChanged

        If cmbCarteira.SelectedIndex > 0 Then

            SessaoRecuperaNotaFiscal()

            If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(cmbCarteira.SelectedValue)

                For Each rowTit As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosNota
                    rowTit.CodigoCarteira = objCarteira.CodigoCarteira
                    rowTit.ContaContabilCliente = objCarteira.CodigoContaCliente
                Next

                SessaoSalvaNotaFiscal()

                grdCondicoes.DataSource = objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D")
                grdCondicoes.DataBind()
            End If
        End If
    End Sub

#End Region

#Region "TAB Produtos"
    Protected Sub lnkAdicionarItem_Click(sender As Object, e As EventArgs) Handles lnkAdicionarItem.Click
        Popup.ConsultaDeProdutoCupomFiscal(Me.Page, "objProdutoNotaNF" & HID.Value)
        ucConsultaProdutoCupomFiscal.InicializarUC()
    End Sub

    Protected Sub imgExcluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        SessaoRecuperaNotaFiscal()
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        Dim i As Integer = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = row.Cells(1).Text)

        If objNotaFiscal.IUD = "I" And objNotaFiscal.CodigoSituacao = 4 Then
            objNotaFiscal.Itens.Remove(objNotaFiscal.Itens(i))
            objNotaFiscal.AtualizaTotais()
            If Not objNotaFiscal.Itens.Any() Then
                LnkLimparParcelamento_Click(lnkLimparParcelamento, Nothing)
            End If
        End If

        SessaoSalvaNotaFiscal()
        AtualizarItensNoGrid()
    End Sub
#End Region

#Region "TAB Vencimentos"
    Protected Sub LnkParcelamento_Click(sender As Object, e As EventArgs) Handles LnkParcelamento.Click
        SessaoRecuperaNotaFiscal()

        If cmbCarteira.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Carteira não foi selecionada.")
            Exit Sub
        ElseIf ddlCondicaoDePagamento.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Condições de Pagamento não foi selecionado.")
            Exit Sub
        ElseIf cmbFormas.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Forma de Pagamento não foi selecionada.")
            Exit Sub
        ElseIf Not objNotaFiscal.Itens.Any() Then
            MsgBox(Me.Page, "Não foi adicionado nenhum Produto na Sacola.")
            Exit Sub
        End If

        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(cmbCarteira.SelectedValue)
        Dim objTributo As New [Lib].Negocio.Encargo()

        If String.IsNullOrWhiteSpace(objCarteira.CodigoContaCliente) AndAlso String.IsNullOrWhiteSpace(objTributo.Codigo) Then
            MsgBox(Me.Page, "Encargo não foi selecionado.")
            Exit Sub
        End If

        If objNotaFiscal.IUD = "I" Then
            cmbFormas.Enabled = True
        Else
            cmbFormas.Enabled = False
        End If

        objNotaFiscal.VencimentosNota.NF = objNotaFiscal
        hdCondicaoDePagamento.Value = ddlCondicaoDePagamento.SelectedValue

        objNotaFiscal.VencimentosNota.ParcelarNotasFiscaisGerais(hdCondicaoDePagamento.Value, objNotaFiscal.TotalNota, 0, 0, 2)

        For Each rowTit As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosNota
            rowTit.CodigoCarteira = objCarteira.CodigoCarteira
            rowTit.CodigoTipoPgto = cmbFormas.SelectedValue
            rowTit.Tributo = IIf(Not String.IsNullOrWhiteSpace(objTributo.Codigo) AndAlso objTributo.Codigo.Length > 0, objTributo.Codigo, "")
            rowTit.UsuarioInclusaoData = Now()
            If Not String.IsNullOrWhiteSpace(objTributo.Codigo) AndAlso objTributo.Codigo.Length > 0 Then
                rowTit.ContaContabilCliente = objTributo.ContaCredito
            Else
                rowTit.ContaContabilCliente = objCarteira.CodigoContaCliente
            End If
        Next

        SessaoSalvaNotaFiscal()

        grdCondicoes.DataSource = objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D")
        grdCondicoes.DataBind()

        If cmbFormas.SelectedValue = 4 Then
            txtCodigoDeBarras.Enabled = True
            BtValidarCodBarras.Enabled = True
            CkbCodigoDeBarras.Enabled = True
            ckPreImpresso.Enabled = True
        Else
            txtCodigoDeBarras.Text = ""
            CkbCodigoDeBarras.Checked = False
            ckPreImpresso.Checked = False
            txtCodigoDeBarras.Enabled = False
            BtValidarCodBarras.Enabled = False
            CkbCodigoDeBarras.Enabled = False
            ckPreImpresso.Enabled = False
        End If

        If (cmbFormas.SelectedValue.ToString.Contains("3") Or
                cmbFormas.SelectedValue.ToString.Contains("6") Or
                cmbFormas.SelectedValue.ToString.Contains("7") Or
                cmbFormas.SelectedValue.ToString.Contains("11")) AndAlso Not divBanco.Visible Then
            divBanco.Visible = True
            ucConsultaDadosBancarios.Limpar()
            ucConsultaDadosBancarios.CarregaGrid(objNotaFiscal.CodigoCliente, objNotaFiscal.EnderecoCliente)
            Popup.ConsultaDeDadosBancarios(Me.Page, "objDadosBancariosNFG" & HID.Value)
        Else
            divBanco.Visible = False
        End If

        cmbCarteira.Enabled = False
        ddlCondicaoDePagamento.Enabled = False
        cmbFormas.Enabled = False
    End Sub

    Protected Sub LnkLimparParcelamento_Click(sender As Object, e As EventArgs) Handles lnkLimparParcelamento.Click
        SessaoRecuperaNotaFiscal()

        If objNotaFiscal.VencimentosNota.Any() Then
            objNotaFiscal.VencimentosNota = Nothing

            grdCondicoes.DataSource = objNotaFiscal.VencimentosNota
            grdCondicoes.DataBind()

            cmbCarteira.Enabled = True
            ddlCondicaoDePagamento.Enabled = True
            cmbFormas.Enabled = True

        End If

        SessaoSalvaNotaFiscal()
    End Sub

    Protected Sub BtValidarCodBarras_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtValidarCodBarras.Click
        If ckPreImpresso.Checked = False Then
            If ddlFormasDePagamento.SelectedValue = 4 Then
                If Trim(txtCodigoDeBarras.Text) <> "" Then
                    Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
                    If CkbCodigoDeBarras.Checked Then txtCodigoDeBarras.Text = Funcoes.FormataLinhaDigitavelOriginal(txtCodigoDeBarras.Text)
                    If Funcoes.ValidaCodigoBarras(txtCodigoDeBarras.Text, CkbCodigoDeBarras.Checked, txtDataVencimento.Text, txtValorVencimento.Text, strEmpresa(0), strEmpresa(1), Banco) Then
                        MsgBox(Me.Page, "Codigo Barras Valido!")
                    Else
                        MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)))
                    End If
                Else
                    MsgBox(Me.Page, "Código de Barras não foi informado")
                End If
            Else
                MsgBox(Me.Page, "Preenchimento do Codigo de Barras Somente Aceito para Boletos Bancarios")
            End If
        Else
            MsgBox(Me.Page, "Sistema não Valida Codigo De Barras de Boletos Pré Impressos...")
        End If
    End Sub

    Protected Sub ddlFormasDePagamento_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlFormasDePagamento.SelectedIndexChanged
        If ddlFormasDePagamento.SelectedValue = 4 Then
            txtCodigoDeBarras.Enabled = True
            BtValidarCodBarras.Enabled = True
            CkbCodigoDeBarras.Enabled = True
            ckPreImpresso.Enabled = True
        Else
            If (ddlFormasDePagamento.SelectedValue.ToString.Contains("3") Or
                ddlFormasDePagamento.SelectedValue.ToString.Contains("6") Or
                ddlFormasDePagamento.SelectedValue.ToString.Contains("7") Or
                ddlFormasDePagamento.SelectedValue.ToString.Contains("11")) AndAlso Not divBanco.Visible Then
                divBanco.Visible = True
                ucConsultaDadosBancarios.Limpar()
                SessaoRecuperaNotaFiscal()
                ucConsultaDadosBancarios.CarregaGrid(objNotaFiscal.CodigoCliente, objNotaFiscal.EnderecoCliente)
                Popup.ConsultaDeDadosBancarios(Me.Page, "objDadosBancariosNFG" & HID.Value)
            Else
                divBanco.Visible = False
            End If

            txtCodigoDeBarras.Text = ""
            CkbCodigoDeBarras.Checked = False
            ckPreImpresso.Checked = False
            txtCodigoDeBarras.Enabled = False
            BtValidarCodBarras.Enabled = False
            CkbCodigoDeBarras.Enabled = False
            ckPreImpresso.Enabled = False
        End If
    End Sub

    Protected Sub grdCondicoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        divTitulo.Visible = True

        If objNotaFiscal.CodigoPedido > 0 Then ddlCondicaoDePagamento.SelectedValue = objNotaFiscal.Pedido.CodigoCondicaoPagamento

        cmbCarteira.SelectedValue = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoCarteira

        Dim Parametros As New Hashtable
        Parametros.Clear()
        Parametros.Add("listarTudo", "N")

        ddl.Carregar(ddlFormasDePagamento, CarregarDDL.Tabela.TipoDePagamento, "", True, Parametros)

        lblTitulo.Text = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Codigo
        txtDataVencimento.Text = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Prorrogacao
        txtValorVencimento.Text = String.Format("{0:N2}", objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).ValorDoDocumento)
        txtCodigoDeBarras.Text = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDeBarras
        CkbCodigoDeBarras.Checked = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDigitado
        ckPreImpresso.Checked = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDeBarrasPreImpresso
        ddlFormasDePagamento.SelectedValue = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoTipoPgto
        ddlFormasDePagamento.Enabled = True
        ddlFormasDePagamento_SelectedIndexChanged(Nothing, Nothing)


        SessaoSalvaNotaFiscal()

        If objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoProvisao = 2 Then
            ddlFormasDePagamento.Enabled = True
            txtValorVencimento.Enabled = True
            cmdOkVencimento.Enabled = True
            txtDataVencimento.Enabled = True

            If ddlFormasDePagamento.SelectedValue = 4 Then
                txtCodigoDeBarras.Enabled = True
                BtValidarCodBarras.Enabled = True
                CkbCodigoDeBarras.Enabled = True
                ckPreImpresso.Enabled = True
            Else
                If (ddlFormasDePagamento.SelectedValue.ToString.Contains("3") Or
                    ddlFormasDePagamento.SelectedValue.ToString.Contains("6") Or
                    ddlFormasDePagamento.SelectedValue.ToString.Contains("7") Or
                    ddlFormasDePagamento.SelectedValue.ToString.Contains("11")) Then
                    divBanco.Visible = True
                Else
                    divBanco.Visible = False
                End If

                txtCodigoDeBarras.Text = ""
                CkbCodigoDeBarras.Checked = False
                ckPreImpresso.Checked = False
                CkbCodigoDeBarras.Checked = False
                ckPreImpresso.Checked = False
                txtCodigoDeBarras.Enabled = False
                BtValidarCodBarras.Enabled = False
                CkbCodigoDeBarras.Enabled = False
                ckPreImpresso.Enabled = False
            End If
        Else
            txtCodigoDeBarras.Enabled = False
            BtValidarCodBarras.Enabled = False
            CkbCodigoDeBarras.Enabled = False
            ckPreImpresso.Enabled = False
            ddlFormasDePagamento.Enabled = False
            txtValorVencimento.Enabled = False
            cmdOkVencimento.Enabled = False
            txtDataVencimento.Enabled = False
        End If
    End Sub

    Protected Sub cmdOkVencimento_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOkVencimento.Click
        SessaoRecuperaNotaFiscal()
        Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")

        If Not IsDate(txtDataVencimento.Text) OrElse Convert.ToDateTime(txtDataVencimento.Text) < Convert.ToDateTime(txtDataNota.Text) Then
            MsgBox(Me.Page, "Data de vencimento não pode ser menor que " & txtDataNota.Text)
            Exit Sub
        End If

        If ddlFormasDePagamento.SelectedValue = 4 AndAlso Not ckPreImpresso.Checked AndAlso Not Funcoes.ValidaCodigoBarras(txtCodigoDeBarras.Text, CkbCodigoDeBarras.Checked, txtDataVencimento.Text, txtValorVencimento.Text, strEmpresa(0), strEmpresa(1), Banco) Then
            MsgBox(Me.Page, "Código de Barras Inválido!")
            Exit Sub
        End If

        Dim DataUtil = Funcoes.ValidaDataUtil(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, CDate(txtDataVencimento.Text))

        LimparVencimentos()
    End Sub

    Private Sub LimparVencimentos()
        lblTitulo.Text = ""
        txtDataVencimento.Text = ""
        CkbCodigoDeBarras.Checked = False
        ckPreImpresso.Checked = False
        txtCodigoDeBarras.Enabled = False
        BtValidarCodBarras.Enabled = False
        CkbCodigoDeBarras.Enabled = False
        ckPreImpresso.Enabled = False
        txtCodigoDeBarras.Text = ""
        ddlFormasDePagamento.SelectedValue = ""
        grdCondicoes.SelectedIndex = -1
        divTitulo.Visible = False
    End Sub

    Protected Sub btnDadosBancarios_Click(sender As Object, e As EventArgs) Handles btnDadosBancarios.Click
        SessaoRecuperaNotaFiscal()
        ucConsultaDadosBancarios.Limpar()
        ucConsultaDadosBancarios.CarregaGrid(objNotaFiscal.CodigoCliente, objNotaFiscal.EnderecoCliente)
        Popup.ConsultaDeDadosBancarios(Me.Page, "objDadosBancariosNFG" & HID.Value)
    End Sub
#End Region

#Region "Methods & Functions"

    Public Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(objNotaFiscal.CodigoEmpresa) Then
            mensagemErro = "Selecione uma Empresa para Continuar"
            Return False
        ElseIf String.IsNullOrWhiteSpace(objNotaFiscal.CodigoCliente) Then
            mensagemErro = "Selecione um Cliente para Continuar"
            Return False
        ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento, "NOTAS FISCAIS") Then
            mensagemErro = "Movimento da Nota Fiscal já fechado para esta data"
            Return False
        ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento, "CONTABIL") Then
            mensagemErro = "Movimento Contábil já fechado para esta data"
            Return False
        ElseIf objNotaFiscal.CodigoSituacao = 0 Then
            mensagemErro = "Situação não foi selecionada"
            Return False
        ElseIf objNotaFiscal.IUD = "I" AndAlso objNotaFiscal.CodigoSituacao <> 4 Then
            mensagemErro = "Situação na gravação deve ser Pendente na Inclusão"
            Return False
        ElseIf objNotaFiscal.CodigoTipoDeDocumento = 0 Then
            mensagemErro = "Tipo de Documento não foi selecionado"
            Return False
        ElseIf objNotaFiscal.TotalNota = 0 Then
            mensagemErro = "Total da Nota Fiscal não pode ser Zero"
            Return False
        ElseIf objNotaFiscal.Itens.Count = 0 Then
            mensagemErro = "Não foi adicionado nenhum Produto na Sacola"
            Return False
        End If

        For Each item In objNotaFiscal.Itens
            If item.QuantidadeFiscal <= 0 OrElse item.QuantidadeFisica <= 0 Then
                mensagemErro = "Não foi adicionado a quantidade do item: " & item.CodigoProduto & " - " & item.Produto.Nome
                Return False
            End If

            If item.Unitario <= 0 Then
                mensagemErro = "Não foi adicionado o Valor do item: " & item.CodigoProduto & " - " & item.Produto.Nome
                Return False
            End If

            If Not item.QuantidadeFiscal = CType(Session("objLoteFornecedor" & HID.Value), DataTable).Compute("SUM(Consumo)", "Produto = " & item.CodigoProduto) Then
                mensagemErro = "Quantidade do Item " & item.CodigoProduto & " está diferente do total de Lote(s)."
                Return False
            End If
        Next

        If objNotaFiscal.VencimentosNota.Count = 0 Then
            mensagemErro = "Financeiro não foi selecionado!"
            Return False
        End If

        Return True
    End Function

    Public Sub SalvarCupomFiscal()
        SessaoRecuperaNotaFiscal()

        If String.IsNullOrWhiteSpace(hdfCodigoCliente.Value) AndAlso Not String.IsNullOrWhiteSpace(txtCPF.Text) Then

            Dim ObjCliente As New [Lib].Negocio.Cliente(txtCPF.Text.RemoveMask, 0)

            Dim Sqls As New ArrayList

            If ObjCliente Is Nothing Then
                ObjCliente = New [Lib].Negocio.Cliente()
                ObjCliente.IUD = "I"
                ObjCliente.CodigoSituacao = eSituacao.Normal
                ObjCliente.Codigo = txtCPF.Text.RemoveMask
                ObjCliente.CodigoEndereco = 0
                ObjCliente.Nome = Funcoes.EliminarCaracteresEspeciais(txtNovoCliente.Text.Replace("-", " ")).ToUpper()
                ObjCliente.Fantasia = Funcoes.EliminarCaracteresEspeciais(txtNovoCliente.Text.Replace("-", " ")).ToUpper()

                ObjCliente.Telefone = txtCelular.Text.RemoveMask

                ObjCliente.CodigoCategoria = 4 'Comerciante

                ObjCliente.ClienteDesde = Today
                ObjCliente.NascimentoConstituicao = Today

                ObjCliente.CodigoPais = 1058 'Brasil
                ObjCliente.CodigoEstado = "MT"
                ObjCliente.CodigoMunicipio = 2704
                ObjCliente.Cidade = "CAMPO GRANDE"
                ObjCliente.CodigoRegiao = 16
                ObjCliente.Endereco = "AV. PRINCIPAL I"
                ObjCliente.Bairro = "NUCLEO INDUSTRIAL"
                ObjCliente.Numero = 1246
                ObjCliente.CEP = "79108550"

                Dim objTipo As New [Lib].Negocio.ClientexTipo(ObjCliente)
                objTipo.IUD = "I"
                objTipo.CodigoTipo = 4 'Clientes
                ObjCliente.Tipos.Add(objTipo)

                ObjCliente.UsuarioInclusao = Session("ssNomeUsuario")
                ObjCliente.UsuarioInclusaoData = DateTime.Now()

                If Not ObjCliente.Salvar Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Return
                End If
            End If

            objNotaFiscal.CodigoCliente = ObjCliente.Codigo
            objNotaFiscal.EnderecoCliente = ObjCliente.CodigoEndereco
        Else
            'Dados da Quimica
            objNotaFiscal.CodigoCliente = objNotaFiscal.Empresa.Codigo
            objNotaFiscal.EnderecoCliente = objNotaFiscal.Empresa.CodigoEndereco
        End If

        SessaoRecuperaNotaFiscal()

        objNotaFiscal.DataNota = DateTime.Now()
        objNotaFiscal.Movimento = CDate(txtDataNota.Text)
        objNotaFiscal.DataInclusao = DateTime.Now()

        objNotaFiscal.CodigoDeposito = objNotaFiscal.CodigoEmpresa
        objNotaFiscal.EnderecoDeposito = objNotaFiscal.EnderecoEmpresa
        objNotaFiscal.CodigoDestino = objNotaFiscal.CodigoCliente
        objNotaFiscal.EnderecoDestino = objNotaFiscal.EnderecoCliente

        If objNotaFiscal.Itens.Count > 0 Then
            objNotaFiscal.CodigoOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").FirstOrDefault.CodigoOperacao
            objNotaFiscal.CodigoSubOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").FirstOrDefault.CodigoSubOperacao
        End If

        If ValidarCampos() Then

            Dim Sqls As New ArrayList
            'CRIA O PEDIDO COM BASE NA NOTA
            Dim objPedido As New Pedido(objNotaFiscal, hdCondicaoDePagamento.Value)

            'NUMERADOR DOS PEDIDOS
            Dim SqlN As String = "exec sp_Numerador '" & objPedido.CodigoEmpresa & "'," & objPedido.EnderecoEmpresa & "," & [Lib].Negocio.eTiposNumerador.Pedido & ""
            Dim dsN As New DataSet
            dsN = Banco.ConsultaDataSet(SqlN, "Numerador")

            Dim CodigoNumerador As Integer = 0
            If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                CodigoNumerador = dsN.Tables(0).Rows(0).Item(0)
            End If

            If Not CodigoNumerador > 0 Then
                MsgBox(Me.Page, "Numerador de Pedidos não cadastrado!")
                Exit Sub
            End If

            objPedido.Codigo = CodigoNumerador
            objNotaFiscal.CodigoPedido = CodigoNumerador
            objNotaFiscal.Pedido = objPedido
            objNotaFiscal.CIFFOB = objPedido.FreteCIFFOB

            For Each row In objNotaFiscal.VencimentosNota
                If Not row.CodigoProvisao = 1 Then
                    If Not objNotaFiscal.SubOperacao.Financeiro Then row.IUD = "D"

                    row.CodigoUnidadeDeNegocio = objNotaFiscal.CodigoUnidadeDeNegocio
                    row.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                    row.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa

                    row.CodigoEmpresaPedido = objNotaFiscal.CodigoEmpresa
                    row.EndEmpresaPedido = objNotaFiscal.EnderecoEmpresa
                    row.CodigoPedido = objNotaFiscal.CodigoPedido

                    row.CodigoCliente = objNotaFiscal.CodigoCliente
                    row.EndCliente = objNotaFiscal.EnderecoCliente

                    row.CodigoDestinatario = objNotaFiscal.CodigoCliente
                    row.EndDestinatario = objNotaFiscal.EnderecoCliente
                    row.NomeDoDestinatario = ""

                    row.Movimento = objNotaFiscal.Movimento

                    row.Historico = "PAGTO " & Trim(objNotaFiscal.TipoDeDocumento.Historico) & " " & Trim(objNotaFiscal.Codigo) & "-" & Trim(objNotaFiscal.Serie) & " - " & Trim(objNotaFiscal.Cliente.Nome) & " - " & Trim(row.Carteira.Descricao)
                    row.Observacoes = objNotaFiscal.Observacoes
                End If
            Next

            objNotaFiscal.CarregandoNota = False

            'Adiciona Lotes Nota Fiscal X Lote
            For Each item In objNotaFiscal.Itens

                For Each drItemLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows

                    If drItemLote("Produto") = item.CodigoProduto Then

                        If drItemLote("Consumo") > 0 Then
                            Dim lote As New NotaFiscalXLote(item)
                            lote.IUD = "I"
                            lote.Lote = drItemLote("Lote")
                            lote.Fabricado = drItemLote("Fabricado")
                            lote.Validade = drItemLote("Validade")
                            lote.Quantidade = drItemLote("Consumo")
                            item.Lotes.Add(lote)
                        End If
                    End If
                Next
            Next

            SessaoSalvaNotaFiscal()

            objPedido.SalvarSql(Sqls)
            objNotaFiscal.SalvarSql(Sqls)

            If Banco.GravaBanco(Sqls) Then
                If objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                    Dim fm As New FilesManager()
                    If fm.IsConnect() Then
                        Dim msgSefaz As String = String.Empty
                        If btnModo.Text.ToUpper().Trim() = "MODO NORMAL" Then
                            If VerificarNFe() Then
                                If EnviarSEFAZ(msgSefaz) Then
                                    MsgBox(Me.Page, "Cupom fiscal " & objNotaFiscal.Codigo & " incluído com sucesso! " & msgSefaz, eTitulo.Sucess)
                                    LimparCampos()
                                Else
                                    MsgBox(Me.Page, "Cupom fiscal " & objNotaFiscal.Codigo & " incluído, porém ainda não foi HOMOLOGADO. " & msgSefaz)
                                    LimparCampos()
                                End If
                            End If
                        Else
                            If EnviarSEFAZ(msgSefaz) Then
                                MsgBox(Me.Page, "Cupom fiscal " & objNotaFiscal.Codigo & " incluído com sucesso! " & msgSefaz, eTitulo.Sucess)
                                LimparCampos()
                            Else
                                MsgBox(Me.Page, "Cupom fiscal " & objNotaFiscal.Codigo & " incluído, porém ainda não foi HOMOLOGADO. " & msgSefaz)
                                LimparCampos()
                            End If
                        End If
                    Else
                        MsgBox(Me.Page, "Cupom fiscal " & objNotaFiscal.Codigo & " incluído, porém Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor! Cupom Fiscal ainda não foi HOMOLOGADO.")
                        LimparCampos()
                    End If
                Else
                    LimparCampos()
                End If

                'FIM SEFAZ

                LimparCampos()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Else
            MsgBox(Me.Page, mensagemErro)
        End If
    End Sub

    Private Function EnviarSEFAZ(ByRef msgNFE As String) As Boolean
        If objNotaFiscal Is Nothing Then
            msgNFE = "É necessário consultar um documento NF-e para realizar o envio para a SEFAZ!"
            Return False
        End If

        If GravarNFeXpress(objNotaFiscal) Then
            Dim obj As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Format("resp-nfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa)

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(90)

            While obj Is Nothing AndAlso Now < tempoLimite
                obj = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                System.Threading.Thread.Sleep(3000)
            End While

            If obj IsNot Nothing Then
                Dim lstMsg As String() = obj.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
                Dim chave As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("CHAVE")).FirstOrDefault()
                Dim recibo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RECIBO")).FirstOrDefault()
                Dim protocolo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NPROTOCOLO")).FirstOrDefault()
                Dim lote As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NUMEROLOTE")).FirstOrDefault()

                Dim strCodigo As String = String.Empty
                If resultado IsNot Nothing AndAlso resultado.Length > 0 Then
                    strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strMsg As String = String.Empty
                If mensagem IsNot Nothing AndAlso mensagem.Length > 0 Then
                    strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strChave As String = String.Empty
                If chave IsNot Nothing AndAlso chave.Length > 0 Then
                    strChave = chave.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strRecibo As String = String.Empty
                If recibo IsNot Nothing AndAlso recibo.Length > 0 Then
                    strRecibo = recibo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strProtocolo As String = String.Empty
                If protocolo IsNot Nothing AndAlso protocolo.Length > 0 Then
                    strProtocolo = protocolo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strLote As String = String.Empty
                If lote IsNot Nothing AndAlso lote.Length > 0 Then
                    strLote = lote.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                'ATUALIZAR NFE PENDENCIAS COM OS DADOS DO RETORNO
                Dim Sqls As New ArrayList

                If (Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "100" AndAlso strCodigo <> "110" AndAlso strCodigo <> "4017" AndAlso strCodigo <> "302") _
                    Or Not strChave.Length = 44 _
                    Or Not strProtocolo.Length = 15 Then

                    lnkNovo.Parent.Visible = False

                    Sql = "UPDATE NFEPendencias " & vbCrLf &
                          "   SET ObservacoesFiscais = '" & objNotaFiscal.Observacoes & "', " & vbCrLf &
                          "       Retorno = '" & strCodigo & "', " & vbCrLf &
                          "       MsgRetorno = '" & strMsg.Replace("'", "") & "' " & vbCrLf
                    If strRecibo.Length > 0 Then Sql &= ", Recibo = '" & strRecibo & "' " & vbCrLf
                    If strChave.Length > 0 AndAlso strChave.Length = 44 Then Sql &= ", ChaveNfe = '" & strChave & "' " & vbCrLf
                    If strLote.Length > 0 Then Sql &= ", Lote = '" & strLote & "' " & vbCrLf
                    If strProtocolo.Length > 0 AndAlso strProtocolo.Length = 15 Then Sql &= ", Protocolo = '" & strProtocolo & "' " & vbCrLf
                    Sql &= "WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf &
                           "  AND EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                           "  AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "' " & vbCrLf &
                           "  AND EndCliente_Id = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                           "  AND EntradaSaida_Id = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                           "  AND Serie_Id = '" & objNotaFiscal.Serie & "' " & vbCrLf &
                           "  AND Nota_Id = " & objNotaFiscal.Codigo & "; " & vbCrLf
                    Sqls.Add(Sql)

                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-statusnfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                    Sqls.Add(Sql)
                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-nfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                    Sqls.Add(Sql)

                    If Banco.GravaBanco(Sqls) Then
                        msgNFE = String.Format("{0} - {1}", strCodigo, strMsg)
                    Else
                        msgNFE = HttpContext.Current.Session("ssMessage")
                    End If

                    Return False
                End If

                msgNFE = String.Format("{0} - {1}", strCodigo, strMsg)

                Sqls.Clear()

                Sql = "DELETE NFEContingencia " & vbCrLf &
                      " WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf &
                      "   AND EndEmpresa_Id = '" & objNotaFiscal.EnderecoEmpresa & "' " & vbCrLf &
                      "   AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "   AND EndCliente_Id = '" & objNotaFiscal.EnderecoCliente & "'" & vbCrLf &
                      "   AND EntradaSaida_Id = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                      "   AND Serie_Id = '" & objNotaFiscal.Serie & "' " & vbCrLf &
                      "   AND Nota_Id = '" & objNotaFiscal.Codigo & "'; " & vbCrLf
                Sqls.Add(Sql)

                If btnModo.Text.ToUpper().Trim() <> "MODO NORMAL" AndAlso strCodigo = "4017" Then
                    Sql = "INSERT INTO NFEContingencia "
                    Sql &= "(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, "
                    Sql &= "NfExpress, Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, Protocolo, RegDPEC) "
                    Sql &= "VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", '" & objNotaFiscal.CodigoCliente & "', '"
                    Sql &= objNotaFiscal.EnderecoCliente & "', '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "', '" & objNotaFiscal.Serie & "', '" & objNotaFiscal.Codigo & "', '"
                    Sql &= objNotaFiscal.DataInclusao.ToSqlDate() & "', '" & Format(objNotaFiscal.DataInclusao, "HH:mm:ss") & "', '" & UsuarioServidor.NomeUsuario & "', '"
                    Sql &= String.Format("nfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "', 'INCLUIR', '"
                    Sql &= strCodigo & "', '" & strMsg & "', '" & strRecibo & "', '" & strChave & "', '" & strLote & "', '2', '" & objNotaFiscal.Observacoes & "', '', '" & strProtocolo & "', ''); "
                    Sqls.Add(Sql)
                Else
                    Sql = "INSERT INTO NFERealizadas "
                    Sql &= "(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, "
                    Sql &= "NfExpress, Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, Protocolo, RegDPEC) "
                    Sql &= "VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ", '" & objNotaFiscal.CodigoCliente & "', '"
                    Sql &= objNotaFiscal.EnderecoCliente & "', '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "', '" & objNotaFiscal.Serie & "', '" & objNotaFiscal.Codigo & "', '"
                    Sql &= objNotaFiscal.DataInclusao.ToSqlDate() & "', '" & Format(objNotaFiscal.DataInclusao, "HH:mm:ss") & "', '" & UsuarioServidor.NomeUsuario & "', '"
                    Sql &= String.Format("nfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "', 'INCLUIR', '"
                    Sql &= strCodigo & "', '" & strMsg & "', '" & strRecibo & "', '" & strChave & "', '" & strLote & "', '1', '" & objNotaFiscal.Observacoes & "', '', '" & strProtocolo & "', ''); "
                    Sqls.Add(Sql)
                End If

                Sql = "DELETE NFEPendencias " & vbCrLf &
                      " WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf &
                      "   AND EndEmpresa_Id = '" & objNotaFiscal.EnderecoEmpresa & "'" & vbCrLf &
                      "   AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "' " & vbCrLf &
                      "   AND EndCliente_Id = '" & objNotaFiscal.EnderecoCliente & "' " & vbCrLf &
                      "   AND EntradaSaida_Id = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                      "   AND Serie_Id = '" & objNotaFiscal.Serie & "'" & vbCrLf &
                      "   AND Nota_Id = '" & objNotaFiscal.Codigo & "'; " & vbCrLf
                Sqls.Add(Sql)

                If strCodigo = "110" OrElse strCodigo = "302" Then
                    Sql = "UPDATE NotasFiscais Set Situacao = 10 "
                Else
                    Sql = "UPDATE NotasFiscais Set Situacao = 1 "
                End If

                Sql &= "WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf &
                       "  AND EndEmpresa_Id = '" & objNotaFiscal.EnderecoEmpresa & "' " & vbCrLf &
                       "  AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "' " & vbCrLf &
                       "  AND EndCliente_Id = '" & objNotaFiscal.EnderecoCliente & "' " & vbCrLf &
                       "  AND EntradaSaida_Id = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                       "  AND Serie_Id = '" & objNotaFiscal.Serie & "' " & vbCrLf &
                       "  AND Nota_Id = '" & objNotaFiscal.Codigo & "'; " & vbCrLf
                Sqls.Add(Sql)

                If strCodigo = "110" OrElse strCodigo = "301" OrElse strCodigo = "302" OrElse strCodigo = "303" Then
                    objNotaFiscal.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)

                    If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                        For Each tit In objNotaFiscal.VencimentosNota
                            If tit.CodigoSituacao = eSituacao.Normal AndAlso tit.CodigoProvisao = eProvisao.Previsao Then
                                Sql = "Update ContasAPagar Set Provisao = 3 " & vbCrLf &
                                        "Where Registro_id = " & tit.Codigo
                                Sqls.Add(Sql)

                                Sql = "Update ContasAReceber Set Provisao = 3 " & vbCrLf &
                                        "Where Registro_id = " & tit.Codigo
                                Sqls.Add(Sql)
                            End If
                        Next
                    End If

                    If objNotaFiscal.NotasReferenciais IsNot Nothing AndAlso objNotaFiscal.NotasReferenciais.Count > 0 Then
                        For Each item In objNotaFiscal.NotasReferenciais
                            item.IUD = "D"
                            item.SalvarSql(Sqls)
                        Next
                    End If

                    Sql = "Delete NotasFiscaisXTransportadores" & vbCrLf &
                          " Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                          "   and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                          "   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                          "   and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                          "   and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          "   and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                          "   and Nota_Id         = " & objNotaFiscal.Codigo
                    Sqls.Add(Sql)


                    Sql = " Delete NotasFiscaisXRomaneios" & vbCrLf &
                          "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                          "    and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                          "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          "    and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                          "    and Romaneio_Id     = " & objNotaFiscal.CodigoRomaneio
                    Sqls.Add(Sql)

                End If

                If Not Banco.GravaBanco(Sqls) Then
                    msgNFE = HttpContext.Current.Session("ssMessage")
                    Return False
                End If
                'FIM DA ROTINA

                'Pega a Chave para Envio de Email
                objNotaFiscal.ChaveNFE = strChave

                SessaoSalvaNotaFiscal()

                Dim usoDenegado As Boolean = False

                If strCodigo = "110" OrElse strCodigo = "302" Then
                    msgNFE = msgNFE & "Uso Denegado : Irregularidade fiscal do destinatario. " & objNotaFiscal.Codigo & "-" & objNotaFiscal.Serie & "!"
                    usoDenegado = True
                ElseIf btnModo.Text.ToUpper().Trim() = "MODO NORMAL" Then
                    If Not String.IsNullOrWhiteSpace(objNotaFiscal.Empresa.EmailNFE) AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.Cliente.EmailNFE) Then
                        [Lib].Negocio.DocumentoEletronico.SendMailNFce(objNotaFiscal, msgNFE, False)
                    End If

                    If [Lib].Negocio.DocumentoEletronico.ImprimirNFCe(objNotaFiscal, 1, msgNFE, True) Then
                        Try
                            Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", objNotaFiscal.ChaveNFE), eTipoDeDocumento.Nota)
                            If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                                Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                                System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                                Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                            End If
                        Catch ex As Exception
                            msgNFE = msgNFE & "Não foi possível encontrar o arquivo DANFE da nota fiscal " & objNotaFiscal.Codigo & "-" & objNotaFiscal.Serie & "!"
                        End Try
                    End If
                Else
                    Try
                        Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", strChave), eTipoDeDocumento.Nota)
                        If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(strChave) Then
                            Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", strChave))
                            System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                            Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", strChave))
                        End If
                    Catch ex As Exception
                        msgNFE = msgNFE & "Não foi possível encontrar o arquivo DANFE da nota fiscal " & objNotaFiscal.Codigo & "-" & objNotaFiscal.Serie & "!"
                    End Try
                End If

                Sqls.Clear()
                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-statusnfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)
                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-nfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)
                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailnfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)
                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfenfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)

                'Excluir tíitulo caso Nota seja Denegada - Furlan - 03/09/2021
                If usoDenegado Then
                    If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                        For Each dr As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosNota
                            dr.IUD = "D"
                            dr.SalvarSql(Sqls, False)
                        Next

                        objNotaFiscal.VencimentosNota.NF = objNotaFiscal
                        'Atualiza a provisão do pedido.
                        For Each dr As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosPedido.Where(Function(t) t.CodigoProvisao = eProvisao.Provisao)
                            dr.IUD = "U"
                            dr.SalvarSql(Sqls, False)
                        Next
                    End If
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    msgNFE = msgNFE & HttpContext.Current.Session("ssMessage")
                    Return False
                ElseIf Not String.IsNullOrEmpty(msgNFE) AndAlso Not strCodigo = "100" Then
                    Return False
                Else
                    Return True
                End If
            Else
                MsgBox(Me.Page, "Falha no retorno da Sefaz ref. a Homologação da NF-e, entre nas Pêndencias e verifique para realizar o reenvio à SEFAZ!")
                Return False
            End If
        End If
    End Function

    Private Function VerificarNFe() As Boolean
        Dim Sqls As New ArrayList
        Dim aux As Boolean = True
        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("statusnfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa)
        obj.Texto = String.Empty
        obj.SalvarSql(Sqls)

        If Not Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-statusnfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa)

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(90)

        While resp Is Nothing AndAlso Now < tempoLimite
            resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
            System.Threading.Thread.Sleep(3000)
        End While

        If resp IsNot Nothing Then
            Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            Sqls.Clear()

            If String.IsNullOrWhiteSpace(strCodigo) Then
                aux = False
                MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE.")
            ElseIf strCodigo = "107" Then
                aux = True

                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-statusnfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)
            ElseIf strCodigo = "4036" Then
                aux = False

                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-statusnfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                Sqls.Add(Sql)

                MsgBox(Me.Page, "Guardian em Modo de Contingência não pode ser usado para Emissão em Modo Normal.")
            Else
                aux = False
                MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
            End If

            If Sqls.Count > 0 Then Banco.GravaBanco(Sqls)
        Else
            aux = False
            MsgBox(Me.Page, "Verifique o Guardian no Servidor, não foi recebida nenhuma resposta do Modo de Operação.")
        End If

        Return aux
    End Function

    Private Function GravarNFeXpress(ByVal nf As [Lib].Negocio.NotaFiscal) As Boolean
        Dim aux As Boolean = True
        Try
            Dim Sqls As New ArrayList
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"
            obj.NomeArquivo = String.Format("nfce{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

            obj.Texto = [Lib].Negocio.DocumentoEletronico.getTextoNFCe4G(nf)

            If String.IsNullOrWhiteSpace(obj.Texto) Then
                MsgBox(Me.Page, "Não foi possível construir o arquivo texto para emissão do cupom fiscal!")
                Return False
            End If
            obj.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Return False
            End If
        Catch ex As Exception
            aux = False
            Throw New Exception(ex.Message)
        End Try
        Return aux
    End Function

    Private Sub EmContingencia(ByVal modo As [Lib].Negocio.eModo)
        Try
            SessaoRecuperaNotaFiscal()
            Dim Sqls As New ArrayList
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"
            obj.NomeArquivo = "contingencianfce.txt"
            obj.Texto = getTextoContingencia(objNotaFiscal, modo)
            obj.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Exit Sub
            End If

            'AGUARDAR RESPOSTA SEFAZ
            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = "resp-contingencianfce.txt"

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(90)

            While resp Is Nothing AndAlso Now < tempoLimite
                resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                System.Threading.Thread.Sleep(3000)
            End While

            If resp IsNot Nothing Then
                Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                If Not String.IsNullOrWhiteSpace(strCodigo) Then
                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                    If modo = eModo.Contingencia AndAlso strCodigo = "4021" AndAlso strMsg.Contains("contingencia2") Then
                        btnModo.BackColor = Drawing.Color.Red
                        btnModo.Text = "MODO CONTINGÊNCIA"
                    ElseIf modo = eModo.Normal AndAlso strCodigo = "4021" AndAlso strMsg.Contains("normal") Then
                        btnModo.BackColor = Drawing.Color.Green
                        btnModo.Text = "MODO NORMAL"
                    End If
                End If
            Else
                MsgBox(Me.Page, "Falha no retorno ref. o Modo de Operação, verifique novamente.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getTextoContingencia(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal modo As [Lib].Negocio.eModo) As String
        Dim sb As New StringBuilder()
        sb.Append("MODALIDADE=" & IIf(modo = eModo.Normal, "0", "2") & ControlChars.CrLf)
        sb.Append("CNPJ=" & nf.CodigoEmpresa & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Function VerificarModo(ByVal Codigo As String, ByVal Empresa As String, ByVal NossaEmissao As Boolean, ByVal verMensagem As Boolean) As Boolean
        Dim aux As Boolean = True

        Try
            Dim Sqls As New ArrayList
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"
            obj.NomeArquivo = String.Format("statusnfce{0:000000000}#{1}.txt", Codigo, Empresa)
            obj.Texto = String.Empty
            obj.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Return False
            End If

            'AGUARDAR RESPOSTA SEFAZ
            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Format("resp-statusnfce{0:000000000}#{1}.txt", Codigo, Empresa)

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(90)

            While resp Is Nothing AndAlso Now < tempoLimite
                resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                System.Threading.Thread.Sleep(3000)
            End While

            If resp IsNot Nothing Then
                Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                Sqls.Clear()

                If String.IsNullOrWhiteSpace(strCodigo) Then
                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE.")
                    aux = False
                ElseIf strCodigo = "107" Then
                    btnModo.BackColor = Drawing.Color.Green
                    btnModo.Text = "MODO NORMAL"
                    aux = True
                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-statusnfce{0:000000000}#{1}.txt", Codigo, Empresa) & "';"
                    Sqls.Add(Sql)
                    If verMensagem Then MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg) & " - MODO NORMAL")
                ElseIf strCodigo = "4036" Then
                    btnModo.BackColor = Drawing.Color.Red
                    btnModo.Text = "MODO CONTINGÊNCIA"
                    aux = True
                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-statusnfce{0:000000000}#{1}.txt", Codigo, Empresa) & "';"
                    Sqls.Add(Sql)
                    If verMensagem Then MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg) & " - MODO CONTINGÊNCIA")
                Else
                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                    aux = False
                End If

                If Sqls.Count > 0 Then Banco.GravaBanco(Sqls)
            Else
                MsgBox(Me.Page, "Verifique o Guardian no Servidor, não foi recebida nenhuma resposta do Modo de Operação. ")
                aux = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            aux = False
        End Try

        Return aux
    End Function

    Protected Sub txt_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim txt As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)

        Dim txtQtde As TextBox = row.FindControl("txtQuantidadeFiscalGrid")
        Dim txtUnitario As TextBox = row.FindControl("txtUnitarioGrid")

        If String.IsNullOrEmpty(txtQtde.Text) OrElse String.IsNullOrEmpty(txtUnitario.Text) OrElse CInt(txtQtde.Text) = 0 OrElse Convert.ToDecimal(txtUnitario.Text) = 0 Then Exit Sub

        SessaoRecuperaNotaFiscal()

        Dim i As Integer = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = row.Cells(1).Text)

        objNotaFiscal.Itens(i).CodigoOperacao = 21
        objNotaFiscal.Itens(i).CodigoSubOperacao = 35

        objNotaFiscal.Itens(i).NotaFiscal.Cliente = New Cliente()
        objNotaFiscal.Itens(i).NotaFiscal.Cliente.CodigoEstado = "MS"
        objNotaFiscal.Itens(i).NotaFiscal.Cliente.InscricaoEstadual = ""
        objNotaFiscal.Itens(i).QuantidadeFiscal = txtQtde.Text
        objNotaFiscal.Itens(i).QuantidadeFisica = txtQtde.Text
        objNotaFiscal.Itens(i).PesoFiscal = txtQtde.Text
        objNotaFiscal.Itens(i).Unitario = txtUnitario.Text


        Dim qtdeLote As Decimal = objNotaFiscal.Itens(i).QuantidadeFiscal
        Dim buscarLote As Boolean = True

        If Not IsDBNull(CType(Session("objLoteFornecedor" & HID.Value), DataTable).Compute("SUM(Consumo)", "Produto = " & objNotaFiscal.Itens(i).CodigoProduto)) Then

            If objNotaFiscal.Itens(i).QuantidadeFiscal = CType(Session("objLoteFornecedor" & HID.Value), DataTable).Compute("SUM(Consumo)", "Produto = " & objNotaFiscal.Itens(i).CodigoProduto) Then
                buscarLote = False
            Else
                For k As Integer = CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Count - 1 To 0 Step -1
                    If CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows(k)("Produto") = objNotaFiscal.Itens(i).CodigoProduto Then
                        CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows(k).Delete()
                    End If
                Next k
            End If

        End If

        If buscarLote Then
            Dim pLote = New OrdemParaProducaoXConsumo()
            Dim ds As New DataSet
            ds = pLote.buscarLoteDeFornecedor(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, "'" & objNotaFiscal.Itens(i).CodigoProduto & "'")

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Não foi encontrado lote para esse Produto!", eTitulo.Info)
                Exit Sub
            End If

            If objNotaFiscal.Itens(i).QuantidadeFiscal > ds.Tables(0).Compute("SUM(Quantidade)", "Produto = " & objNotaFiscal.Itens(i).CodigoProduto) Then
                MsgBox(Me.Page, "Quantidade informada é maior que a quatidade de Lote disponível!", eTitulo.Info)
                txtQtde.Text = 0
                txtUnitario.Text = 0
                Exit Sub
            End If

            For Each dr In ds.Tables(0).Rows
                Dim drItem As DataRow = CType(Session("objLoteFornecedor" & HID.Value), DataTable).NewRow()
                drItem("Produto") = dr("Produto")
                drItem("Lote") = dr("Lote")
                drItem("Fabricado") = CDate(dr("Fabricado")).ToString("dd/MM/yyyy")
                drItem("Validade") = CDate(dr("Validade")).ToString("dd/MM/yyyy")
                drItem("Quantidade") = dr("Quantidade")

                If qtdeLote > 0 Then

                    If qtdeLote > dr("Quantidade") Then

                        drItem("Consumo") = dr("Quantidade")

                        qtdeLote -= drItem("Consumo")
                    Else

                        drItem("Consumo") = qtdeLote

                        qtdeLote = 0

                    End If
                Else
                    drItem("Consumo") = 0
                End If

                CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Add(drItem)
            Next
        End If


        If Not String.IsNullOrWhiteSpace(txtQtde.Text) AndAlso Not String.IsNullOrWhiteSpace(txtUnitario.Text) Then
            objNotaFiscal.Itens(i).ValorTotal = Math.Round((CDbl(txtQtde.Text) * CDbl(txtUnitario.Text)), 2).ToString("N2")
        End If

        SessaoSalvaNotaFiscal()

        AtualizarItensNoGrid()

    End Sub

    Protected Sub imgSelecionar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            SessaoRecuperaNotaFiscal()

            Dim chkProduto As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(chkProduto.NamingContainer, GridViewRow)

            If String.IsNullOrWhiteSpace(grdProdutos.Rows(row.RowIndex).Cells(5).Text) OrElse CDec(grdProdutos.Rows(row.RowIndex).Cells(5).Text) = 0 Then
                MsgBox(Me.Page, "Quantidade ou unitário não foi informado!", eTitulo.Info)
                Exit Sub
            End If

            ucConsultaLote.Limpar()
            ucConsultaLote.SetarHID(HID.Value)
            ucConsultaLote.MostrarLote(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Itens(row.RowIndex))

            Popup.ConsultaDeLote(Me.Page, "objConsultaDeLote" & HID.Value)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub imgLimpar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try

            SessaoRecuperaNotaFiscal()

            If objNotaFiscal.IUD = "U" Then Exit Sub

            Dim chkProduto As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(chkProduto.NamingContainer, GridViewRow)

            CType(grdProdutos.Rows(row.RowIndex).FindControl("txtQuantidadeFiscalGrid"), TextBox).Enabled = True
            CType(grdProdutos.Rows(row.RowIndex).FindControl("txtUnitarioGrid"), TextBox).Enabled = True

            CType(grdProdutos.Rows(row.RowIndex).FindControl("txtQuantidadeFiscalGrid"), TextBox).Text = "0"
            CType(grdProdutos.Rows(row.RowIndex).FindControl("txtUnitarioGrid"), TextBox).Text = "0,00"
            grdProdutos.Rows(row.RowIndex).Cells(5).Text = "0,00"

            For Each drItemLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                If drItemLote("Produto") = grdProdutos.Rows(row.RowIndex).Cells(1).Text Then
                    drItemLote("Consumo") = 0
                End If
            Next

            Dim i As Integer = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = row.Cells(1).Text)

            objNotaFiscal.Itens(i).QuantidadeFiscal = 0
            objNotaFiscal.Itens(i).QuantidadeFisica = 0
            objNotaFiscal.Itens(i).PesoFiscal = 0
            objNotaFiscal.Itens(i).Unitario = 0
            objNotaFiscal.Itens(i).ValorTotal = 0

            SessaoSalvaNotaFiscal()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Public Sub InutilizarNFCe()
        Dim Sqls As New ArrayList
        Dim obj As New [Lib].Negocio.Fil()

        obj.IUD = "I"
        obj.NomeArquivo = String.Format("inutilizanfce{0}-{1}#{2}.txt", objNotaFiscal.Codigo.ToString().Trim(), objNotaFiscal.Codigo.ToString().Trim(), objNotaFiscal.Empresa.Codigo)
        obj.Texto = GetTextoInutilizacao()

        obj.SalvarSql(Sqls)

        If Not Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Exit Sub
        End If

        'AGUARDANDO RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Empty

        fileName = String.Format("resp-inutilizanfce{0}-{1}#{2}.txt", objNotaFiscal.Codigo.ToString().Trim(), objNotaFiscal.Codigo.ToString().Trim(), objNotaFiscal.Empresa.Codigo)

        While resp Is Nothing
            resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
        End While

        If resp IsNot Nothing Then
            Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
            Dim protocolo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NPROTOCOLO")).FirstOrDefault()

            Dim strCodigo As String = String.Empty
            If resultado IsNot Nothing AndAlso resultado.Length > 0 Then
                strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            End If

            Dim strMsg As String = String.Empty
            If mensagem IsNot Nothing AndAlso mensagem.Length > 0 Then
                strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            End If

            Dim strProtocolo As String = String.Empty
            If protocolo IsNot Nothing AndAlso protocolo.Length > 0 Then
                strProtocolo = protocolo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            End If
        End If
    End Sub

    Public Sub ReenviarNFCe()

        SessaoRecuperaNotaFiscal()

        Dim fm As New FilesManager()
        If fm.IsConnect() Then
            Dim msgSefaz As String = String.Empty
            If btnModo.Text.ToUpper().Trim() = "MODO NORMAL" Then
                If VerificarNFe() Then
                    If EnviarSEFAZ(msgSefaz) Then
                        MsgBox(Me.Page, "Cupom fiscal " & objNotaFiscal.Codigo & " incluído com sucesso! " & msgSefaz, eTitulo.Sucess)
                        LimparCampos()
                    Else
                        MsgBox(Me.Page, "Cupom fiscal " & objNotaFiscal.Codigo & " incluído, porém ainda não foi HOMOLOGADO. " & msgSefaz)
                        LimparCampos()
                    End If
                End If
            Else
                If EnviarSEFAZ(msgSefaz) Then
                    MsgBox(Me.Page, "Cupom fiscal " & objNotaFiscal.Codigo & " incluído com sucesso! " & msgSefaz, eTitulo.Sucess)
                    LimparCampos()
                Else
                    MsgBox(Me.Page, "Cupom fiscal " & objNotaFiscal.Codigo & " incluído, porém ainda não foi HOMOLOGADO. " & msgSefaz)
                    LimparCampos()
                End If
            End If
        Else
            MsgBox(Me.Page, "Cupom fiscal " & objNotaFiscal.Codigo & " incluído, porém Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor! Cupom Fiscal ainda não foi HOMOLOGADO.")
            LimparCampos()
        End If
    End Sub

    Private Function GetTextoInutilizacao() As String
        Dim sb As New StringBuilder()

        sb.Append("CODIGOUF=" & objNotaFiscal.Empresa.Municipio.EstadoIbge & ControlChars.CrLf)
        sb.Append("ANO =" & Mid(Now.Year.ToString, 3, 2) & ControlChars.CrLf)
        sb.Append("CNPJ =" & objNotaFiscal.Empresa.Codigo & ControlChars.CrLf)
        sb.Append("MODELO = 65" & ControlChars.CrLf)
        sb.Append("SERIE =" & objNotaFiscal.Serie.Trim() & ControlChars.CrLf)
        sb.Append("NFEINI =" & objNotaFiscal.Codigo.ToString().Trim() & ControlChars.CrLf)
        sb.Append("NFEFIM =" & objNotaFiscal.Codigo.ToString().Trim() & ControlChars.CrLf)
        sb.Append("JUSTIFICATIVA =" & Funcoes.EliminarCaracteresEspeciais("Numeração não será utilizada") & ControlChars.CrLf)

        Return sb.ToString()
    End Function

    Public Sub CancelarCupomFiscal()
        Try
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal.CodigoSituacao = 2 Then
                MsgBox(Me.Page, "Este cupom já está cancelada!")
                Exit Sub
            ElseIf objNotaFiscal.CodigoSituacao = 7 Then
                MsgBox(Me.Page, "Este cupom está no aguardo da SEFAZ para o seu cancelamento, não se pode alterar a observação!")
                Exit Sub
            ElseIf Not objNotaFiscal.NossaEmissao Then
                MsgBox(Me.Page, "Este cupom não é de nossa emissão, sendo assim não pode ser cancelado!")
                Exit Sub
            ElseIf Not objNotaFiscal.Eletronica Then
                MsgBox(Me.Page, "Este cupom não é eletrônico!")
                Exit Sub
            ElseIf String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                MsgBox(Me.Page, "Este cupom não possui a chave de NFC-e da SEFAZ!")
                Exit Sub
            ElseIf String.IsNullOrWhiteSpace(objNotaFiscal.ProtocoloNota) Then
                MsgBox(Me.Page, "Este cupom não possui o número de protocolo da SEFAZ!")
                Exit Sub
            ElseIf Not btnModo.Text.ToUpper().Trim() = "MODO NORMAL" Then
                MsgBox(Me.Page, "Cancelamento não pode ser feito em Modo de Contingência. ")
                Exit Sub
            End If

            objNotaFiscal.ObservacaoCancelamento = "Cancelamento de Cupom Fiscal"
            objNotaFiscal.IUD = "C"

            If objNotaFiscal.Salvar() Then
                If VerificarNFe() Then
                    Dim msgNFE As String = String.Empty

                    Dim Sqls As New ArrayList

                    If [Lib].Negocio.DocumentoEletronico.CancelarNFCe(objNotaFiscal, msgNFE) Then

                        CancelarNFCe(Sqls, objNotaFiscal)

                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-statusnfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-nfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailnfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfenfce{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        Else
                            If Not String.IsNullOrEmpty(msgNFE) Then
                                MsgBox(Me.Page, msgNFE)
                            End If
                            MsgBox(Me.Page, "Cupom Fiscal cancelado com sucesso!")
                            LimparCampos()
                        End If
                    Else
                        If Trim(msgNFE.Split("-")(0)) = "501" Then
                            Sqls.Clear()

                            Sql = "DELETE NFEPendencias " & vbCrLf &
                                  " WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf &
                                  "   AND EndEmpresa_Id = '" & objNotaFiscal.EnderecoEmpresa & "' " & vbCrLf &
                                  "   AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                                  "   AND EndCliente_Id = '" & objNotaFiscal.EnderecoCliente & "'" & vbCrLf &
                                  "   AND EntradaSaida_Id = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                                  "   AND Serie_Id = '" & objNotaFiscal.Serie & "' " & vbCrLf &
                                  "   AND Nota_Id = '" & objNotaFiscal.Codigo & "'; " & vbCrLf
                            Sqls.Add(Sql)

                            Sql = "UPDATE NotasFiscais Set Situacao = 1, " & vbCrLf

                            If objNotaFiscal.ObservacoesControleInterno.Length > 0 Then
                                Sql &= "ObservacoesControleInterno = '" & objNotaFiscal.ObservacoesControleInterno & ". TENTATIVA DE CANCELAMENTO NEGADA PELA SEFAZ POR ESTAR FORA DO PRAZO - USUARIO: " & UsuarioServidor.NomeUsuario & " AS " & Now.ToString("dd/MM/yyyy HH:mm:ss") & "'" & vbCrLf
                            Else
                                Sql &= "ObservacoesControleInterno = 'TENTATIVA DE CANCELAMENTO NEGADA PELA SEFAZ POR ESTAR FORA DO PRAZO - USUARIO: " & UsuarioServidor.NomeUsuario & " AS " & Now.ToString("dd/MM/yyyy HH:mm:ss") & "'" & vbCrLf
                            End If
                            Sql &= "WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf &
                                   "  AND EndEmpresa_Id = '" & objNotaFiscal.EnderecoEmpresa & "' " & vbCrLf &
                                   "  AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "' " & vbCrLf &
                                   "  AND EndCliente_Id = '" & objNotaFiscal.EnderecoCliente & "' " & vbCrLf &
                                   "  AND EntradaSaida_Id = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                                   "  AND Serie_Id = '" & objNotaFiscal.Serie & "' " & vbCrLf &
                                   "  AND Nota_Id = '" & objNotaFiscal.Codigo & "'; " & vbCrLf
                            Sqls.Add(Sql)

                            If Not Banco.GravaBanco(Sqls) Then
                                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            End If
                        End If

                        MsgBox(Me.Page, msgNFE)

                        LimparCampos()

                    End If
                End If
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Public Sub CancelarNFCe(ByRef Sqls As ArrayList, ByRef nf As [Lib].Negocio.NotaFiscal)

        Sql = "UPDATE NotasFiscaisXEncargos Set" & vbCrLf &
              "    Base       = 0," & vbCrLf &
              "    Percentual = 0," & vbCrLf &
              "    Valor      = 0 " & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Sql = " UPDATE NotasFiscaisXItens Set" & vbCrLf &
              "    PesoFiscal       = 0," & vbCrLf &
              "    QuantidadeFisica = 0," & vbCrLf &
              "    QuantidadeFiscal = 0," & vbCrLf &
              "    Unitario         = 0," & vbCrLf &
              "    Valor            = 0 " & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Dim obs As String = nf.ObservacoesControleInterno
        If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
            obs = obs & ". CANCELADO PELO IP " & ddlUsuarios.SelectedItem.Text
        Else
            obs = "CANCELADO PELO IP " & ddlUsuarios.SelectedItem.Text
        End If

        Sql = "UPDATE NotasFiscais Set" & vbCrLf &
              "   Situacao = 2 " & vbCrLf &
              "	 ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
              "	 ,UsuarioAlteracaoData = getdate() " & vbCrLf &
              "  ,ObservacoesControleInterno  = '" & obs & "'" & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)



        Sql = "UPDATE ContasAReceber SET" & vbCrLf &
              "    ContasAReceber.Provisao = 3" & vbCrLf &
              "  FROM ContasAReceber " & vbCrLf &
              " INNER JOIN (select Empresa_Id, EndEmpresa_Id, Cliente_id, Endcliente_Id, Titulo_Id" & vbCrLf &
              "               from NotaFiscalXTitulo " & vbCrLf &
              "				 where Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "				   and EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "				   and Cliente_id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "    			   and Endcliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "				   and EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "				   and Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "				   and Nota_id         = " & nf.Codigo & vbCrLf &
              "             ) AS TxN " & vbCrLf &
              "    ON ContasAReceber.Empresa     = TxN.Empresa_Id " & vbCrLf &
              "	  AND ContasAReceber.EndEmpresa  = TxN.EndEmpresa_Id " & vbCrLf &
              "	  AND ContasAReceber.Cliente     = TxN.Cliente_id " & vbCrLf &
              "	  AND ContasAReceber.EndCliente  = TxN.Endcliente_Id " & vbCrLf &
              "	  AND ContasAReceber.Registro_Id = TxN.Titulo_Id " & vbCrLf &
              " WHERE ContasAReceber.Provisao = 2; "
        Sqls.Add(Sql)

        Sql = "DELETE NotaFiscalXTitulo " & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)


        Sql = "DELETE RAZAO" & vbCrLf &
      " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
      "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
      "   And Cliente_nf      ='" & nf.CodigoCliente & "'" & vbCrLf &
      "   And EndCliente_nf   = " & nf.EnderecoCliente & vbCrLf &
      "   And EntradaSaida_nf ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
      "   And Serie_nf        ='" & nf.Serie & "'" & vbCrLf &
      "   And Numero_nf       = " & nf.Codigo & vbCrLf &
      "   And Lote_Id        in (9,10,11); "
        Sqls.Add(Sql)

        Sql = "DELETE NFEPendencias " & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   And EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   And Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   And EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   And EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   And Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   And Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)

        Sql = "UPDATE NFERealizadas SET" & vbCrLf &
              "    Operacao   = 'Cancelado'," & vbCrLf &
              "    Retorno    = '101'," & vbCrLf &
              "    MsgRetorno = 'Cancelamento de NFC-e homologado'" & vbCrLf &
              " WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
              "   AND EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
              "   AND Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
              "   AND EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
              "   AND EntradaSaida_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   AND Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
              "   AND Nota_Id         = " & nf.Codigo & ";"
        Sqls.Add(Sql)
    End Sub

#End Region

#Region "ToolBar"

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("CupomFiscal", "LEITURA") Then
                Popup.ConsultaMonitorCupomFiscal(Me.Page, "objNotaFiscal" & HID.Value)
                ucMonitorCupomFiscal.InicializarUC(CDate(txtDataNota.Text).ToString("yyyy-MM-dd"))
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível consultar a nota")
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CupomFiscal", "GRAVAR") Then

                If Not String.IsNullOrWhiteSpace(txtCPF.Text) Then

                    Dim nCliente As String = txtCPF.Text.RemoveMask

                    If Not nCliente.Length = 11 Then
                        MsgBox(Me.Page, "Cliente informado não é CPF!")
                        Exit Sub
                    End If
                End If

                SalvarCupomFiscal()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível salvar o cupom fiscal." & ex.Message)
        End Try
    End Sub

    Protected Sub lnkReenviar_Click(sender As Object, e As EventArgs) Handles lnkReenviar.Click
        Try
            ReenviarNFCe()
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível atualizar a nota." & ex.Message)
        End Try
    End Sub

    Protected Sub lnkVisualizar_Click(sender As Object, e As EventArgs) Handles lnkVisualizar.Click
        Try
            SessaoRecuperaNotaFiscal()
            If objNotaFiscal IsNot Nothing Then
                Dim fileName As String = Server.MapPath(String.Format("~/Files/nota{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa))
                If System.IO.File.Exists(fileName) Then
                    System.IO.File.Delete(fileName)
                End If

                Dim str As String = String.Empty

                str = [Lib].Negocio.DocumentoEletronico.getTextoNFCe4G(objNotaFiscal)

                Using sw As New StreamWriter(fileName)
                    sw.WriteLine(str)
                    sw.Close()
                End Using

                Response.Clear()
                Response.ClearHeaders()
                Response.AddHeader("content-length", str.Length.ToString())
                Response.ContentType = "text/plain"
                Response.AppendHeader("content-disposition", "attachment;filename=" & String.Format("nota{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa))
                Response.Write(str)
                Response.End()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnRecontabilizar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnRecontabilizar.Click
        Try
            If Funcoes.VerificaPermissao("RECONTABILIZAR", "ALTERAR") Then

                SessaoRecuperaNotaFiscal()

                If objNotaFiscal.Itens.Count > 0 Then
                    Dim sqls As New ArrayList
                    objNotaFiscal.Razao.ExcluiContabilizacaoNotaSQL(sqls)
                    objNotaFiscal.Razao.ContabilizarNotaSql(sqls)

                    If Banco.GravaBanco(sqls) Then
                        Dim objNotaAntes = New NotaFiscal(objNotaFiscal)
                        objNotaFiscal = New NotaFiscal(objNotaAntes)
                        objNotaFiscal.NotaFiscalOriginal = New NotaFiscal(objNotaFiscal)
                        SessaoSalvaNotaFiscal()

                        ddlProdutoContabilizacao.Items.Clear()

                        ddlProdutoContabilizacao.Items.Add(New ListItem("Todos os Lançamentos", "0"))
                        For Each row In objNotaFiscal.Itens
                            ddlProdutoContabilizacao.Items.Add(New ListItem("Produto " & row.CodigoProduto & " - " & row.Produto.Nome, row.CodigoProduto))
                        Next

                        objNotaFiscal.LancamentosContabeis.CalcularSaldo()
                        gridRazao.DataSource = objNotaFiscal.LancamentosContabeis.OrderBy(Function(s) s.Sequencia)
                        gridRazao.DataBind()

                        lblDebito.Text = objNotaFiscal.LancamentosContabeis.Sum(Function(s) s.DebitoOficial).ToString("N2")
                        lblCredito.Text = objNotaFiscal.LancamentosContabeis.Sum(Function(s) s.CreditoOficial).ToString("N2")

                        MsgBox(Me.Page, "Nota Recontabilizada com Sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, "Erro Durante a Recontabilizacão da nota")
                    End If
                Else
                    MsgBox(Me.Page, "Consute a Nota Fiscal para recontabilização")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para recontabilizar Nota Fiscal")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkCancelar_Click(sender As Object, e As EventArgs) Handles lnkCancelar.Click
        Try
            If Funcoes.VerificaPermissao("CupomFiscal", "EXCLUIR") Then
                CancelarCupomFiscal()
                LimparCampos()
            Else
                MsgBox(Me.Page, "Usuario sem permissao para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível excluir a nota")
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível limpar a nota")
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CupomFiscal")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region


End Class