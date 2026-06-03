Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucPedidoLancamentoItem
    Inherits BaseUserControl


#Region "Inicialização"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                ddl.Carregar(ddlGrupoProduto, CarregarDDL.Tabela.GrupoProduto)
                ddl.Carregar(ddlClassificacao, CarregarDDL.Tabela.TabelaDeClassificacoes)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub SetarParametros(LinhaProdutoPedido As Integer)
        chkRetencao.Parent.Visible = False
        HIDLinhaProduto.Value = LinhaProdutoPedido

        Limpar()
        SessaoRecuperaPedido()

        HFMediaOficialCompra.Value = "0"
        HFMediaMoedaCompra.Value = "0"

        If LinhaProdutoPedido >= 0 Then

            '#AcrescimoDesconto
            If objPedido.Itens(LinhaProdutoPedido).Encargos IsNot Nothing AndAlso objPedido.Itens(LinhaProdutoPedido).Encargos.Count > 0 Then
                If objPedido.Itens(LinhaProdutoPedido).Encargos.Find(Function(s) s.CodigoEncargo = "ACRESCIMOS") IsNot Nothing Then
                    objPedido.Itens(LinhaProdutoPedido).Acrescimo = objPedido.Itens(LinhaProdutoPedido).Encargos.Where(Function(s) s.CodigoEncargo = "ACRESCIMOS").SingleOrDefault.ValorOficial
                End If
                If objPedido.Itens(LinhaProdutoPedido).Encargos.Find(Function(s) s.CodigoEncargo = "DESCONTOS") IsNot Nothing Then
                    Dim ValorOficialDoItem As Decimal = objPedido.Itens(LinhaProdutoPedido).Encargos.Find(Function(s) s.CodigoEncargo = "LIQUIDO").ValorOficial
                    Dim ValorOficialDoDesconto As Decimal = objPedido.Itens(LinhaProdutoPedido).Encargos.Find(Function(s) s.CodigoEncargo = "DESCONTOS").ValorOficial

                    If ValorOficialDoItem > 0 AndAlso ValorOficialDoDesconto > 0 Then
                        If ValorOficialDoItem - ValorOficialDoDesconto > 0 Then
                            objPedido.Itens(LinhaProdutoPedido).Desconto = ValorOficialDoDesconto
                        Else
                            MsgBox(Me.Page, "Valor do Desconto não deve ser maior do que o valor total do item. Valor Total do Item: " & ValorOficialDoItem & "  Desconto: " & ValorOficialDoDesconto)
                        End If
                    End If
                End If
            End If
            '##

            If objPedido.Itens(LinhaProdutoPedido).Produto.Agrupar.Equals("S") AndAlso objPedido.Itens.Count > 0 Then
                lnkProdutoNovo.Parent.Visible = True
            Else
                lnkProdutoNovo.Parent.Visible = False
            End If

            'Seta valores
            ddlGrupoProduto.SelectedValue = objPedido.Itens(LinhaProdutoPedido).Produto.CodigoGrupo
            ddl.Carregar(ddlProdutos, CarregarDDL.Tabela.Produto, " Situacao = 1 and Grupo ='" & ddlGrupoProduto.SelectedValue & "'")
            ddlProdutos.SelectedValue = objPedido.Itens(LinhaProdutoPedido).CodigoProduto
            HFC.Value = objPedido.Itens(LinhaProdutoPedido).UnidadeComercializacaoFatorDeConversao

            If Funcoes.VerificaPermissao("AjustarEncargosPedido", "GRAVAR") Then
                txtDataMovimentoItem.Enabled = True
            Else
                txtDataMovimentoItem.Enabled = False
            End If

            txtDataMovimentoItem.Text = Date.Now.ToShortDateString()
            txtDataEntregaItem.Text = objPedido.DataEntregaFinal.ToShortDateString()
            MudouDataEntrega()

            txtResumoQtde.Text = objPedido.Itens(LinhaProdutoPedido).Lancamentos.QuantidadeTotalPrdFaturamento.ToString("N4")
            txtResumoUntMedio.Text = CDec(objPedido.Itens(LinhaProdutoPedido).UnitarioMedioFaturamento).ToString("N10")

            txtResumoQtdeCom.Text = objPedido.Itens(LinhaProdutoPedido).Lancamentos.QuantidadeTotalPrdComercializacao.ToString("N4")

            txtResumoUntMedioCom.Text = CDec(objPedido.Itens(LinhaProdutoPedido).UnitarioMedioComercializacao).ToString("N10")

            txtResumoValor.Text = objPedido.Itens(LinhaProdutoPedido).PedidoValor.ToString("N2")

            ddlClassificacao.SelectedValue = objPedido.Itens(LinhaProdutoPedido).Classificacao.Codigo
            ddlClassificacao.Enabled = False

            carregarUnidadedeComercializacao(objPedido.Itens(LinhaProdutoPedido).Produto)
            ddlUnidadeComercializacao.SelectedValue = objPedido.Itens(LinhaProdutoPedido).CodigoUnidadeComercializacao
            ddlUnidadeComercializacao.Enabled = False

            ddlTipoLancamento.Items.Clear()
            ddlTipoLancamento.Items.Add(New ListItem("Complemento Qtde", "CQ"))
            ddlTipoLancamento.Items.Add(New ListItem("Complemento Valor", "CV"))
            ddlTipoLancamento.Items.Add(New ListItem("Estorno Qtde", "EQ"))
            ddlTipoLancamento.Items.Add(New ListItem("Estorno Valor", "EV"))

            gridProdutos.DataSource = objPedido.Itens(LinhaProdutoPedido).Lancamentos.ToArray
            gridProdutos.DataBind()

            'Desativa Campos
            ddlGrupoProduto.Enabled = False
            ddlProdutos.Enabled = False
            lnkBuscaProduto.Visible = False
        Else
            lnkProdutoNovo.Parent.Visible = False
            txtDataMovimentoItem.Text = Date.Now.ToShortDateString()
            txtDataEntregaItem.Text = objPedido.DataEntregaFinal.ToShortDateString()

            txtResumoQtde.Text = "0,00"
            txtResumoUntMedio.Text = "0"
            txtResumoValor.Text = "0,00"

            ddlTipoLancamento.Items.Clear()
            ddlTipoLancamento.Items.Add(New ListItem("Normal", "N"))
            ddlTipoLancamento.Items.Add(New ListItem("Complemento Valor", "CV"))
        End If
        MudouTipodeLancamento(ddlTipoLancamento.SelectedValue)
    End Sub
#End Region

#Region "Variáveis Locais"
    Private objPedido As [Lib].Negocio.Pedido
#End Region

#Region "Sessão"
    Private Sub SessaoRecuperaPedido()
        objPedido = CType(Session("objPedido" & HID.Value), [Lib].Negocio.Pedido)
        If objPedido Is Nothing Then objPedido = New [Lib].Negocio.Pedido
    End Sub

    Private Sub SessaoSalvaPedido()
        Session("objPedido" & HID.Value) = objPedido
    End Sub

    Public Property UsuarioSupervisor(Processo As String) As Usuario
        Get
            If UsuarioSupervisorProcesso = Processo Then
                Return Session("objSupervisor" & HID.Value)
            Else
                Return Nothing
            End If
        End Get

        Set(ByVal value As Usuario)
            Session("objSupervisor" & HID.Value) = value
            If Session("objSupervisor" & HID.Value) Is Nothing Then
                UsuarioSupervisorProcesso = ""
            Else
                UsuarioSupervisorProcesso = Processo
            End If
        End Set
    End Property

    Public Property UsuarioSupervisorProcesso() As String
        Get
            Return Session("objSupervisorProcesso" & HID.Value)
        End Get
        Set(ByVal value As String)
            Session("objSupervisorProcesso" & HID.Value) = value
        End Set
    End Property

#End Region

    Public Sub MudouTipodeLancamento(pTipo As String)
        SessaoRecuperaPedido()

        txtUnitarioCompra.Visible = True
        txtQuantidade.Parent.Visible = True
        txtUnitario.Parent.Visible = True

        txtValorTotal.Enabled = False
        txtValorTotalMoeda.Parent.Visible = False

        txtQuantidade.Text = "0,0000"
        txtUnitario.Text = "0,0000000000"

        txtQuantidadeFat.Text = "0,0000"
        txtUnitarioFat.Text = "0,0000000000"

        txtValorTotal.Text = "0,00"
        txtValorTotalMoeda.Text = "0,00"

        If Not objPedido.SubOperacao.QuantidadePedido AndAlso Not objPedido.SubOperacao.Classe = [Lib].Negocio.eClassesOperacoes.AFIXAR Then
            txtQuantidade.Parent.Visible = False
        End If

        If Not objPedido.SubOperacao.UnitarioPedido AndAlso Not objPedido.SubOperacao.Classe = [Lib].Negocio.eClassesOperacoes.AFIXAR Then
            txtUnitarioCompra.Visible = False
        End If

        If pTipo = "CV" Or pTipo = "EV" Then
            txtQuantidade.Text = "0,0000"
            txtUnitario.Text = "0,0000000000"
            txtQuantidadeFat.Text = "0,0000"
            txtUnitarioFat.Text = "0,0000000000"
            txtUnitarioCompra.Visible = False

            txtQuantidade.Parent.Visible = False
            txtUnitario.Parent.Visible = False

            txtValorTotal.Enabled = True

            If objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso Funcoes.VerificaPermissao("PedidoXItemXLancamentoMoeda", "Gravar") Then
                txtValorTotalMoeda.Parent.Visible = True
                txtValorTotalMoeda.Enabled = True
                'Busca o campo TotalOficial para para mostrá-lo ou ocultá-lo dependenco da Moeda
                For i As Integer = 0 To gridProdutos.Columns.Count - 1
                    If gridProdutos.Columns(i).HeaderText.Equals("TotalOficial") Then
                        gridProdutos.Columns(i).Visible = True
                    End If
                Next
            End If
        ElseIf pTipo = "CQ" Or pTipo = "EQ" Then 'And TdUnitario.Visible Then
            txtUnitario.Text = CDec(objPedido.Itens(HIDLinhaProduto.Value).UnitarioMedioComercializacao).ToString("N10")
            txtUnitarioFat.Text = CDec(objPedido.Itens(HIDLinhaProduto.Value).UnitarioMedioFaturamento).ToString("N10")
            If txtUnitarioFat.Visible Then txtUnitarioFat.Text = objPedido.Itens(HIDLinhaProduto.Value).UnitarioMedioFaturamento.ToString("N10")
        ElseIf pTipo = "N" AndAlso Not objPedido.SubOperacao.QuantidadePedido AndAlso Not objPedido.SubOperacao.UnitarioPedido Then
            txtValorTotal.Enabled = True
        End If
    End Sub

    Public Sub MudouDataEntrega()
        If objPedido Is Nothing Then SessaoRecuperaPedido()

        HFMediaOficialCompra.Value = "0"
        HFMediaMoedaCompra.Value = "0"

        If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso objPedido.SubOperacao.PrecoFixo Then
            Dim ValorCompra As New [Lib].Negocio.PedidoAvaliacaoProduto()
            ValorCompra.AvaliarProduto(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, objPedido.CodigoSafra, ddlProdutos.SelectedValue, objPedido.DataEntregaFinal)

            HFMediaOficialCompra.Value = ValorCompra.ValorOficial
            HFMediaMoedaCompra.Value = ValorCompra.ValorMoeda

            If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                txtUnitarioCompra.Text = CDec(HFMediaOficialCompra.Value).ToString("N10")
                txtUnitarioCompraCom.Text = (CDec(HFMediaOficialCompra.Value) * CDec(HFC.Value)).ToString("N10")
            Else
                txtUnitarioCompra.Text = CDec(HFMediaMoedaCompra.Value).ToString("N10")
                txtUnitarioCompraCom.Text = (CDec(HFMediaOficialCompra.Value) * CDec(HFC.Value)).ToString("N10")
            End If
        End If
    End Sub

    Public Function ValidaLancamento(objProduto As [Lib].Negocio.Produto) As ArrayList
        '0 - True or False
        '1 - Mensagens
        '2 - Codigo da configuracao das Operacoesxencargos só grava nos lancamentos Normais
        Dim Msg As String = ""
        Dim Valido As Boolean = True
        Dim CodigoOxE As Integer

        If Not IsNumeric(txtQuantidade.Text) Then txtQuantidade.Text = "0,0000"
        If Not IsNumeric(txtUnitario.Text) Then txtUnitario.Text = "0,0000000000"
        If Not IsNumeric(txtValorTotal.Text) Then txtValorTotal.Text = "0,00"

        If Not ddlUnidadeComercializacao.SelectedValue = "TON" AndAlso objProduto.CodigoEmbalagem = 1 AndAlso CDec(txtQuantidade.Text) > 0 Then
            'ainda vou testar - Furlan - 23/08/2022
            Dim vString As String = txtQuantidade.Text
            Dim vDecimal As Decimal = CDec(vString.Split(",")(1))

            If vDecimal > 0 Then
                Msg &= "Produto a GRANEL não pode ter casa decimal."
                Valido = False
            End If
        End If

        If objProduto.NCM.ToString.Length = 0 Then
            Msg &= "Código NCM do Produto não está cadastrado, o mesmo deve ser informado.  "
            Valido = False
        End If

        If ddlClassificacao.SelectedIndex <= 0 Then
            Msg &= "Informe a Classificação!"
            Valido = False
        End If

        If Not objPedido.FiscalAberto AndAlso Not Funcoes.VerificaAcesso(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, IIf(objPedido.DataVencimentoPedido > Now.Date, Now.Date, objPedido.DataVencimentoPedido), "PEDIDOS") Then
            Msg &= "Movimento fechado para alterar o pedido nesta data. "
            Valido = False
        End If

        If ddlTipoLancamento.SelectedValue = "N" Then
            If objPedido.SubOperacao.QuantidadePedido And CDec(txtQuantidade.Text) = 0 Then
                Msg &= "A quantidade do pedido deve ser maior que zero! "
                Valido = False
            ElseIf objPedido.SubOperacao.UnitarioPedido And CDec(txtUnitario.Text) = 0 Then
                Msg &= "O valor unitário do pedido deve ser maior que zero! "
                Valido = False
            End If

            Dim Parametros As New OperacaoXEstado
            Parametros.Empresa = Left(objPedido.CodigoEmpresa, 8)
            Parametros.CodigoGrupoProduto = objProduto.CodigoGrupo
            Parametros.CodigoProduto = objProduto.Codigo
            Parametros.CodigoOperacao = objPedido.CodigoOperacao
            Parametros.CodigoSubOperacao = objPedido.CodigoSubOperacao
            Parametros.EstadoOrigem = objPedido.Empresa.CodigoEstado
            Parametros.EstadoDestino = objPedido.Cliente.CodigoEstado
            Dim OxE As New OperacaoXEstado(Parametros)

            If OxE.Encargos.Count = 0 Then
                Msg &= "Não foram encontrados encargos para este produto! "
                Valido = False
            Else
                CodigoOxE = OxE.Codigo
            End If

        ElseIf ddlTipoLancamento.SelectedValue = "CQ" Or ddlTipoLancamento.SelectedValue = "EQ" Then
            If objPedido.SubOperacao.QuantidadePedido And CDec(txtQuantidade.Text) = 0 Then
                Msg &= "Preencha a Quantidade para efetuar o Complemento/Estorno! "
                Valido = False
            ElseIf objPedido.SubOperacao.UnitarioPedido And CDec(txtUnitario.Text) = 0 Then
                Msg &= "Preencha o valor unitário para efetuar o Complemento/Estorno! "
                Valido = False
            End If

            txtValorTotal.Text = (CDec(txtQuantidade.Text) * CDec(txtUnitario.Text)).ToString("N2")

        ElseIf ddlTipoLancamento.SelectedValue = "CV" Or ddlTipoLancamento.SelectedValue = "EV" Then
            txtQuantidade.Text = "0,0000"
            txtUnitario.Text = "0,0000000000"

            If objPedido.Itens.Count = 0 Then
                Msg &= "Faça Primeiro o Lançamento Normal para em seguida fazer o complemento da diferença! "
                Valido = False
            End If
            If CDec(txtValorTotal.Text) = 0 Then
                Msg &= "Preencha o Valor do Lançamento! "
                Valido = False
            End If
        End If

        Dim valida As New ArrayList
        valida.Add(Valido)
        valida.Add(Msg)
        valida.Add(CodigoOxE)
        Return valida
    End Function

    Public Sub ValidaSupervisor(ByVal Processo As String)
        'Abrirá user control para que o supervisor coloque login e senha e além do login seja também validada a permissão.
        UsuarioSupervisor("") = Nothing
        Dim ucSupervisor As ucSupervisor = DirectCast(Me.NamingContainer.FindControl("ucSupervisor"), ucSupervisor)
        If ucSupervisor IsNot Nothing Then
            ucSupervisor.Limpar()
            ucSupervisor.Processo = Processo
            ucSupervisor.MainUserControl = Me
            Popup.Supervisor(Me.Page, "objSupervisor" & HID.Value, "txtUsuario")
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objProdutoPXI" & HID.Value) IsNot Nothing Then
            SessaoRecuperaPedido()
            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)
            ddlGrupoProduto.SelectedValue = objProduto.CodigoGrupo

            ddlProdutos.Items.Clear()
            ddl.Carregar(ddlProdutos, CarregarDDL.Tabela.Produto, " Situacao = 1")
            ddlProdutos.SelectedValue = objProduto.Codigo.Trim()

            Dim i As Integer = objPedido.Itens.FindIndex(Function(s) s.CodigoProduto = ddlProdutos.SelectedValue)

            If i >= 0 Then
                SetarParametros(i)
            Else
                Dim prd As New Produto(ddlProdutos.SelectedValue)
                'HIDBaseDeCalculo.Value = prd.BaseCalculo
                'txtBaseDeCalculo.Text = prd.BaseCalculo
                carregarUnidadedeComercializacao(prd)
            End If

            If IsDate(txtDataEntregaItem.Text) Then MudouDataEntrega()
            MudouTipodeLancamento(ddlTipoLancamento.SelectedValue)


            Session.Remove("objProdutoPXI" & HID.Value)
        End If

    End Sub

    Public Overrides Sub Carregar(par As Dictionary(Of String, Object))
        If Session("objSupervisor" & HID.Value) IsNot Nothing Then
            UsuarioSupervisor(par("Processo")) = par("Usuario")
            lnkNovo_Click(New Object, New EventArgs)
            Session.Remove("objSupervisor" & HID.Value)
        End If

    End Sub

    Public Overrides Sub Limpar()

        If CInt(HIDLinhaProduto.Value) = -1 Then
            ddlGrupoProduto.SelectedIndex = 0
            ddlGrupoProduto.Enabled = True

            ddlProdutos.Items.Clear()
            ddlProdutos.Enabled = True
            lnkBuscaProduto.Visible = True

            txtDataMovimentoItem.Enabled = True

            ddlClassificacao.SelectedIndex = 1
            ddlClassificacao.Enabled = True

            ddlTipoLancamento.Items.Clear()
            ddlTipoLancamento.Items.Add(New ListItem("Normal", "N"))
            MudouTipodeLancamento("N")

            lblUnidFat.Text = ""
            ddlUnidadeComercializacao.Items.Clear()
            ddlUnidadeComercializacao.Enabled = True
            HFC.Value = 1

            gridProdutos.DataSource = Nothing
            gridProdutos.DataBind()
        End If

        txtDataMovimentoItem.Text = Today.ToString("dd/MM/yyyy")

        If Not IsNumeric(txtQuantidadeFat.Text) Then txtQuantidadeFat.Text = "0,0000"
        txtUnitarioFat.Text = "0,0000000000"
        If Not IsNumeric(txtQuantidade.Text) Then txtQuantidade.Text = "0,0000"
        txtUnitario.Text = "0,0000000000"
        If Not IsNumeric(txtValorTotal.Text) Then txtValorTotal.Text = "0,00"

        txtValorTotal.Enabled = False
        txtValorTotalMoeda.Enabled = False

        Session.Remove("_MainUserControl")
    End Sub

    Public Sub desabilitaProdutoNovo()
        lnkProdutoNovo.Parent.Visible = False
    End Sub

    Protected Sub ddlGrupoProduto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProduto.SelectedIndexChanged
        Try
            ddl.Carregar(ddlProdutos, CarregarDDL.Tabela.Produto, " Situacao = 1 and Grupo ='" & ddlGrupoProduto.SelectedValue & "'")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlProdutos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProdutos.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()

            Dim i As Integer = objPedido.Itens.FindIndex(Function(s) s.CodigoProduto = ddlProdutos.SelectedValue)

            If i >= 0 Then
                SetarParametros(i)
            Else
                Dim prd As New Produto(ddlProdutos.SelectedValue)
                carregarUnidadedeComercializacao(prd)
            End If

            If IsDate(txtDataEntregaItem.Text) Then MudouDataEntrega()
            MudouTipodeLancamento(ddlTipoLancamento.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub carregarUnidadedeComercializacao(pProduto As Produto)
        ddlUnidadeComercializacao.Items.Clear()

        lblUnidFat.Text = pProduto.Unidade

        For Each un In pProduto.UnidadesDeComercializacao.Where(Function(s) s.Ativo = True)
            ddlUnidadeComercializacao.Items.Add(New ListItem(un.CodigoUnidade & " - " & un.FatorConversao.ToString("N4"), un.CodigoUnidade))
        Next

        If HIDLinhaProduto.Value = -1 Then
            txtUnitarioFat.Visible = False
            txtQuantidadeFat.Visible = False
            HFC.Value = 1
        Else
            If pProduto.Unidade = objPedido.Itens(HIDLinhaProduto.Value).CodigoUnidadeComercializacao Then
                txtUnitarioFat.Visible = False
                txtQuantidadeFat.Visible = False
            Else
                txtUnitarioFat.Visible = True
                txtQuantidadeFat.Visible = True
            End If
            HFC.Value = objPedido.Itens(HIDLinhaProduto.Value).UnidadeComercializacaoFatorDeConversao
        End If
        SetarMsgRetencao()
    End Sub

    Public Sub SetarMsgRetencao()
        '*************************************************************************************
        '*************** RETENCAO ************************************************************
        '*************************************************************************************
        If HIDLinhaProduto.Value = -1 Then
            Dim objProduto As New Produto(ddlProdutos.SelectedValue)
            Dim Parametros As New OperacaoXEstado
            Parametros.Empresa = Left(objPedido.CodigoEmpresa, 8)
            Parametros.CodigoGrupoProduto = objProduto.CodigoGrupo
            Parametros.CodigoProduto = objProduto.Codigo
            Parametros.CodigoOperacao = objPedido.CodigoOperacao
            Parametros.CodigoSubOperacao = objPedido.CodigoSubOperacao
            Parametros.EstadoOrigem = objPedido.Empresa.CodigoEstado
            Parametros.EstadoDestino = objPedido.Cliente.CodigoEstado
            Dim OxE As New OperacaoXEstado(Parametros)

            If OxE.Encargos.Count = 0 Then Exit Sub

            If objPedido.Cliente.CodigoEstado <> "EX" Then
                If objPedido.Cliente.Codigo.Length = 14 Then
                    If OxE.Encargos.DescRetencaoPJ.Length > 0 Then
                        chkRetencao.Parent.Visible = True
                        chkRetencao.Text = "haverá retenção dos seguintes encargos -> " & OxE.Encargos.DescRetencaoPJ
                        chkRetencao.Checked = True
                    Else
                        chkRetencao.Parent.Visible = False
                        chkRetencao.Checked = False
                    End If
                Else
                    If OxE.Encargos.DescRetencaoPF.Length > 0 Then
                        chkRetencao.Parent.Visible = True
                        chkRetencao.Text = "haverá retenção dos seguintes encargos -> " & OxE.Encargos.DescRetencaoPF
                        chkRetencao.Checked = True
                    Else
                        chkRetencao.Parent.Visible = False
                        chkRetencao.Checked = False
                    End If
                End If
            End If
        Else
            If objPedido.Cliente.CodigoEstado <> "EX" Then
                If objPedido.Cliente.Codigo.Length = 14 Then
                    If objPedido.Itens(HIDLinhaProduto.Value).OperacaoxEstado.Encargos.DescRetencaoPJ.Length > 0 Then
                        chkRetencao.Parent.Visible = True
                        chkRetencao.Text = "haverá retenção dos seguintes encargos -> " & objPedido.Itens(HIDLinhaProduto.Value).OperacaoxEstado.Encargos.DescRetencaoPJ
                        chkRetencao.Checked = objPedido.Itens(HIDLinhaProduto.Value).Retencao
                        chkRetencao.Enabled = Not objPedido.Itens(HIDLinhaProduto.Value).TemNota
                    Else
                        chkRetencao.Parent.Visible = False
                        chkRetencao.Checked = False
                        objPedido.Itens(HIDLinhaProduto.Value).Retencao = False
                    End If
                Else
                    If objPedido.Itens(HIDLinhaProduto.Value).OperacaoxEstado.Encargos.DescRetencaoPF.Length > 0 Then
                        chkRetencao.Parent.Visible = True
                        chkRetencao.Text = "haverá retenção dos seguintes encargos -> " & objPedido.Itens(HIDLinhaProduto.Value).OperacaoxEstado.Encargos.DescRetencaoPF
                        chkRetencao.Checked = objPedido.Itens(HIDLinhaProduto.Value).Retencao
                        chkRetencao.Enabled = Not objPedido.Itens(HIDLinhaProduto.Value).TemNota
                    Else
                        chkRetencao.Parent.Visible = False
                        chkRetencao.Checked = False
                        objPedido.Itens(HIDLinhaProduto.Value).Retencao = False
                    End If
                End If
            End If

        End If
    End Sub

    Protected Sub ddlTipoLancamento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTipoLancamento.SelectedIndexChanged
        Try
            MudouTipodeLancamento(ddlTipoLancamento.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirLancamento_Click(sender As Object, e As EventArgs)
        Try
            SessaoRecuperaPedido()

            Dim prd() As String = ddlProdutos.SelectedValue.Split("-")
            Dim ItemPedido As PedidoXItem
            ItemPedido = objPedido.Itens.Find(Function(s) s.CodigoProduto = prd(0))

            Dim lnkExc As LinkButton = CType(sender, LinkButton)
            Dim linha As Integer = CType(lnkExc.NamingContainer, GridViewRow).RowIndex

            If ItemPedido.Lancamentos(linha).TipoLancamento = eTiposLancamentosPedidos.Normal And ItemPedido.Lancamentos.Count > 1 Then
                MsgBox(Me.Page, "O Lançamento Normal tem que ser o último a ser excluído")
                Exit Sub
            End If

            ItemPedido.Lancamentos.RemoverLancamento(ItemPedido.Lancamentos(linha))

            If ItemPedido.Lancamentos.Count = 0 Then
                objPedido.Itens.Remove(ItemPedido)
                For Each rep In objPedido.Representantes
                    rep.Comissoes.RemoveAll(Function(s) s.CodigoProduto = prd(0))
                Next
                HIDLinhaProduto.Value = -1
            Else
                ItemPedido.Encargos = Nothing
                ItemPedido.Encargos.CriaListar()
            End If


            If objPedido.Itens.Count > 0 Then
                txtResumoQtde.Text = objPedido.Itens(HIDLinhaProduto.Value).Lancamentos.QuantidadeTotalPrdComercializacao.ToString("N4")
                txtResumoUntMedio.Text = objPedido.Itens(HIDLinhaProduto.Value).UnitarioMedioFaturamento.ToString("N10")
                txtResumoValor.Text = objPedido.Itens(HIDLinhaProduto.Value).PedidoValor.ToString("N2")

                gridProdutos.DataSource = ItemPedido.Lancamentos.ToArray
                gridProdutos.DataBind()
                SessaoSalvaPedido()
            Else
                Limpar()
            End If

            If ItemPedido.Lancamentos.Count = 0 Then SetarParametros(HIDLinhaProduto.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
        End Try
    End Sub

    Protected Sub gridProdutos_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridProdutos.RowDataBound
        Try

            If e.Row.RowType = DataControlRowType.DataRow Then
                SessaoRecuperaPedido()

                Dim prd() As String = ddlProdutos.SelectedValue.Split("-")
                Dim ItemPedido As PedidoXItem
                ItemPedido = objPedido.Itens.Find(Function(s) s.CodigoProduto = prd(0))

                Dim lbl As Label = CType(e.Row.FindControl("lblQtdeCom"), Label)
                lbl.Visible = Not ItemPedido.UnidadeFaturamento = ItemPedido.CodigoUnidadeComercializacao

                lbl = CType(e.Row.FindControl("lblUnidCom"), Label)
                lbl.Visible = Not ItemPedido.UnidadeFaturamento = ItemPedido.CodigoUnidadeComercializacao

                lbl = CType(e.Row.FindControl("lblUnitCompCom"), Label)
                lbl.Visible = Not ItemPedido.UnidadeFaturamento = ItemPedido.CodigoUnidadeComercializacao

                lbl = CType(e.Row.FindControl("lblUnitCom"), Label)
                lbl.Visible = Not ItemPedido.UnidadeFaturamento = ItemPedido.CodigoUnidadeComercializacao

                Dim lnk As LinkButton = CType(e.Row.FindControl("lnkExcluirLancamento"), LinkButton)
                lnk.Visible = ItemPedido.Lancamentos(e.Row.RowIndex).IUD = "I"
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            SessaoRecuperaPedido()

            If gridProdutos.Rows.Count = 0 And ddlTipoLancamento.SelectedValue <> "N" Then
                MsgBox(Me.Page, "O Primeiro Lançamento tem que ser Normal")
                Exit Sub
            End If

            Dim prdLancamento As New Produto(ddlProdutos.SelectedValue)
            Dim Valida As ArrayList = ValidaLancamento(prdLancamento)
            If Not Valida(0) Then
                Dim erroMSG As String = Valida(1)
                MsgBox(Me.Page, erroMSG)
                Exit Sub
            End If

            '**************************************************************************************
            '*************  ITEM DO PEDIDO ********************************************************
            '**************************************************************************************
            Dim ItemPedido As [Lib].Negocio.PedidoXItem
            If HIDLinhaProduto.Value = -1 Then
                ItemPedido = New [Lib].Negocio.PedidoXItem(objPedido)
                ItemPedido.IUD = "I"
                ItemPedido.CodigoOperacaoXEstado = Valida(2)
                ItemPedido.Retencao = chkRetencao.Checked

                ItemPedido.CodigoProduto = ddlProdutos.SelectedValue
                ItemPedido.Descricao = prdLancamento.Nome
                ItemPedido.CodigoClassificacao = ddlClassificacao.SelectedValue
                ItemPedido.CodigoUnidadeComercializacao = ddlUnidadeComercializacao.SelectedValue

                objPedido.Itens.Add(ItemPedido)
                HIDLinhaProduto.Value = objPedido.Itens.FindIndex(Function(s) s.CodigoProduto = ItemPedido.CodigoProduto)

                For Each rep In objPedido.Representantes.Where(Function(s) s.IUD <> "D" And Not s.PercentualFixo)
                    If rep.ValorComissao = 0 Then
                        Dim PxT As New PedidoXRepresentantesxTabelaDeComissao(rep, rep.Comissoes.TabelasSafra)
                        PxT.IUD = "I"
                        PxT.CodigoProduto = ItemPedido.CodigoProduto
                        PxT.NomeProduto = ItemPedido.Descricao
                        rep.Comissoes.Add(PxT)
                    End If
                Next
            Else
                ItemPedido = objPedido.Itens(HIDLinhaProduto.Value)
                ItemPedido.IUD = "U"
            End If

            '**************************************************************************************
            '*************  LANCAMENTOS DOS ITENS DO PEDIDO ***************************************
            '**************************************************************************************
            Dim Lancamento As New [Lib].Negocio.LancamentoItemPedido(ItemPedido)

            Lancamento.IUD = "I"
            If objPedido.Itens(HIDLinhaProduto.Value).Lancamentos.Count = 0 Then
                Lancamento.CodigoPedidoItem = 1
            Else
                Lancamento.CodigoPedidoItem = objPedido.Itens(HIDLinhaProduto.Value).Lancamentos.Max(Function(s) s.CodigoPedidoItem) + 1
            End If
            Lancamento.TipoLancamento = IIf(ddlTipoLancamento.SelectedValue = "N", eTiposLancamentosPedidos.Normal, (IIf(Left(ddlTipoLancamento.SelectedValue, 1) = "C", eTiposLancamentosPedidos.Complemento, eTiposLancamentosPedidos.Estorno)))
            Lancamento.Movimento = CDate(txtDataMovimentoItem.Text).ToShortDateString()
            Lancamento.DataEntrega = CDate(txtDataEntregaItem.Text).ToShortDateString()
            Lancamento.UsuarioLiberacao = HttpContext.Current.Session("ssNomeUsuario")

            '********** UNITARIO DE COMPRA ****************************************
            If IsNumeric(HFMediaOficialCompra.Value) AndAlso CDec(HFMediaOficialCompra.Value) > 0 Then
                Lancamento.UnitarioOficialCompra = HFMediaOficialCompra.Value
                Lancamento.UnitarioMoedaCompra = HFMediaMoedaCompra.Value
            Else
                Lancamento.UnitarioOficialCompra = 0
                Lancamento.UnitarioMoedaCompra = 0
            End If
            '**********************************************************************

            If ItemPedido.CodigoUnidadeComercializacao <> ItemPedido.Produto.Unidade Then
                Lancamento.QuantidadeFaturamento = CDec(txtQuantidadeFat.Text)
                'Lancamento.QuantidadeFaturamento = CDec(txtQuantidade.Text)
                Lancamento.QuantidadeComercializacao = CDec(txtQuantidade.Text)

                If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    If CDec(txtUnitarioFat.Text) > 0 Then
                        Lancamento.UnitarioOficial = CDec(txtUnitarioFat.Text)
                    Else
                        Lancamento.UnitarioOficial = CDec(txtUnitario.Text)
                    End If

                    Lancamento.TotalOficial = CDec(txtValorTotal.Text)

                    If ddlTipoLancamento.SelectedValue = "N" AndAlso objPedido.SubOperacao.UnitarioPedido AndAlso Lancamento.UnitarioOficial = 0 Then
                        MsgBox(Me.Page, "O valor unitário do pedido deve ser maior que zero! ")
                        Exit Sub
                    End If
                Else
                    If CDec(txtUnitarioFat.Text) > 0 Then
                        Lancamento.UnitarioMoeda = CDec(txtUnitarioFat.Text)
                    Else
                        Lancamento.UnitarioMoeda = CDec(txtUnitario.Text)
                    End If

                    'Caso seja moeda estrangeira e seja Complemento ou Estorno de Valor e o usuário tenha permissão
                    'ele conseguirá lançar os valores em reais e em moeda estrangeira.
                    If Funcoes.VerificaPermissao("PedidoXItemXLancamentoMoeda", "Gravar") AndAlso (ddlTipoLancamento.SelectedValue = "CV" Or ddlTipoLancamento.SelectedValue = "EV") Then
                        Lancamento.TotalMoeda = CDec(txtValorTotalMoeda.Text)
                    Else
                        Lancamento.TotalMoeda = CDec(txtValorTotal.Text)
                    End If

                    If ddlTipoLancamento.SelectedValue = "N" AndAlso objPedido.SubOperacao.UnitarioPedido AndAlso Lancamento.UnitarioMoeda = 0 Then
                        MsgBox(Me.Page, "O valor unitário do pedido deve ser maior que zero! ")
                        Exit Sub
                    End If
                End If

            Else
                Lancamento.QuantidadeFaturamento = CDec(txtQuantidade.Text)
                Lancamento.QuantidadeComercializacao = CDec(txtQuantidade.Text)

                If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    Lancamento.UnitarioOficial = CDec(txtUnitario.Text)
                    Lancamento.TotalOficial = CDec(txtValorTotal.Text)

                    If ddlTipoLancamento.SelectedValue = "N" AndAlso objPedido.SubOperacao.UnitarioPedido AndAlso Lancamento.UnitarioOficial = 0 Then
                        MsgBox(Me.Page, "O valor unitário do pedido deve ser maior que zero! ")
                        Exit Sub
                    End If
                Else
                    Lancamento.UnitarioMoeda = CDec(txtUnitario.Text)
                    'Lancamento.TotalMoeda = CDec(txtValorTotal.Text)
                    'Caso seja moeda estrangeira e seja Complemento ou Estorno de Valor e o usuário tenha permissão
                    'ele conseguirá lançar os valores em reais e em moeda estrangeira.
                    If Funcoes.VerificaPermissao("PedidoXItemXLancamentoMoeda", "Gravar") AndAlso (ddlTipoLancamento.SelectedValue = "CV" Or ddlTipoLancamento.SelectedValue = "EV") Then
                        Lancamento.TotalMoeda = CDec(txtValorTotalMoeda.Text)
                    Else
                        Lancamento.TotalMoeda = CDec(txtValorTotal.Text)
                    End If

                    If ddlTipoLancamento.SelectedValue = "N" AndAlso objPedido.SubOperacao.UnitarioPedido AndAlso Lancamento.UnitarioMoeda = 0 Then
                        MsgBox(Me.Page, "O valor unitário do pedido deve ser maior que zero! ")
                        Exit Sub
                    End If
                End If
            End If

            If objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then

                If Lancamento.UnitarioMoeda > 0 Then
                    If objPedido.IndiceFixado Then
                        Lancamento.UnitarioOficial = Funcoes.ConverteMoeda(Lancamento.UnitarioMoeda, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 10)
                    Else
                        Lancamento.UnitarioOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(Lancamento.UnitarioMoeda, objPedido.Indexador.Codigo, Lancamento.Movimento), 10, MidpointRounding.AwayFromZero)
                    End If
                End If


                If objPedido.IndiceFixado > 0 Then
                    If Funcoes.VerificaPermissao("PedidoXItemXLancamentoMoeda", "Gravar") AndAlso (ddlTipoLancamento.SelectedValue = "CV" Or ddlTipoLancamento.SelectedValue = "EV") Then
                        Lancamento.TotalOficial = CDec(txtValorTotal.Text)
                    Else
                        Lancamento.TotalOficial = Funcoes.ConverteMoeda(Lancamento.TotalMoeda, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 10)
                    End If

                Else
                    If Funcoes.VerificaPermissao("PedidoXItemXLancamentoMoeda", "Gravar") AndAlso (ddlTipoLancamento.SelectedValue = "CV" Or ddlTipoLancamento.SelectedValue = "EV") Then
                        Lancamento.TotalOficial = CDec(txtValorTotal.Text)
                    Else
                        Lancamento.TotalOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(Lancamento.TotalMoeda, objPedido.Indexador.Codigo, Lancamento.Movimento), 10, MidpointRounding.AwayFromZero)
                    End If
                End If
            End If

            '**********************************************************************************************************************************************************************************************************************************
            '**********************************************************************************************************************************************************************************************************************************
            '**********************************************************************************************************************************************************************************************************************************

            Dim ret As ArrayList = objPedido.Itens(HIDLinhaProduto.Value).Lancamentos.AdicionarLancamento(Lancamento)

            If Not ret(0) Then
                MsgBox(Me.Page, ret(1))
            Else
                Dim MediaCompra As Decimal

                If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                    MediaCompra = CDec(HFMediaOficialCompra.Value)
                Else
                    MediaCompra = CDec(HFMediaMoedaCompra.Value)
                End If

                'Cometado, quando o processo da Onsoft estiver no AR vamos rever - Furlan - 07/05/2024
                'If MediaCompra > 0 _
                '   AndAlso (objPedido.Itens(HIDLinhaProduto.Value).QuantidadePedidoFaturamento > 0 AndAlso objPedido.Itens(HIDLinhaProduto.Value).PedidoValor > 0) _
                '   AndAlso MediaCompra > objPedido.Itens(HIDLinhaProduto.Value).UnitarioMedioFaturamento _
                '   AndAlso ([Enum].Parse(GetType(eClassesOperacoes), objPedido.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.VENDAS) _
                '   AndAlso objPedido.Itens(HIDLinhaProduto.Value).SaldoItem.SaldoQtdeGlobalFiscal + objPedido.Itens(HIDLinhaProduto.Value).SaldoItem.SaldoQtdeDiretoFiscal > 0 Then

                '    If Not UsuarioSupervisor("PRECOUNITARIOMENORDOQUECOMPRA") Is Nothing Then
                '        Lancamento.UsuarioLiberacao = UsuarioSupervisor("PRECOUNITARIOMENORDOQUECOMPRA").Usuario_Id
                '    Else
                '        objPedido.Itens(HIDLinhaProduto.Value).Lancamentos.RemoverLancamento(Lancamento)
                '        MsgBox(Me.Page, "O Preço Unitário definido não pode ser menor que o Preço Unitário Médio de Compra! - Abrirá uma tela para que um Supervisor Valide essa Operação!")
                '        ValidaSupervisor("PRECOUNITARIOMENORDOQUECOMPRA")
                '        Exit Sub
                '    End If
                'End If

                If Lancamento.ItemPedido.Produto.Agrupar.Equals("S") Then
                    lnkProdutoNovo.Parent.Visible = True
                Else
                    lnkProdutoNovo.Parent.Visible = False
                End If

                txtResumoQtde.Text = objPedido.Itens(HIDLinhaProduto.Value).Lancamentos.QuantidadeTotalPrdComercializacao.ToString("N4")
                txtResumoUntMedio.Text = objPedido.Itens(HIDLinhaProduto.Value).UnitarioMedioFaturamento.ToString("N10")
                txtResumoValor.Text = objPedido.Itens(HIDLinhaProduto.Value).PedidoValor.ToString("N2")

                ItemPedido.Encargos = Nothing
                ItemPedido.Encargos.CriaListar()

                '#AcrescimoDesconto
                If ItemPedido.Encargos.Find(Function(s) s.CodigoEncargo = "ACRESCIMOS") IsNot Nothing AndAlso ItemPedido.Acrescimo > 0 Then
                    ItemPedido.Encargos.Find(Function(s) s.CodigoEncargo = "ACRESCIMOS").ValorOficial = ItemPedido.Acrescimo
                    ItemPedido.Encargos.AtualizaLiquido()
                End If

                If ItemPedido.Encargos.Find(Function(s) s.CodigoEncargo = "DESCONTOS") IsNot Nothing AndAlso ItemPedido.Desconto > 0 Then
                    ItemPedido.Encargos.Find(Function(s) s.CodigoEncargo = "DESCONTOS").ValorOficial = ItemPedido.Desconto
                    ItemPedido.Encargos.AtualizaLiquido()
                End If
                '##

                ItemPedido.Pedido.Representantes.RecalcularComissoesFixas()

                Limpar()
            End If

            If (ddlTipoLancamento.SelectedValue = "EQ" OrElse ddlTipoLancamento.SelectedValue = "EV" OrElse ddlTipoLancamento.SelectedValue = "CQ" OrElse ddlTipoLancamento.SelectedValue = "CV") Then
                If objPedido.FiscalAberto Then
                    CType(Me.Page, PedidosXItens).HabilitarAlteracao()
                ElseIf Funcoes.VerificaAcesso(objPedido.Empresa.Codigo, objPedido.Empresa.CodigoEndereco, IIf(objPedido.DataVencimentoPedido > Now.Date, Now.Date, objPedido.DataVencimentoPedido), "PEDIDOS") Then
                    CType(Me.Page, PedidosXItens).HabilitarAlteracao()
                End If
            End If


            SessaoSalvaPedido()
            SetarParametros(HIDLinhaProduto.Value)
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao Inserir o Lançamento." & Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString))
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub LnkSair_Click(sender As Object, e As EventArgs) Handles LnkSair.Click
        Try
            If objPedido Is Nothing Then SessaoRecuperaPedido()

            If TypeOf Me.Page Is PedidosXItens Then
                CType(Me.Page, PedidosXItens).AtualizarResumoItens()
            End If
            Popup.CloseDialog(Me.Page, "divPedidoLancamentoItem")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBuscaProduto_Click(sender As Object, e As EventArgs) Handles lnkBuscaProduto.Click
        Try
            Dim ucConsultaProduto As ucConsultaProduto = CType(Me.Page.FindControlRecursive("ucConsultaProduto"), ucConsultaProduto)
            If ucConsultaProduto IsNot Nothing Then
                ucConsultaProduto.Limpar()
                ucConsultaProduto.MainUserControl = Me
                Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
                Popup.ConsultaDeProduto(Me.Page, "objProdutoPXI" & HID.Value, txtNome.ClientID, True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtDataEntregaItem_TextChanged(sender As Object, e As EventArgs) Handles txtDataEntregaItem.TextChanged
        Try
            MudouDataEntrega()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkProdutoNovo_Click(sender As Object, e As EventArgs) Handles lnkProdutoNovo.Click
        Try
            SetarParametros(-1)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeComercializacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeComercializacao.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()
            HFC.Value = ddlUnidadeComercializacao.SelectedItem.Text.Split("-")(1)

            If lblUnidFat.Text.Trim <> ddlUnidadeComercializacao.SelectedItem.Text.Split("-")(0).Trim Then
                txtUnitarioFat.Visible = True
                txtQuantidadeFat.Visible = True

                txtUnitarioFat.Text = "0,0000000000"

                txtUnitario.Text = "0,0000000000"
                txtQuantidade.Text = "0,0000"
                txtQuantidadeFat.Focus()
                txtQuantidadeFat.Text &= " "
            Else
                txtUnitarioFat.Text = "0,0000000000"
                txtQuantidadeFat.Text = "0,0000"
                txtUnitarioFat.Visible = False
                txtQuantidadeFat.Visible = False
                txtQuantidade.Focus()
                txtQuantidade.Text &= " "
            End If

            Dim umc As Decimal
            If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                umc = CDec(HFMediaOficialCompra.Value)
                txtUnitarioCompra.Text = (umc * HFC.Value).ToString("N10")
            Else
                umc = CDec(HFMediaMoedaCompra.Value)
                txtUnitarioCompra.Text = (umc * HFC.Value).ToString("N10")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkRetencao_CheckedChanged(sender As Object, e As EventArgs) Handles chkRetencao.CheckedChanged
        If HIDLinhaProduto.Value >= 0 Then
            SessaoRecuperaPedido()
            If objPedido.IUD <> "I" And objPedido.Itens(HIDLinhaProduto.Value).IUD <> "I" Then objPedido.Itens(HIDLinhaProduto.Value).IUD = "U"
            objPedido.Itens(HIDLinhaProduto.Value).Retencao = chkRetencao.Checked
            objPedido.Itens(HIDLinhaProduto.Value).Encargos = Nothing
            SessaoSalvaPedido()
        End If
    End Sub

    Protected Sub txtValorTotalMoeda_TextChanged(sender As Object, e As EventArgs) Handles txtValorTotalMoeda.TextChanged
        SessaoRecuperaPedido()
        'Caso o pedido seja em moeda estrangeira e o valor em Reais seja igual a zero então é feito o cálculo do valor em reais correspondente ao em  dólar lançado
        'Utilizando o índice do pedido
        If objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso Not String.IsNullOrWhiteSpace(txtValorTotalMoeda.Text) _
            AndAlso IsNumeric(txtValorTotalMoeda.Text) AndAlso CDec(txtValorTotalMoeda.Text) > 0 _
            AndAlso IsNumeric(txtValorTotal.Text) AndAlso CDec(txtValorTotal.Text) <= 0 Then
            txtValorTotal.Text = Funcoes.ConverteMoeda(txtValorTotalMoeda.Text, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 10).ToString("N2")
        End If
    End Sub


End Class