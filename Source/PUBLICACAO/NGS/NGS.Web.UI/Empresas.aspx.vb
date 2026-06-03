Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Empresas
    Inherits BasePage

    Dim Mensagem As String
    Dim objEmpresa As ClienteXEmpresa

#Region "Session"
    Private Sub SessaoSalvaEmpresa()
        Session("objEmpresa" + HID.Value) = objEmpresa
    End Sub

    Private Sub SessaoRecuperaEmpresa()
        objEmpresa = CType(Session("objEmpresa" + HID.Value), [Lib].Negocio.ClienteXEmpresa)
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Empresas", "ACESSAR") Then

                ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas)
                ddl.Carregar(ddlServidor, CarregarDDL.Tabela.Servidores)
                ddl.Carregar(ddlEstadoExpCRC, CarregarDDL.Tabela.EstadosUF)
                ddl.Carregar(ddlNaturezaJuridica, CarregarDDL.Tabela.NaturezaJuridica)
                ddl.Carregar(ddlCnae, CarregarDDL.Tabela.CNAE)
                CarregarIND_ALIQ_CSLL()

                Dim encs As New ListEncargo(True, " isnull(VerificaEmpresa,0) = 1")
                If encs.DescEncargoEmpresa.Length > 0 Then
                    chkObgEncargo.Text = "Quando Parametrizado nas operacões de (VENDAS,TRANSFERENCIAS,DOACAO,FISCAL), obriga indiferentemente do tipo de pessoa a destacar nas notas o(s) encargo(s): " & encs.DescEncargoEmpresa
                Else
                    chkObgEncargo.Checked = False
                    chkObgEncargo.Parent.Visible = False
                End If

                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Not Session("ContaInicial" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaInicial" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaInicial.Text = ct.Conta
            txtContaInicial.ToolTip = ct.Titulo
            Session.Remove("ContaInicial" & HID.Value)
        ElseIf Not Session("ContaFinal" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaFinal" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaFinal.Text = ct.Conta
            txtContaFinal.ToolTip = ct.Titulo
            Session.Remove("ContaFinal" & HID.Value)
        ElseIf Not Session("ContaPatrimonio" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaPatrimonio" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaPatrimonio.Text = ct.Conta
            txtContaPatrimonio.ToolTip = ct.Titulo
            Session.Remove("ContaPatrimonio" & HID.Value)
        ElseIf Not Session("ContaAtiva" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaAtiva" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaVariacaoAtiva.Text = ct.Conta
            txtContaVariacaoAtiva.ToolTip = ct.Titulo
            Session.Remove("ContaAtiva" & HID.Value)
        ElseIf Not Session("ContaPassiva" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaPassiva" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaVariacaoPassiva.Text = ct.Conta
            txtContaVariacaoPassiva.ToolTip = ct.Titulo
            Session.Remove("ContaPassiva" & HID.Value)
        ElseIf Not Session("ContaCliente" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaCliente" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaVariacaoCliente.Text = ct.Conta
            txtContaVariacaoCliente.ToolTip = ct.Titulo
            Session.Remove("ContaCliente" & HID.Value)
        ElseIf Not Session("ContaFornecedor" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaFornecedor" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaVariacaoFornecedor.Text = ct.Conta
            txtContaVariacaoFornecedor.ToolTip = ct.Titulo
            Session.Remove("ContaFornecedor" & HID.Value)
        ElseIf Not Session("ContaBanco" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaBanco" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaContabilGrupoBanco.Text = ct.Conta
            txtContaContabilGrupoBanco.ToolTip = ct.Titulo
            Session.Remove("ContaBanco" & HID.Value)
        ElseIf Not Session("ContaJuroRecebido" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaJuroRecebido" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaJuroRecebido.Text = ct.Conta
            txtContaJuroRecebido.ToolTip = ct.Titulo
            Session.Remove("ContaJuroRecebido" & HID.Value)
        ElseIf Not Session("ContaJuroPago" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaJuroPago" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaJuroPago.Text = ct.Conta
            txtContaJuroPago.ToolTip = ct.Titulo
            Session.Remove("ContaJuroPago" & HID.Value)
        ElseIf Not Session("ContaDuplicatasDescontada" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaDuplicatasDescontada" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaDuplicatasDescontada.Text = ct.Conta
            txtContaDuplicatasDescontada.ToolTip = ct.Titulo
            Session.Remove("ContaDuplicatasDescontada" & HID.Value)
        ElseIf Not Session("ContaEstoque" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaEstoque" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaEstoque.Text = ct.Conta
            txtContaEstoque.ToolTip = ct.Titulo
            Session.Remove("ContaEstoque" & HID.Value)

        ElseIf Not Session("ContaEstoqueEmNossoPoder" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaEstoqueEmNossoPoder" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaEstoqueEmNossoPoder.Text = ct.Conta
            txtContaEstoqueEmNossoPoder.ToolTip = ct.Titulo
            Session.Remove("ContaEstoqueEmNossoPoder" & HID.Value)

        ElseIf Not Session("ContaEstoqueEmPoderDeTerceiros" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaEstoqueEmPoderDeTerceiros" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaEstoqueEmPoderDeTerceiros.Text = ct.Conta
            txtContaEstoqueEmPoderDeTerceiros.ToolTip = ct.Titulo
            Session.Remove("ContaEstoqueEmPoderDeTerceiros" & HID.Value)

        ElseIf Not Session("ContaGrupoComissoes" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaGrupoComissoes" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaGrupoComissoes.Text = ct.Conta
            txtContaGrupoComissoes.ToolTip = ct.Titulo
            Session.Remove("ContaGrupoComissoes" & HID.Value)
        ElseIf Not Session("ContaAdiantamentoDeFrete" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaAdiantamentoDeFrete" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaAdiantamentoDeFrete.Text = ct.Conta
            txtContaAdiantamentoDeFrete.ToolTip = ct.Titulo
            Session.Remove("ContaAdiantamentoDeFrete" & HID.Value)
        ElseIf Not Session("ContaPedagioDeFrete" & HID.Value) Is Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaPedagioDeFrete" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaPedagioDeFrete.Text = ct.Conta
            txtContaPedagioDeFrete.ToolTip = ct.Titulo
            Session.Remove("ContaPedagioDeFrete" & HID.Value)
        ElseIf Session("objProdutoDeFrete" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoDeFrete" & HID.Value)
            txtProdutoDeFrete.Text = objProduto.Codigo
            Session.Remove("objProdutoDeFrete" & HID.Value)
        ElseIf Session("objProdutoDeEstadia" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoDeEstadia" & HID.Value)
            txtProdutoDeEstadia.Text = objProduto.Codigo
            Session.Remove("objProdutoDeEstadia" & HID.Value)
        ElseIf Session("objProdutoDeMDFe" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoDeMDFe" & HID.Value)
            txtProdutoDeMDFe.Text = objProduto.Codigo
            Session.Remove("objProdutoDeMDFe" & HID.Value)
        ElseIf Session("ContaGrupoTEDDOC" & HID.Value) IsNot Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaGrupoTEDDOC" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaTEDDOC.Text = ct.Conta
            txtContaTEDDOC.ToolTip = ct.Titulo
            Session.Remove("ContaGrupoTEDDOC" & HID.Value)
        ElseIf Session("ContaDescontoObtido" & HID.Value) IsNot Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaDescontoObtido" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaDescontoObtido.Text = ct.Conta
            txtContaDescontoObtido.ToolTip = ct.Titulo
            Session.Remove("ContaDescontoObtido" & HID.Value)
        ElseIf Session("ContaDescontoConcedido" & HID.Value) IsNot Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ContaDescontoConcedido" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtContaDescontoConcedido.Text = ct.Conta
            txtContaDescontoConcedido.ToolTip = ct.Titulo
            Session.Remove("ContaDescontoConcedido" & HID.Value)
        ElseIf Session("FornecedorDeFrete" & HID.Value) IsNot Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("FornecedorDeFrete" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtFornecedorFrete.Text = ct.Conta
            btnFornecedorFrete.ToolTip = ct.Titulo
            Session.Remove("FornecedorDeFrete" & HID.Value)
        ElseIf Session("ssCaixaCompensacao" & HID.Value) IsNot Nothing Then
            Dim ct As [Lib].Negocio.PlanoDeConta = CType(Session("ssCaixaCompensacao" & HID.Value), [Lib].Negocio.PlanoDeConta)
            txtCaixaCompensacao.Text = ct.Conta
            btnCaixaCompensacao.ToolTip = ct.Titulo
            Session.Remove("ssCaixaCompensacao" & HID.Value)
        End If
    End Sub

    Private Sub Limpar()
        DdlEmpresa.SelectedIndex = 0

        txtRegistroNaJunta.Text = ""
        txtEstadoDaJunta.Text = ""
        txtDataDaFundacao.Text = ""

        txtNire.Text = ""
        txtDataDaNire.Text = ""

        txtInscricaoMunicipal.Text = ""
        txtAtividadeEconomicaEstadual.Text = ""
        txtAtividadeEconomicaFederal.Text = ""
        txtMatriz.Text = ""
        txtRegistroEC.Text = ""
        txtRegistroEP.Text = ""
        txtRegistroEI.Text = ""

        DdlCodigoDaQualificacaoDoTitular.SelectedIndex = 0

        txtNomeDoContador.Text = ""
        txtCpfDoContador.Text = ""
        txtCnpjDoContador.Text = ""
        txtCrcDoContador.Text = ""
        ddlEstadoExpCRC.SelectedIndex = 0
        txtEmailContador.Text = ""
        txtTelefoneContador.Text = ""

        txtQualificacaoDoContador.Text = ""
        DdlCodigoDeQualificacaoDoContador.SelectedIndex = 0
        ddlCrt.SelectedValue = ""

        txtNomeDoTitular.Text = ""
        txtCpfDoTitular.Text = ""
        txtQualificacaoDoTitular.Text = ""
        txtContaInicial.Text = ""
        txtContaFinal.Text = ""
        txtContaVariacaoAtiva.Text = ""
        txtContaVariacaoPassiva.Text = ""
        txtContaVariacaoCliente.Text = ""
        txtContaVariacaoFornecedor.Text = ""
        txtContaContabilGrupoBanco.Text = ""
        txtContaEstoque.Text = ""
        txtContaGrupoComissoes.Text = ""
        txtContaDescontoConcedido.Text = String.Empty
        txtContaDescontoObtido.Text = String.Empty
        txtContaJuroPago.Text = String.Empty
        txtContaJuroRecebido.Text = String.Empty
        txtContaTEDDOC.Text = String.Empty
        txtContaDuplicatasDescontada.Text = String.Empty

        txtContaAdiantamentoDeFrete.Text = ""
        txtContaPedagioDeFrete.Text = ""
        txtFornecedorFrete.Text = String.Empty
        txtCaixaCompensacao.Text = String.Empty

        txtContaEstoqueEmNossoPoder.Text = String.Empty
        txtContaEstoqueEmPoderDeTerceiros.Text = String.Empty

        txtProdutoDeFrete.Text = ""
        txtProdutoDeEstadia.Text = ""
        txtProdutoDeMDFe.Text = ""

        chknfe.Checked = False
        chkEmitirCTe.Checked = False
        chknfProdutor.Checked = False
        chkObrigaNavio.Checked = False
        chkObrigaChaveNf.Checked = False
        chkObrigaChaveNfg.Checked = False
        chkCertidaoNegativa.Checked = False
        chkNossaEmissao.Checked = False
        chkFluxoDeCaixa.Checked = False
        chkSugereDeposito.Checked = False
        chkRegistroDeExportacao.Checked = False
        chkFretePedido.Checked = False
        chkPedidoBloqueado.Checked = False
        chkLiberaCarregamento.Checked = False
        chkBaixaFinanceiroPorLote.Checked = False
        chkUsarDescricaoProduto.Checked = False
        chkUsarRegistroMinAgr.Checked = False
        chkObgEncargo.Checked = False

        ddlServidor.SelectedIndex = 0

        txtPrazoCancelamentoNFE.Text = ""
        txtNumeroDeViasNFE.Text = ""

        RadControlaDataMOvimentoNFGNao.Checked = True
        RadControlaDataMOvimentoNFGSim.Checked = False
        RadControlaEmissaChequeNao.Checked = True
        RadControlaEmissaChequeSim.Checked = False
        radArquivoNFENao.Checked = True
        radArquivoNFESim.Checked = False
        radConferenciaNFENao.Checked = True
        radConferenciaNFESim.Checked = False

        txtObservacaoSefazNFE.Text = ""
        lblDownloadNFe.Text = String.Empty
        lnkExcluir.Parent.Visible = False

        ddlUsuarios.Items.Clear()
        ddlUsuarios.Items.Add("Inc.- " & Session("ssNomeUsuario"))

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPlanoDeContas.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)
    End Sub

    Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(txtRegistroNaJunta.Text) Then
            Mensagem &= "Registro na Junta é obrigatório! \n"
        End If
        If String.IsNullOrWhiteSpace(txtDataDaFundacao.Text) Then
            Mensagem &= "Data da Fundação é obrigatório! \n"
        End If
        If String.IsNullOrWhiteSpace(txtEstadoDaJunta.Text) Then
            Mensagem &= "Estado da Junta é obrigatório! \n"
        End If
        If String.IsNullOrWhiteSpace(txtNire.Text) Then
            Mensagem &= "Numero da Nire é obrigatório! \n"
        End If
        If String.IsNullOrWhiteSpace(txtDataDaNire.Text) Then
            Mensagem &= "Data da Nire é obrigatório! \n"
        End If
        If String.IsNullOrWhiteSpace(txtInscricaoMunicipal.Text) Then
            Mensagem &= "Inscrição Municipal é obrigatório! \n"
        End If
        If String.IsNullOrWhiteSpace(txtAtividadeEconomicaEstadual.Text) Then
            Mensagem &= "Atividade Economica Estadual é obrigatório! \n"
        End If
        If String.IsNullOrWhiteSpace(txtAtividadeEconomicaFederal.Text) Then
            Mensagem &= "Atividade Economica Federal é obrigatório! \n"
        End If
        If String.IsNullOrWhiteSpace(txtMatriz.Text) Then
            Mensagem &= "Matriz S/N é obrigatório! \n"
        End If
        If String.IsNullOrWhiteSpace(DdlCodigoDeRelacionamento.Text) Then
            Mensagem = "Código de Relacionamento é obrigatório! \n"
            MsgBox(Me.Page, Mensagem)
            Return False
        End If
        If String.IsNullOrWhiteSpace(DdlEscrituracaoContabil.Text) Then
            Mensagem = "Forma de Escrituração Contábil é obrigatório! \n"
            MsgBox(Me.Page, Mensagem)
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtNomeDoTitular.Text) Then
            Mensagem = "Nome do Titular é obrigatório! \n"
            MsgBox(Me.Page, Mensagem)
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtCpfDoTitular.Text) Then
            Mensagem = "CPF do Titular é obrigatório! \n"
            MsgBox(Me.Page, Mensagem)
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtQualificacaoDoTitular.Text) Then
            Mensagem = "Qualificação do Titular é obrigatório! \n"
            MsgBox(Me.Page, Mensagem)
            Return False
        End If
        If String.IsNullOrWhiteSpace(DdlCodigoDaQualificacaoDoTitular.Text) Then
            Mensagem = "Código de Qualificação do Titular é obrigatório! \n"
            MsgBox(Me.Page, Mensagem)
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtNomeDoContador.Text) Then
            Mensagem = "Nome do Contador é obrigatório! \n"
            MsgBox(Me.Page, Mensagem)
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtCpfDoContador.Text) AndAlso String.IsNullOrWhiteSpace(txtCnpjDoContador.Text) Then
            Mensagem = "CPF ou CNPJ do Contador é obrigatório! \n"
            MsgBox(Me.Page, Mensagem)
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtCrcDoContador.Text) Then
            Mensagem = "CRC do Contador é obrigatório! \n"
        End If

        If String.IsNullOrWhiteSpace(txtQualificacaoDoContador.Text) Then
            Mensagem = "Qualificação do Contador é obrigatório! \n"
        End If
        If String.IsNullOrWhiteSpace(DdlCodigoDeQualificacaoDoContador.Text) Then
            Mensagem = "Código de Qualificação do Contador é obrigatório! \n"
        End If
        If String.IsNullOrWhiteSpace(ddlCrt.Text) Then
            Mensagem = "Código de Regime Tributário é obrigatório! \n"
        End If

        If String.IsNullOrWhiteSpace(txtPrazoCancelamentoNFE.Text) And chknfe.Checked Then
            Mensagem = " Prazo para canelamento da NFe é obrigatório! \n"
        End If

        If Not String.IsNullOrEmpty(Mensagem) Then
            MsgBox(Me.Page, Mensagem)
            Return False
        Else
            Return True
        End If
    End Function

    Function ValidarExcluir()
        If DdlEmpresa.SelectedValue = "" Then
            MsgBox(Me.Page, "Empresa é obrigatório!")
            Return False
        End If
        Return True
    End Function

#Region "Consulta Contas Contabeis"
    Protected Sub btnContaInicial_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaInicial" & HID.Value)
    End Sub

    Protected Sub btnContaFinal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnContaFinal.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaFinal" & HID.Value)
    End Sub

    Protected Sub btnContaPatrimonio_Click(sender As Object, e As EventArgs) Handles btnContaPatrimonio.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaPatrimonio" & HID.Value)
    End Sub

    Protected Sub btnContaAtiva_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaAtiva" & HID.Value)
    End Sub

    Protected Sub btnContaPassiva_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaPassiva" & HID.Value)
    End Sub

    Protected Sub btnContaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaCliente" & HID.Value)
    End Sub

    Protected Sub btnContaFornecedor_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaFornecedor" & HID.Value)
    End Sub

    Protected Sub btnContaBanco_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnContaBanco.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaBanco" & HID.Value)
    End Sub

    Protected Sub btnContaJuroRecebido_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnContaJuroRecebido.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaJuroRecebido" & HID.Value)
    End Sub

    Protected Sub btnContaJuroPago_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnContaJuroPago.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaJuroPago" & HID.Value)
    End Sub

    Protected Sub btnContaDescontoObtido_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnContaDescontoObtido.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaDescontoObtido" & HID.Value)
    End Sub

    Protected Sub btnContaDescontoConcedido_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnContaDescontoConcedido.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaDescontoConcedido" & HID.Value)
    End Sub

    Protected Sub btnContaDuplicatasDescontada_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnContaDuplicatasDescontada.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaDuplicatasDescontada" & HID.Value)
    End Sub

    Protected Sub btnContaGrupoComissoes_Click(sender As Object, e As EventArgs) Handles btnContaGrupoComissoes.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaGrupoComissoes" & HID.Value)
    End Sub

    Protected Sub btnContaTEDDOC_Click(sender As Object, e As EventArgs) Handles btnContaTEDDOC.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaGrupoTEDDOC" & HID.Value)
    End Sub

    Protected Sub btnContaAdiantamentoDeFrete_Click(sender As Object, e As EventArgs) Handles btnContaAdiantamentoDeFrete.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaAdiantamentoDeFrete" & HID.Value)
    End Sub

    Protected Sub btnContaPedagioDeFrete_Click(sender As Object, e As EventArgs) Handles btnContaPedagioDeFrete.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaPedagioDeFrete" & HID.Value)
    End Sub

    Protected Sub btnProdutoDeFrete_Click(sender As Object, e As EventArgs) Handles btnProdutoDeFrete.Click
        ucConsultaProduto.Limpar()
        Session("Where" & HID.Value) = "Situacao = 1"
        Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
        Popup.ConsultaDeProduto(Me.Page, "objProdutoDeFrete" & HID.Value, txtNome.ClientID, True)
    End Sub

    Protected Sub btnProdutoDeEstadia_Click(sender As Object, e As EventArgs) Handles btnProdutoDeEstadia.Click
        ucConsultaProduto.Limpar()
        Session("Where" & HID.Value) = "Situacao = 1"
        Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
        Popup.ConsultaDeProduto(Me.Page, "objProdutoDeEstadia" & HID.Value, txtNome.ClientID, True)
    End Sub

    Protected Sub btnProdutoDeMDFe_Click(sender As Object, e As EventArgs) Handles btnProdutoDeMDFe.Click
        ucConsultaProduto.Limpar()
        Session("Where" & HID.Value) = "Situacao = 1"
        Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
        Popup.ConsultaDeProduto(Me.Page, "objProdutoDeMDFe" & HID.Value, txtNome.ClientID, True)
    End Sub

    Protected Sub btnFornecedorFrete_Click(sender As Object, e As EventArgs) Handles btnFornecedorFrete.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "FornecedorDeFrete" & HID.Value)
    End Sub

    Protected Sub btnContaEstoque_Click(sender As Object, e As EventArgs) Handles btnContaEstoque.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaEstoque" & HID.Value)
    End Sub

    Protected Sub btnContaEstoqueEmNossoPoder_Click(sender As Object, e As EventArgs) Handles btnContaEstoqueEmNossoPoder.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaEstoqueEmNossoPoder" & HID.Value)
    End Sub

    Protected Sub btnContaEstoqueEmPoderDeTerceiros_Click(sender As Object, e As EventArgs) Handles btnContaEstoqueEmPoderDeTerceiros.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ContaEstoqueEmPoderDeTerceiros" & HID.Value)
    End Sub

    Protected Sub btnCaixaCompensacao_Click(sender As Object, e As EventArgs) Handles btnCaixaCompensacao.Click
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "ssCaixaCompensacao" & HID.Value)
    End Sub
#End Region

#Region "Botoes"
    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Dim Empresa() As String = DdlEmpresa.SelectedValue.Split("-")

        SessaoRecuperaEmpresa()
        objEmpresa.IUD = "I"

        objEmpresa.AtividadeEconomica = txtAtividadeEconomicaFederal.Text
        objEmpresa.AtividadeEstadual = txtAtividadeEconomicaEstadual.Text
        objEmpresa.DataDeFundacao = CDate(txtDataDaFundacao.Text)
        objEmpresa.RegistroNaJunta = txtRegistroNaJunta.Text
        objEmpresa.EstadoDaJunta = txtEstadoDaJunta.Text
        objEmpresa.NomeDoContador = txtNomeDoContador.Text
        objEmpresa.CPFContador = txtCpfDoContador.Text.Replace(".", "").Replace("/", "").Replace("-", "")
        objEmpresa.CNPJContador = txtCnpjDoContador.Text.Replace(".", "").Replace("/", "").Replace("-", "")
        objEmpresa.CRCContador = txtCrcDoContador.Text
        objEmpresa.EstadoExpCRC = ddlEstadoExpCRC.SelectedValue
        objEmpresa.TelefoneContador = txtTelefoneContador.Text
        objEmpresa.EmailContador = txtEmailContador.Text
        objEmpresa.NomeDoTitular = txtNomeDoTitular.Text
        objEmpresa.CPFTitular = txtCpfDoTitular.Text.Replace(".", "").Replace("/", "").Replace("-", "")
        objEmpresa.CondicaoTitular = txtQualificacaoDoTitular.Text
        objEmpresa.Matriz = txtMatriz.Text
        objEmpresa.InscricaoMunicipal = txtInscricaoMunicipal.Text
        objEmpresa.RegistroEp = txtRegistroEP.Text
        objEmpresa.RegistroEi = txtRegistroEI.Text
        objEmpresa.RegistroEc = txtRegistroEC.Text
        objEmpresa.Marca = txtMarca.Text
        objEmpresa.RegistroNire = txtNire.Text
        objEmpresa.DataRegistroNire = CDate(txtDataDaNire.Text)
        objEmpresa.QualificacaoContador = txtQualificacaoDoContador.Text
        objEmpresa.CodigoQualificacaoContador = DdlCodigoDeQualificacaoDoContador.SelectedValue
        objEmpresa.QualificacaoTitular = txtQualificacaoDoTitular.Text
        objEmpresa.CodigoQualificacaoTitular = DdlCodigoDaQualificacaoDoTitular.SelectedValue
        objEmpresa.CodigoDeRelacionamento = DdlCodigoDeRelacionamento.SelectedValue
        objEmpresa.EscrituracaoContabil = DdlEscrituracaoContabil.SelectedValue
        objEmpresa.NotaFiscalEletronica = chknfe.Checked
        objEmpresa.EmitirCTe = chkEmitirCTe.Checked
        objEmpresa.ObrigaNfProdutor = chknfProdutor.Checked
        objEmpresa.ObrigaNavio = chkObrigaNavio.Checked
        objEmpresa.ObrigaChaveNf = chkObrigaChaveNf.Checked
        objEmpresa.ObrigaChaveNfg = chkObrigaChaveNfg.Checked
        objEmpresa.CodigoContadeCustoInicial = Funcoes.AlinharEsquerda(txtContaInicial.Text, 9, "0")
        objEmpresa.CodigoContadeCustoFinal = Funcoes.AlinharEsquerda(txtContaFinal.Text, 9, "9")
        objEmpresa.CodigoContaPatrimonio = txtContaPatrimonio.Text
        objEmpresa.Crt = ddlCrt.SelectedValue
        objEmpresa.CertidaoNegativa = chkCertidaoNegativa.Checked
        objEmpresa.FluxoDeCaixa = chkFluxoDeCaixa.Checked
        objEmpresa.SugereDeposito = chkSugereDeposito.Checked
        objEmpresa.NossaEmissao = chkNossaEmissao.Checked
        objEmpresa.RegistroDeExportacao = chkRegistroDeExportacao.Checked
        objEmpresa.CodigoContaVariacaoMonetariaAtiva = txtContaVariacaoAtiva.Text
        objEmpresa.CodigoContaVariacaoMonetariaPassiva = txtContaVariacaoPassiva.Text
        objEmpresa.CodigoContaVariacaoMonetariaCliente = txtContaVariacaoCliente.Text
        objEmpresa.CodigoContaVariacaoMonetariaFornecedor = txtContaVariacaoFornecedor.Text
        objEmpresa.CodigoContaGrupoBanco = txtContaContabilGrupoBanco.Text
        objEmpresa.FretePedido = chkFretePedido.Checked
        objEmpresa.BaixaFinanceiroPorLote = chkBaixaFinanceiroPorLote.Checked
        objEmpresa.PedidoBloqueado = chkPedidoBloqueado.Checked
        objEmpresa.LiberaCarregamento = chkLiberaCarregamento.Checked
        objEmpresa.CodigoContaJuroPago = txtContaJuroPago.Text
        objEmpresa.CodigoContaJuroRecebido = txtContaJuroRecebido.Text
        objEmpresa.CodigoContaGrupoDuplicatasDescontada = txtContaDuplicatasDescontada.Text
        objEmpresa.CodigoContaGrupoEstoque = txtContaEstoque.Text
        objEmpresa.CodigoContaGrupoComissoes = txtContaGrupoComissoes.Text
        objEmpresa.CodigoContaAdiantamentoDeFrete = txtContaAdiantamentoDeFrete.Text
        objEmpresa.CodigoContaPedagioDeFrete = txtContaPedagioDeFrete.Text
        objEmpresa.CodigoProdutoDeFrete = txtProdutoDeFrete.Text
        objEmpresa.CodigoProdutoDeEstadia = txtProdutoDeEstadia.Text
        objEmpresa.CodigoProdutoDeMDFe = txtProdutoDeMDFe.Text
        objEmpresa.Servidor = ddlServidor.SelectedValue
        objEmpresa.ViasNFE = txtNumeroDeViasNFE.Text
        objEmpresa.PrazoCancelamentoNFE = txtPrazoCancelamentoNFE.Text
        objEmpresa.ControlaEmissaoCheque = RadControlaEmissaChequeSim.Checked
        objEmpresa.ControlaDataMovimentoNFG = RadControlaDataMOvimentoNFGSim.Checked
        objEmpresa.ObservacaoSefazNFE = txtObservacaoSefazNFE.Text
        objEmpresa.PathDownloadNFe = lblDownloadNFe.Text
        objEmpresa.DiasNFRetroativa = txtDiasNFRetroativa.Text

        objEmpresa.ArquivoNFE = radArquivoNFESim.Checked
        objEmpresa.ConferenciaNFE = radConferenciaNFESim.Checked

        objEmpresa.UsarDescricaoProduto = chkUsarDescricaoProduto.Checked
        objEmpresa.UsarRegistroMinAgr = chkUsarRegistroMinAgr.Checked

        objEmpresa.CodigoContaGrupoTedDoc = txtContaTEDDOC.Text
        objEmpresa.CodigoContaDescontoConcedido = txtContaDescontoConcedido.Text
        objEmpresa.CodigoContaDescontoObtido = txtContaDescontoObtido.Text
        objEmpresa.CodigoContaFornecedorFrete = txtFornecedorFrete.Text
        objEmpresa.CodigoContaCaixaCompensacao = txtCaixaCompensacao.Text
        objEmpresa.CodigoContaEstoqueEmNossoPoder = txtContaEstoqueEmNossoPoder.Text
        objEmpresa.CodigoContaEstoqueEmPoderDeTerceiros = txtContaEstoqueEmPoderDeTerceiros.Text
        objEmpresa.DataFinanceiro = CDate(txtDataFinanceiro.Text)
        objEmpresa.ObrigaEncargo = chkObgEncargo.Checked
        objEmpresa.CodigoCnae = ddlCnae.SelectedValue
        If Not String.IsNullOrWhiteSpace(ddlNaturezaJuridica.SelectedItem.Value) Then objEmpresa.CodigoNaturezaJuridica = ddlNaturezaJuridica.SelectedValue

        If objEmpresa.Matriz = "S" Then
            objEmpresa.ParametrosECF.IUD = "I"

            'objEmpresa.ParametrosECF.IND_SIT_INI_PER = row("IND_SIT_INI_PER")
            'objEmpresa.ParametrosECF.SIT_ESPECIAL = row("SIT_ESPECIAL")
            'objEmpresa.ParametrosECF.TIP_ECF = row("TIP_ECF")

            objEmpresa.ParametrosECF.OPT_REFIS = chkOPT_REFIS.Checked
            objEmpresa.ParametrosECF.OPT_PAES = chkOPT_PAES.Checked
            objEmpresa.ParametrosECF.FORMA_TRIB = ddlFORMA_TRIB.SelectedValue
            objEmpresa.ParametrosECF.FORMA_APUR = ddlFORMA_APUR.SelectedValue
            objEmpresa.ParametrosECF.COD_QUALIF_PJ = ddlCOD_QUALIF_PJ.SelectedValue
            objEmpresa.ParametrosECF.FORMA_TRIB_PER = ddlFORMA_TRIB_PER.SelectedValue
            objEmpresa.ParametrosECF.MES_BAL_RED = ddlMES_BAL_RED.SelectedValue
            objEmpresa.ParametrosECF.TIP_ESC_PRE = ddlTIP_ESC_PRE.SelectedValue
            objEmpresa.ParametrosECF.TIP_ENT = ddlTIP_ENT.SelectedValue
            objEmpresa.ParametrosECF.FORMA_APUR_I = ddlFORMA_APUR_I.SelectedValue
            objEmpresa.ParametrosECF.APUR_CSLL = ddlAPUR_CSLL.SelectedValue
            objEmpresa.ParametrosECF.OPT_EXT_RTT = chkOPT_EXT_RTT.Checked
            objEmpresa.ParametrosECF.DIF_FCONT = chkDIF_FCONT.Checked
            objEmpresa.ParametrosECF.IND_ALIQ_CSLL = chkIND_ALIQ_CSLL.Checked
            objEmpresa.ParametrosECF.IND_ALIQ_CSLL_I = ddlIND_ALIQ_CSLL_I.SelectedIndex
            objEmpresa.ParametrosECF.IND_QTE_SCP = txtIND_QTE_SCP.Text
            objEmpresa.ParametrosECF.IND_ADM_FUN_CLU = chkIND_ADM_FUN_CLU.Checked
            objEmpresa.ParametrosECF.IND_PART_CONS = chkIND_PART_CONS.Checked
            objEmpresa.ParametrosECF.IND_OP_EXT = chkIND_OP_EXT.Checked
            objEmpresa.ParametrosECF.IND_OP_VINC = chkIND_OP_VINC.Checked
            objEmpresa.ParametrosECF.IND_PJ_ENQUAD = chkIND_PJ_ENQUAD.Checked
            objEmpresa.ParametrosECF.IND_PART_EXT = chkIND_PART_EXT.Checked
            objEmpresa.ParametrosECF.IND_ATIV_RURAL = chkIND_ATIV_RURAL.Checked
            objEmpresa.ParametrosECF.IND_LUC_EXP = chkIND_LUC_EXP.Checked
            objEmpresa.ParametrosECF.IND_RED_ISEN = chkIND_RED_ISEN.Checked
            objEmpresa.ParametrosECF.IND_FIN = chkIND_FIN.Checked
            objEmpresa.ParametrosECF.IND_DOA_ELEIT = chkIND_DOA_ELEIT.Checked
            objEmpresa.ParametrosECF.IND_PART_COLIG = chkIND_PART_COLIG.Checked
            objEmpresa.ParametrosECF.IND_VEND_EXP = chkIND_VEND_EXP.Checked
            objEmpresa.ParametrosECF.IND_REC_EXT = chkIND_REC_EXT.Checked
            objEmpresa.ParametrosECF.IND_ATIV_EXT = chkIND_ATIV_EXT.Checked
            objEmpresa.ParametrosECF.IND_COM_EXP = chkIND_COM_EXP.Checked
            objEmpresa.ParametrosECF.IND_PGTO_EXT = chkIND_PGTO_EXT.Checked
            objEmpresa.ParametrosECF.IND_E_COM_TI = chkIND_E_COM_TI.Checked
            objEmpresa.ParametrosECF.IND_ROY_REC = chkIND_ROY_REC.Checked
            objEmpresa.ParametrosECF.IND_ROY_PAG = chkIND_ROY_PAG.Checked
            objEmpresa.ParametrosECF.IND_REND_SERV = chkIND_REND_SERV.Checked
            objEmpresa.ParametrosECF.IND_PGTO_REM = chkIND_PGTO_REM.Checked
            objEmpresa.ParametrosECF.IND_INOV_TEC = chkIND_INOV_TEC.Checked
            objEmpresa.ParametrosECF.IND_CAP_INF = chkIND_CAP_INF.Checked
            objEmpresa.ParametrosECF.IND_PJ_HAB = chkIND_PJ_HAB.Checked
            objEmpresa.ParametrosECF.IND_POLO_AM = chkIND_POLO_AM.Checked
            objEmpresa.ParametrosECF.IND_ZON_EXP = chkIND_ZON_EXP.Checked
            objEmpresa.ParametrosECF.IND_AREA_COM = chkIND_AREA_COM.Checked
        End If

        If Not objEmpresa.Salvar Then
            MsgBox(Me.Page, Session("ssMessage"))
        Else
            Limpar()
            MsgBox(Me.Page, "Registro gravado com Sucesso.", eTitulo.Sucess)
        End If
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        'Dim Empresa() As String = DdlEmpresa.SelectedValue.Split("-")
        'If ValidarExcluir() Then
        '    Sql = " DELETE FROM ClientesXEmpresas"
        '    Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' "

        '    Sqls.Add(Sql)

        '    If Not Banco.GravaBanco(Sqls) Then
        '        MsgBox(Me.Page, Session("ssMessage"))
        '    Else
        '        Limpar()
        '        MsgBox(Me.Page, "Registro excluído com Sucesso.", eTitulo.Sucess)
        '    End If
        'End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível limpar o registro.")
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Empresas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub cmdAjuda_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim NomeArquivo As String = "Manual/Empresas.mht"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
    End Sub
#End Region

    Protected Sub DdlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlEmpresa.SelectedIndexChanged
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedItem.Value) Then
            Limpar()
            Exit Sub
        End If

        Dim Empresa() As String = DdlEmpresa.SelectedValue.Split("-")

        objEmpresa = New ClienteXEmpresa(Empresa(0), Empresa(1))

        txtRegistroNaJunta.Text = objEmpresa.RegistroNaJunta
        txtEstadoDaJunta.Text = objEmpresa.EstadoDaJunta
        If IsDate(objEmpresa.DataDeFundacao) Then
            txtDataDaFundacao.Text = Format(objEmpresa.DataDeFundacao, "dd/MM/yyyy")
        End If

        txtNire.Text = objEmpresa.RegistroNire
        If IsDate(objEmpresa.DataRegistroNire) Then
            txtDataDaNire.Text = Format(objEmpresa.DataRegistroNire, "dd/MM/yyyy")
        End If

        ddlCnae.SelectedValue = objEmpresa.CodigoCnae
        If objEmpresa.CodigoNaturezaJuridica = 0 Then
            ddlNaturezaJuridica.SelectedIndex = 0
        Else
            ddlNaturezaJuridica.SelectedValue = objEmpresa.CodigoNaturezaJuridica
        End If

        txtInscricaoMunicipal.Text = objEmpresa.InscricaoMunicipal
        txtAtividadeEconomicaEstadual.Text = objEmpresa.AtividadeEstadual
        txtAtividadeEconomicaFederal.Text = objEmpresa.AtividadeEconomica
        txtMatriz.Text = objEmpresa.Matriz
        txtRegistroEC.Text = objEmpresa.RegistroEc
        txtRegistroEP.Text = objEmpresa.RegistroEp
        txtRegistroEI.Text = objEmpresa.RegistroEi
        txtMarca.Text = objEmpresa.Marca

        txtNomeDoContador.Text = objEmpresa.NomeDoContador

        If Not String.IsNullOrWhiteSpace(objEmpresa.CPFContador) Then txtCpfDoContador.Text = Funcoes.FormatarCpfCnpj(objEmpresa.CPFContador)

        If Not String.IsNullOrWhiteSpace(objEmpresa.CNPJContador) Then txtCnpjDoContador.Text = Funcoes.FormatarCpfCnpj(objEmpresa.CNPJContador)

        txtCrcDoContador.Text = objEmpresa.CRCContador

        ddlEstadoExpCRC.SelectedValue = objEmpresa.EstadoExpCRC
        txtTelefoneContador.Text = objEmpresa.TelefoneContador
        txtEmailContador.Text = objEmpresa.EmailContador
        txtQualificacaoDoContador.Text = objEmpresa.QualificacaoContador
        DdlCodigoDeQualificacaoDoContador.SelectedValue = objEmpresa.CodigoQualificacaoContador

        txtNomeDoTitular.Text = objEmpresa.NomeDoTitular
        txtCpfDoTitular.Text = Funcoes.FormatarCpfCnpj(objEmpresa.CPFTitular)
        txtQualificacaoDoTitular.Text = objEmpresa.QualificacaoTitular
        DdlCodigoDaQualificacaoDoTitular.SelectedValue = objEmpresa.CodigoQualificacaoTitular

        chknfe.Checked = objEmpresa.NotaFiscalEletronica
        chkEmitirCTe.Checked = objEmpresa.EmitirCTe
        chknfProdutor.Checked = objEmpresa.ObrigaNfProdutor
        chkObrigaNavio.Checked = objEmpresa.ObrigaNavio
        chkObrigaChaveNf.Checked = objEmpresa.ObrigaChaveNf
        chkObrigaChaveNfg.Checked = objEmpresa.ObrigaChaveNfg
        chkCertidaoNegativa.Checked = objEmpresa.CertidaoNegativa
        chkObgEncargo.Checked = objEmpresa.ObrigaEncargo
        chkSugereDeposito.Checked = objEmpresa.SugereDeposito
        chkFluxoDeCaixa.Checked = objEmpresa.FluxoDeCaixa
        chkRegistroDeExportacao.Checked = objEmpresa.RegistroDeExportacao
        chkFretePedido.Checked = objEmpresa.FretePedido
        chkPedidoBloqueado.Checked = objEmpresa.PedidoBloqueado
        chkLiberaCarregamento.Checked = objEmpresa.LiberaCarregamento
        chkNossaEmissao.Checked = objEmpresa.NossaEmissao
        chkBaixaFinanceiroPorLote.Checked = objEmpresa.BaixaFinanceiroPorLote

        txtContaInicial.Text = CInt(IIf(objEmpresa.CodigoContadeCustoInicial.Length = 0, "0", StrReverse(objEmpresa.CodigoContadeCustoInicial.ToString))).ToString
        If objEmpresa.ContadeCustoInicial IsNot Nothing Then txtContaInicial.ToolTip = objEmpresa.ContadeCustoInicial.Titulo

        txtContaFinal.Text = IIf(objEmpresa.CodigoContadeCustoFinal.Length = 0, "9", objEmpresa.CodigoContadeCustoFinal.Substring(0, CInt(StrReverse(objEmpresa.CodigoContadeCustoFinal.Replace("9", "0"))).ToString.Length))
        If objEmpresa.ContadeCustoFinal IsNot Nothing Then txtContaFinal.ToolTip = objEmpresa.ContadeCustoFinal.Titulo

        txtContaPatrimonio.Text = objEmpresa.CodigoContaPatrimonio
        If objEmpresa.ContaPatrimonio IsNot Nothing Then txtContaPatrimonio.ToolTip = objEmpresa.ContaPatrimonio.Titulo

        txtContaVariacaoAtiva.Text = objEmpresa.CodigoContaVariacaoMonetariaAtiva
        If objEmpresa.ContaVariacaoMonetariaAtiva IsNot Nothing Then txtContaVariacaoAtiva.ToolTip = objEmpresa.ContaVariacaoMonetariaAtiva.Titulo

        txtContaVariacaoPassiva.Text = objEmpresa.CodigoContaVariacaoMonetariaPassiva
        If objEmpresa.ContaVariacaoMonetariaPassiva IsNot Nothing Then txtContaVariacaoPassiva.ToolTip = objEmpresa.ContaVariacaoMonetariaPassiva.Titulo

        txtContaVariacaoCliente.Text = objEmpresa.CodigoContaVariacaoMonetariaCliente
        If objEmpresa.ContaVariacaoMonetariaCliente IsNot Nothing Then txtContaVariacaoCliente.ToolTip = objEmpresa.ContaVariacaoMonetariaCliente.Titulo

        txtContaVariacaoFornecedor.Text = objEmpresa.CodigoContaVariacaoMonetariaFornecedor
        If objEmpresa.ContaVariacaoMonetariafornecedor IsNot Nothing Then txtContaVariacaoFornecedor.ToolTip = objEmpresa.ContaVariacaoMonetariafornecedor.Titulo

        txtContaContabilGrupoBanco.Text = objEmpresa.CodigoContaGrupoBanco
        If objEmpresa.ContaGrupoBanco IsNot Nothing Then txtContaContabilGrupoBanco.ToolTip = objEmpresa.ContaGrupoBanco.Titulo

        txtContaJuroPago.Text = objEmpresa.CodigoContaJuroPago
        If objEmpresa.ContaJuroPago IsNot Nothing Then txtContaJuroPago.ToolTip = objEmpresa.ContaJuroPago.Titulo

        txtContaJuroRecebido.Text = objEmpresa.CodigoContaJuroRecebido
        If objEmpresa.ContaJuroRecebido IsNot Nothing Then txtContaJuroRecebido.ToolTip = objEmpresa.ContaJuroRecebido.Titulo

        txtContaDuplicatasDescontada.Text = objEmpresa.CodigoContaGrupoDuplicatasDescontada
        If objEmpresa.ContaGrupoDuplicatasDescontada IsNot Nothing Then txtContaDuplicatasDescontada.ToolTip = objEmpresa.ContaGrupoDuplicatasDescontada.Titulo

        txtContaEstoque.Text = objEmpresa.CodigoContaGrupoEstoque
        If objEmpresa.ContaGrupoEstoque IsNot Nothing Then txtContaEstoque.ToolTip = objEmpresa.ContaGrupoEstoque.Titulo

        txtContaGrupoComissoes.Text = objEmpresa.CodigoContaGrupoComissoes
        If objEmpresa.ContaGrupoComissoes IsNot Nothing Then txtContaGrupoComissoes.ToolTip = objEmpresa.ContaGrupoComissoes.Titulo

        txtContaTEDDOC.Text = objEmpresa.CodigoContaGrupoTedDoc
        If objEmpresa.ContaGrupoTedDoc IsNot Nothing Then txtContaTEDDOC.ToolTip = objEmpresa.ContaGrupoTedDoc.Titulo

        txtContaDescontoConcedido.Text = objEmpresa.CodigoContaDescontoConcedido
        If objEmpresa.ContaDescontoConcedido IsNot Nothing Then txtContaDescontoConcedido.Text = objEmpresa.ContaDescontoConcedido.Titulo

        txtContaDescontoObtido.Text = objEmpresa.CodigoContaDescontoObtido
        If objEmpresa.ContaDescontoObtido IsNot Nothing Then txtContaDescontoObtido.Text = objEmpresa.ContaDescontoObtido.Titulo

        txtContaPedagioDeFrete.Text = objEmpresa.CodigoContaPedagioDeFrete
        If objEmpresa.ContaPedagioDeFrete IsNot Nothing Then txtContaPedagioDeFrete.Text = objEmpresa.ContaPedagioDeFrete.Titulo

        txtContaAdiantamentoDeFrete.Text = objEmpresa.CodigoContaAdiantamentoDeFrete
        If objEmpresa.ContaAdiantamentoDeFrete IsNot Nothing Then txtContaAdiantamentoDeFrete.Text = objEmpresa.ContaAdiantamentoDeFrete.Titulo

        txtProdutoDeFrete.Text = objEmpresa.CodigoProdutoDeFrete
        txtProdutoDeEstadia.Text = objEmpresa.CodigoProdutoDeEstadia
        txtProdutoDeMDFe.Text = objEmpresa.CodigoProdutoDeMDFe

        txtFornecedorFrete.Text = objEmpresa.CodigoContaFornecedorFrete
        If objEmpresa.ContaFornecedorFrete IsNot Nothing Then txtFornecedorFrete.ToolTip = objEmpresa.ContaFornecedorFrete.Titulo

        txtCaixaCompensacao.Text = objEmpresa.CodigoContaCaixaCompensacao
        If objEmpresa.ContaCaixaCompensacao IsNot Nothing Then txtCaixaCompensacao.ToolTip = objEmpresa.ContaCaixaCompensacao.Titulo

        txtContaEstoqueEmNossoPoder.Text = objEmpresa.CodigoContaEstoqueEmNossoPoder
        If objEmpresa.ContaEstoqueEmNossoPoder IsNot Nothing Then txtContaEstoqueEmNossoPoder.ToolTip = objEmpresa.ContaEstoqueEmNossoPoder.Titulo

        txtContaEstoqueEmPoderDeTerceiros.Text = objEmpresa.CodigoContaEstoqueEmPoderDeTerceiros
        If objEmpresa.ContaEstoqueEmPoderDeTerceiros IsNot Nothing Then txtContaEstoqueEmPoderDeTerceiros.ToolTip = objEmpresa.ContaEstoqueEmPoderDeTerceiros.Titulo

        txtPrazoCancelamentoNFE.Text = objEmpresa.PrazoCancelamentoNFE
        txtNumeroDeViasNFE.Text = objEmpresa.ViasNFE

        ddlCrt.SelectedValue = objEmpresa.Crt
        ddlServidor.SelectedValue = objEmpresa.Servidor

        If objEmpresa.ControlaEmissaoCheque Then
            RadControlaEmissaChequeSim.Checked = True
        Else
            RadControlaEmissaChequeNao.Checked = True
        End If

        If objEmpresa.ControlaDataMovimentoNFG Then
            RadControlaDataMOvimentoNFGSim.Checked = True
        Else
            RadControlaDataMOvimentoNFGNao.Checked = True
        End If

        radArquivoNFENao.Checked = Not objEmpresa.ArquivoNFE
        radArquivoNFESim.Checked = objEmpresa.ArquivoNFE
        radConferenciaNFENao.Checked = Not objEmpresa.ConferenciaNFE
        radConferenciaNFESim.Checked = objEmpresa.ConferenciaNFE

        chkUsarDescricaoProduto.Checked = objEmpresa.UsarDescricaoProduto
        chkUsarRegistroMinAgr.Checked = objEmpresa.UsarRegistroMinAgr

        txtObservacaoSefazNFE.Text = objEmpresa.ObservacaoSefazNFE

        lblDownloadNFe.Text = objEmpresa.PathDownloadNFe
        txtDiasNFRetroativa.Text = objEmpresa.DiasNFRetroativa

        txtDataFinanceiro.Text = objEmpresa.DataFinanceiro.ToString("dd/MM/yyyy")

        ddlUsuarios.Items.Clear()
        If objEmpresa.UsuarioAlteracao IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objEmpresa.UsuarioAlteracao) AndAlso objEmpresa.UsuarioInclusao <> objEmpresa.UsuarioAlteracao Then
            ddlUsuarios.Items.Add("Alt.- " & objEmpresa.UsuarioAlteracao & " " & objEmpresa.DataAlteracao.ToString("dd/MM/yyyy"))
            ddlUsuarios.Items.Add("Inc.- " & objEmpresa.UsuarioInclusao & " " & objEmpresa.DataInclusao.ToString("dd/MM/yyyy"))
        Else
            ddlUsuarios.Items.Add("Inc.- " & objEmpresa.UsuarioInclusao & " " & objEmpresa.DataInclusao.ToString("dd/MM/yyyy"))
        End If

        If objEmpresa.Matriz = "S" Then
            TabECF.Visible = True

            chkOPT_REFIS.Checked = objEmpresa.ParametrosECF.OPT_REFIS
            chkOPT_PAES.Checked = objEmpresa.ParametrosECF.OPT_PAES
            ddlFORMA_TRIB.SelectedValue = objEmpresa.ParametrosECF.FORMA_TRIB
            DivIsentaImune.Visible = (ddlFORMA_TRIB.SelectedValue = 8 Or ddlFORMA_TRIB.SelectedValue = 9)
            ddlFORMA_APUR.SelectedValue = objEmpresa.ParametrosECF.FORMA_APUR
            ddlCOD_QUALIF_PJ.SelectedValue = objEmpresa.ParametrosECF.COD_QUALIF_PJ
            ddlFORMA_TRIB_PER.SelectedValue = objEmpresa.ParametrosECF.FORMA_TRIB_PER
            ddlMES_BAL_RED.SelectedValue = objEmpresa.ParametrosECF.MES_BAL_RED
            ddlTIP_ESC_PRE.SelectedValue = objEmpresa.ParametrosECF.TIP_ESC_PRE

            ddlTIP_ENT.SelectedValue = objEmpresa.ParametrosECF.TIP_ENT
            ddlFORMA_APUR_I.SelectedValue = objEmpresa.ParametrosECF.FORMA_APUR_I
            ddlAPUR_CSLL.SelectedValue = objEmpresa.ParametrosECF.APUR_CSLL

            chkOPT_EXT_RTT.Checked = objEmpresa.ParametrosECF.OPT_EXT_RTT
            chkDIF_FCONT.Checked = objEmpresa.ParametrosECF.DIF_FCONT

            chkIND_ALIQ_CSLL.Checked = objEmpresa.ParametrosECF.IND_ALIQ_CSLL
            ddlIND_ALIQ_CSLL_I.SelectedIndex = objEmpresa.ParametrosECF.IND_ALIQ_CSLL_I
            txtIND_QTE_SCP.Text = objEmpresa.ParametrosECF.IND_QTE_SCP
            chkIND_ADM_FUN_CLU.Checked = objEmpresa.ParametrosECF.IND_ADM_FUN_CLU
            chkIND_PART_CONS.Checked = objEmpresa.ParametrosECF.IND_PART_CONS
            chkIND_OP_EXT.Checked = objEmpresa.ParametrosECF.IND_OP_EXT
            chkIND_OP_VINC.Checked = objEmpresa.ParametrosECF.IND_OP_VINC
            chkIND_PJ_ENQUAD.Checked = objEmpresa.ParametrosECF.IND_PJ_ENQUAD
            chkIND_PART_EXT.Checked = objEmpresa.ParametrosECF.IND_PART_EXT
            chkIND_ATIV_RURAL.Checked = objEmpresa.ParametrosECF.IND_ATIV_RURAL
            chkIND_LUC_EXP.Checked = objEmpresa.ParametrosECF.IND_LUC_EXP
            chkIND_RED_ISEN.Checked = objEmpresa.ParametrosECF.IND_RED_ISEN
            chkIND_FIN.Checked = objEmpresa.ParametrosECF.IND_FIN
            chkIND_DOA_ELEIT.Checked = objEmpresa.ParametrosECF.IND_DOA_ELEIT
            chkIND_PART_COLIG.Checked = objEmpresa.ParametrosECF.IND_PART_COLIG
            chkIND_VEND_EXP.Checked = objEmpresa.ParametrosECF.IND_VEND_EXP
            chkIND_REC_EXT.Checked = objEmpresa.ParametrosECF.IND_REC_EXT
            chkIND_ATIV_EXT.Checked = objEmpresa.ParametrosECF.IND_ATIV_EXT
            chkIND_COM_EXP.Checked = objEmpresa.ParametrosECF.IND_COM_EXP
            chkIND_PGTO_EXT.Checked = objEmpresa.ParametrosECF.IND_PGTO_EXT
            chkIND_E_COM_TI.Checked = objEmpresa.ParametrosECF.IND_E_COM_TI
            chkIND_ROY_REC.Checked = objEmpresa.ParametrosECF.IND_ROY_REC
            chkIND_ROY_PAG.Checked = objEmpresa.ParametrosECF.IND_ROY_PAG
            chkIND_REND_SERV.Checked = objEmpresa.ParametrosECF.IND_REND_SERV
            chkIND_PGTO_REM.Checked = objEmpresa.ParametrosECF.IND_PGTO_REM
            chkIND_INOV_TEC.Checked = objEmpresa.ParametrosECF.IND_INOV_TEC
            chkIND_CAP_INF.Checked = objEmpresa.ParametrosECF.IND_CAP_INF
            chkIND_POLO_AM.Checked = objEmpresa.ParametrosECF.IND_POLO_AM
            chkIND_ZON_EXP.Checked = objEmpresa.ParametrosECF.IND_ZON_EXP
            chkIND_AREA_COM.Checked = objEmpresa.ParametrosECF.IND_AREA_COM
            chkIND_PJ_HAB.Checked = objEmpresa.ParametrosECF.IND_PJ_HAB
        Else
            TabECF.Visible = False
        End If

        lnkExcluir.Parent.Visible = True
        SessaoSalvaEmpresa()
    End Sub

    Protected Sub ddlFORMA_TRIB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlFORMA_TRIB.SelectedIndexChanged
        DivIsentaImune.Visible = (ddlFORMA_TRIB.SelectedValue = 8 Or ddlFORMA_TRIB.SelectedValue = 9)
    End Sub
    Private Sub CarregarIND_ALIQ_CSLL()
        ddlIND_ALIQ_CSLL_I.Items.Clear()
        ddlIND_ALIQ_CSLL_I.Items.Insert(0, "")
        ddlIND_ALIQ_CSLL_I.Items.Insert(1, "9%")
        ddlIND_ALIQ_CSLL_I.Items.Insert(2, "17%")
        ddlIND_ALIQ_CSLL_I.Items.Insert(3, "20%")
        ddlIND_ALIQ_CSLL_I.SelectedIndex = 0
    End Sub

End Class