Imports System.IO
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports CrystalDecisions.CrystalReports.Engine

Partial Class PedidosXItens
    Inherits BasePage

#Region "Inicialização"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If IsConnect AndAlso Not IsPostBack Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Comercial.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("PedidosXItens", "ACESSAR") Then
                    ddl.Carregar(cmbUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
                    ddl.Carregar(cmbSituacao, CarregarDDL.Tabela.Situacao)
                    ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao)
                    ddl.Carregar(cmbSafra, CarregarDDL.Tabela.Safra)
                    ddl.Carregar(cmbMoeda, CarregarDDL.Tabela.Moeda, "", True)
                    ddl.Carregar(cmbFinalidade, CarregarDDL.Tabela.Finalidade)

                    ddl.Carregar(cmbCondicoes, CarregarDDL.Tabela.CondicaoDePagamento)
                    ddl.Carregar(lstCondicoesPgtoEntrega, CarregarDDL.Tabela.CondicaoDePagamento)

                    ddl.Carregar(ddlCondPagPed, CarregarDDL.Tabela.CondicaoDePagamento)
                    ddl.Carregar(ddlCondPagNotas, CarregarDDL.Tabela.CondicaoDePagamento)
                    ddl.Carregar(cmbEmbalagem, CarregarDDL.Tabela.Embalagem, "", True)

                    divComissoes.Visible = False

                    ddlTipoDeposito.DataTextField = "Descricao"
                    ddlTipoDeposito.DataValueField = "Codigo"
                    ddlTipoDeposito.Items.Add(New ListItem("DE - DEPÓSITO", "DE"))
                    ddlTipoDeposito.Items.Add(New ListItem("OD - ORIGEM/DESTINO", "OD"))
                    ddlTipoDeposito.Items.Add(New ListItem("LE - LOCAL DE EMBARQUE/COLETA", "LE"))
                    ddlTipoDeposito.Items.Add(New ListItem("TR - TRANSBORDO", "TR"))
                    ddlTipoDeposito.Items.Add(New ListItem("PM - PROPRIETÁRIO DA MERCADORIA", "PM"))
                    Funcoes.InserirLinhaEmBranco(ddlTipoDeposito)
                    LimparPedido("")
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Variáveis Locais"

    Private objPedido As [Lib].Negocio.Pedido
    Private objEstado As [Lib].Negocio.Estado
    Private objMunicipio As [Lib].Negocio.Municipio
#End Region

#Region "Sessões"
    Private Sub SessaoRecuperaPedido()
        objPedido = CType(Session("objPedido" & HID.Value), [Lib].Negocio.Pedido)
    End Sub

    Private Sub SessaoSalvaPedido()
        Session("objPedido" & HID.Value) = objPedido
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objClientePXI" & HID.Value) Is Nothing Then
            SessaoRecuperaPedido()
            Dim cli As Cliente = Session("objClientePXI" & HID.Value)
            objPedido.CodigoCliente = cli.Codigo
            objPedido.EnderecoCliente = cli.CodigoEndereco
            objPedido.CodigoPraca = objPedido.Cliente.Codigo
            objPedido.EnderecoPraca = objPedido.Cliente.CodigoEndereco
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objPedido.Cliente)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            txtPraca.Text = txtCliente.Text
            txtCodigoPraca.Value = txtCodigoCliente.Value
            objPedido.EstadoEntrega = objPedido.Cliente.CodigoEstado
            objPedido.CidadeEntrega = objPedido.Cliente.Cidade
            'txtUFEntrega.Text = objPedido.EstadoEntrega
            'txtCidadeEntrega.Text = objPedido.CidadeEntrega
            'btnUFEntrega.Enabled = True
            ddlTipoDeposito.SelectedValue = "DE"
            chkDepPrincipal.Checked = True
            ddl.Carregar(cmbDepositos, CarregarDDL.Tabela.Depositos, "", True)
            cmbDepositos.SelectedValue = objPedido.CodigoEmpresa & "-" & objPedido.EnderecoEmpresa
            Session.Remove("objClientePXI" & HID.Value)
            If objPedido.Codigo > 0 Then
                objPedido.ContaBancariaSelecionada = Nothing
                txtDadosBancarios.Text = ""
                For Each objItem As [Lib].Negocio.PedidoXItem In objPedido.Itens
                    objItem.Encargos.Clear()
                    objItem.Encargos.CriaListar()
                Next
                If cmbCondicoes.SelectedIndex > 0 Then CalcularParcelamento(False) Else LimparCondicoes(True)
            End If
            SessaoSalvaPedido()
        ElseIf Not Session("objPracaPXI" & HID.Value) Is Nothing Then
            SessaoRecuperaPedido()
            objPedido.CodigoPraca = CType(Session("objPracaPXI" & HID.Value), [Lib].Negocio.Cliente).Codigo
            objPedido.EnderecoPraca = CType(Session("objPracaPXI" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
            Dim itemPraca As ListItem = Funcoes.FormatarListItemCliente(objPedido.Praca)
            txtPraca.Text = itemPraca.Text
            txtCodigoPraca.Value = itemPraca.Value
            Session.Remove("objPracaPXI" & HID.Value)
            SessaoSalvaPedido()
        ElseIf Not Session("objClientePXDEP" & HID.Value) Is Nothing Then
            cmbDepositos.Items.Clear()
            cmbDepositos.DataTextField = "Descricao"
            cmbDepositos.DataValueField = "Codigo"
            Dim Dep As Cliente = CType(Session("objClientePXDEP" & HID.Value), [Lib].Negocio.Cliente)
            cmbDepositos.Items.Add(Funcoes.FormatarListItemCliente(Dep))
            Funcoes.InserirLinhaEmBranco(cmbDepositos)
            cmbDepositos.SelectedValue = Dep.Codigo & "-" & Dep.CodigoEndereco
            Session.Remove("objClientePXDEP" & HID.Value)
        ElseIf Not Session("objTransportadorPXI" & HID.Value) Is Nothing Then
            Dim objTransportadorPXI As Cliente = CType(Session("objTransportadorPXI" & HID.Value), [Lib].Negocio.Cliente)
            If (obj IsNot Nothing) Then
                hdfTransportadoras.Value = objTransportadorPXI.Codigo & "-" & objTransportadorPXI.CodigoEndereco
                Funcoes.FormatarClienteTXT(txtTransportadoras, objTransportadorPXI)
            End If
            Session.Remove("objTransportadorPXI" & HID.Value)
        ElseIf Not Session("objRepresentantePXI" & HID.Value) Is Nothing Then
            Dim objRepresentantePXI As Cliente = CType(Session("objRepresentantePXI" & HID.Value), [Lib].Negocio.Cliente)
            If (objRepresentantePXI IsNot Nothing) Then
                hdfRepresentante.Value = objRepresentantePXI.Codigo & "-" & CStr(objRepresentantePXI.CodigoEndereco)
                Funcoes.FormatarClienteTXT(txtRepresentante, objRepresentantePXI)
            End If
            Session.Remove("objRepresentantePXI" & HID.Value)
        ElseIf Session("objEstadoPDXI" & HID.Value) IsNot Nothing Then
            SessaoRecuperaPedido()
            objEstado = CType(Session("objEstadoPDXI" & HID.Value), [Lib].Negocio.Estado)
            objPedido.EstadoEntrega = objEstado.Codigo
            'txtUFEntrega.Text = objPedido.EstadoEntrega
            SessaoSalvaPedido()
            Session.Remove("objEstadoPDXI" & HID.Value)
            HttpContext.Current.Session("ssUF") = objPedido.EstadoEntrega
            ucConsultaCodMunicipios.Limpar()
            Popup.ConsultaDeMunicipios(Me.Page, "objMunicipioPDXI" & HID.Value)
        ElseIf Session("objMunicipioPDXI" & HID.Value) IsNot Nothing Then
            SessaoRecuperaPedido()
            objMunicipio = CType(Session("objMunicipioPDXI" & HID.Value), [Lib].Negocio.Municipio)
            If (objMunicipio IsNot Nothing) Then
                objPedido.CidadeEntrega = objMunicipio.CodigoMunicipio
                'txtCidadeEntrega.Text = objPedido.CidadeEntrega
                SessaoSalvaPedido()
                Session.Remove("objMunicipioPDXI" & HID.Value)
            End If
        ElseIf Session("objPedidosXItens" & HID.Value) IsNot Nothing Then
            objPedido = CType(Session("objPedidosXItens" & HID.Value), [Lib].Negocio.Pedido)
            objPedido.IUD = "U"
            CarregarFormularioComAClasse()
            Session.Remove("objPedidosXItens" & HID.Value)
            SessaoSalvaPedido()
        ElseIf Session("objPedidoTroca" & HID.Value) IsNot Nothing Then
            SessaoRecuperaPedido()
            Dim PedidoOrigem As [Lib].Negocio.Pedido = CType(Session("objPedidoTroca" & HID.Value), [Lib].Negocio.Pedido)
            objPedido.CodigoEmpresaTroca = PedidoOrigem.CodigoEmpresa
            objPedido.EnderecoEmpresaTroca = PedidoOrigem.EnderecoEmpresa
            objPedido.CodigoPedidoTroca = PedidoOrigem.Codigo
            Dim itemEmpresaTroca As ListItem = Funcoes.FormatarListItemCliente(PedidoOrigem.Empresa)
            lblEmpresaTroca.Text = itemEmpresaTroca.Text
            Dim itemClienteTroca As ListItem = Funcoes.FormatarListItemCliente(PedidoOrigem.Cliente)
            lblClienteTroca.Text = itemClienteTroca.Text
            lblPedidoTroca.Text = PedidoOrigem.Codigo
            lblCNTroca.Text = PedidoOrigem.PedidoEfetivo
            txtVendaTroca.Text = objPedido.PedidoTroca.Itens.LiquidoOficial.ToString("N2")
            txtCompraTroca.Text = objPedido.Itens.LiquidoOficial.ToString("N2")
            txtSaldoTroca.Text = (objPedido.PedidoTroca.Itens.LiquidoOficial - objPedido.Itens.LiquidoOficial).ToString("N2")
            SessaoSalvaPedido()
        ElseIf Not Session("objContaBancariaPXI" & HID.Value) Is Nothing Then
            SessaoRecuperaPedido()
            objPedido.ContaBancariaSelecionada = CType(Session("objContaBancariaPXI" & HID.Value), [Lib].Negocio.ClienteXContaBancaria)
            With objPedido.ContaBancariaSelecionada
                txtDadosBancarios.Text = "BANCO: " & .CodigoBanco.ToString() & " - " & .Banco.Descricao & " - " &
                                         "AGÊNCIA: " & .CodigoAgencia.ToString() & "-" & .DigitoAgencia & " - " &
                                         "CONTA: " & .ContaCorrente & "-" & .DigitoConta
            End With
            Session.Remove("objContaBancariaPXI" & HID.Value)
            SessaoSalvaPedido()
        ElseIf Session("objNavioXInvoice" & HID.Value) IsNot Nothing Then
            SessaoRecuperaPedido()
            Dim objNavioXInvoice = CType(Session("objNavioXInvoice" & HID.Value), [Lib].Negocio.NavioXInvoice)
            objPedido.InvoiceNavio = objNavioXInvoice.Codigo
            txtNaviosXInvoice.Text = objNavioXInvoice.Navio_Id & " - " & objNavioXInvoice.Descricao
            Session.Remove("objNavioXInvoice" & HID.Value)
            SessaoSalvaPedido()
        End If
    End Sub
#End Region

#Region "Methods"

    Public Sub HabilitarAlteracao()
        lnkAtualizar.Parent.Visible = True
    End Sub

    Public Sub DesabilitaAlteracao()
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Sub CarregarFormularioComAClasse()
        Try
            '*** Botoes ****
            lnkNovoPedido.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            lnkConsultar.Parent.Visible = False

            'Pedido pode ser bloqueado novamente mesmo depois de já ter Faturamento - Furlan - 14/05/2025
            If Not objPedido.PedidoBloqueado AndAlso objPedido.CodigoSituacao = eSituacao.Normal AndAlso Funcoes.VerificaPermissao("BloquearPedido", "ACESSAR") Then
                lnkBloquear.Parent.Visible = True
            End If

            MudarSituacao("FISCAL", False)
            MudarSituacao("FINANCEIRO", False)

            ddlUsuarios.Items.Clear()
            If objPedido.UsuarioCancelamento IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objPedido.UsuarioCancelamento) Then
                ddlUsuarios.Items.Add("Can.- " & objPedido.UsuarioCancelamento)
            End If
            If objPedido.UsuarioAlteracao IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objPedido.UsuarioAlteracao) Then
                ddlUsuarios.Items.Add("Alt.- " & objPedido.UsuarioAlteracao)
            End If
            ddlUsuarios.Items.Add("Inc.- " & objPedido.UsuarioInclusao)


            '** Cabeçalho **
            '*************************
            '***  Primeira Coluna ****
            '*************************
            txtRegistro.Text = objPedido.Codigo.ToString()
            txtRegistro.Enabled = False
            imgExtratoPedido.Visible = True
            txtPedido.Text = objPedido.PedidoEfetivo
            txtContrato.Text = objPedido.Contrato
            txtXPedNFe.Text = objPedido.XPedNFe
            txtItemXPedNFe.Text = objPedido.ItemXPedNFe
            txtDataPedido.Text = objPedido.DataPedido.ToString("dd/MM/yyyy")
            txtDataEntregaInicial.Text = objPedido.DataEntregaInicial.ToString("dd/MM/yyyy")
            txtDataEntregaFinal.Text = objPedido.DataEntregaFinal.ToString("dd/MM/yyyy")
            txtDataVencimentoPedido.Text = objPedido.DataVencimentoPedido.ToString("dd/MM/yyyy")

            chkPeso.Checked = CInt(objPedido.OrigemDestino)
            ddlComercializacao.SelectedValue = objPedido.Comercializacao

            '*************************
            '***   Segunda Coluna  ***
            '*************************
            If objPedido.CodigoSituacao > 0 Then
                cmbSituacao.SelectedValue = objPedido.CodigoSituacao
            End If

            If objPedido.CodigoFinalidade > 0 Then
                cmbFinalidade.SelectedValue = objPedido.CodigoFinalidade
            End If

            If Not String.IsNullOrWhiteSpace(objPedido.CodigoSafra) Then
                cmbSafra.SelectedValue = objPedido.CodigoSafra
            End If

            If objPedido.CodigoMoeda > 0 Then
                cmbMoeda.SelectedValue = objPedido.CodigoMoeda
            End If
            cmbMoeda.Enabled = False

            If objPedido.Embalagem > 0 Then
                cmbEmbalagem.SelectedValue = objPedido.Embalagem
            End If

            txtLocalDeEmbarque.Text = objPedido.LocalDeEmbarque

            If objPedido.TipoCondicaoEntrega = "AC" Then
                rbPagAnterior.Checked = True
            ElseIf objPedido.TipoCondicaoEntrega = "PC" Then
                rbPagPosterior.Checked = True
            End If

            If objPedido.TipoPagamentoPtax = "XA" Then
                rbPtaxAnterior.Checked = True
            ElseIf objPedido.TipoPagamentoPtax = "XP" Then
                rbPtaxPosterior.Checked = True
            End If

            ddl.Carregar(cmbIndexador, CarregarDDL.Tabela.Indexador, "indexador_id = 99 or isnull(Moeda,0) = " & objPedido.CodigoMoeda, False)

            If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                cmbIndexador.Visible = False
                txtIndiceFixado.Visible = False
                imgAjudaFixarDolar1.Visible = False
            Else
                cmbIndexador.Visible = True
                txtIndiceFixado.Visible = True
                imgAjudaFixarDolar1.Visible = True
            End If

            cmbIndexador.SelectedValue = IIf(objPedido.Moeda.Classificacao = eTiposMoeda.Oficial, 99, objPedido.CodigoIndexador)
            cmbIndexador.Enabled = False

            If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                chkTemVariacao.Checked = False
                chkTemVariacao.Visible = False
            Else
                chkTemVariacao.Visible = True
                chkTemVariacao.Checked = objPedido.TemVariacao
                chkTemVariacao.Enabled = Not objPedido.TemFinanceiro(1)
            End If

            If objPedido.SubOperacao.PrecoFixo AndAlso objPedido.IndiceFixado = 0 AndAlso Not objPedido.CodigoIndexador = 99 Then objPedido.IndiceFixado = Funcoes.PegarValorConversao(objPedido.CodigoIndexador, objPedido.DataPedido)

            txtIndiceFixado.Text = objPedido.IndiceFixado.ToString("N8")

            'VAMOS ACRESCENTAR UM PROCESSO PARA DAR PARA PERMISSÃO A DETERMINADO USUÁRIO SE PODE OU NÃO ALTERAR
            chkIndexadorFixo.Checked = objPedido.IndexadorFixo
            If Funcoes.VerificaPermissao("IndexadorFixo", "ALTERAR") Then
                chkIndexadorFixo.Enabled = True
            Else
                chkIndexadorFixo.Enabled = False
            End If

            If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                txtIndiceFixado.Enabled = False
            Else
                txtIndiceFixado.Enabled = LiberarIndiceFixado()
            End If

            txtTaxa.Text = objPedido.Taxa.ToString("N2")

            'Criado processo para Liberar ajuste do índice fixado - furlan - 09/10/2024
            If Funcoes.VerificaPermissao("LIBERAINDICEFIXADO", "GRAVAR") Then
                cmbIndexador.Enabled = True
                txtIndiceFixado.Enabled = True
            End If

            If objPedido.InvoiceNavio > 0 Then
                Dim objNavioXInvoice = New NavioXInvoice(objPedido.InvoiceNavio)
                txtNaviosXInvoice.Text = objNavioXInvoice.Navio_Id & " - " & objNavioXInvoice.Descricao
            End If

            '**************************
            '***   Terceira Coluna  ***
            '**************************
            cmbUnidadeNegocio.SelectedValue = objPedido.CodigoUnidadeNegocio
            cmbUnidadeNegocio.Enabled = False
            ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.Empresas, objPedido.CodigoUnidadeNegocio, True)
            cmbEmpresa.SelectedValue = (objPedido.Empresa.Codigo & "-" & objPedido.Empresa.CodigoEndereco)
            cmbEmpresa.Enabled = False

            If objPedido.Empresa.Empresa.FretePedido Then
                'pnlRoteiro.Visible = True
                TabTransportes.Visible = True
                'divRoteiroLog.Style.Add("display", "none")
            Else
                pnlRoteiro.Visible = False
                TabTransportes.Visible = False
            End If

            cmbOperacao.SelectedValue = objPedido.CodigoOperacao
            cmbOperacao.Enabled = False
            ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.SubOperacaoSemOperacao, " so.Operacao_Id = " & objPedido.CodigoOperacao)
            cmbSubOperacao.SelectedValue = objPedido.CodigoSubOperacao
            cmbSubOperacao.Enabled = False

            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objPedido.Cliente)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            cmdConsultarCliente.Enabled = False

            Dim itemPraca As ListItem = Funcoes.FormatarListItemCliente(objPedido.Praca)
            txtPraca.Text = itemPraca.Text
            txtCodigoPraca.Value = itemPraca.Value

            '** Dados Bancários **
            If Not objPedido.ContaBancariaSelecionada Is Nothing Then
                With objPedido.ContaBancariaSelecionada
                    txtDadosBancarios.Text = "BANCO: " & .CodigoBanco.ToString() & " - " & .Banco.Descricao & vbCrLf &
                                             "AGÊNCIA: " & .CodigoAgencia.ToString() & "-" & .DigitoAgencia & vbCrLf &
                                             "CONTA: " & .ContaCorrente & "-" & .DigitoConta
                End With
            End If

            lblCifFob.Text = "Frete pela Empresa:"
            If objPedido.SubOperacao.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida Then
                rbCIF.Text = "Sim (CIF)"
                rbFOB.Text = "Não (FOB)"
            Else
                rbCIF.Text = "Não (CIF)"
                rbFOB.Text = "Sim (FOB)"
            End If

            '***********************************
            '***  Rodape do Abeçalho Coluna  ***
            '***********************************
            rbCIF.Checked = (objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.CIF)
            rbFOB.Checked = (objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.FOB)
            rbTER.Checked = (objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.TER)
            rbNenhum.Checked = (objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.NEN)

            chkTroca.Checked = objPedido.Troca
            chkAntecipada.Checked = objPedido.Antecipada
            chkRecompra.Checked = objPedido.Recompra

            chkAgruparFinanceiro.Checked = objPedido.AgruparFinanceiro

            'txtUFEntrega.Text = objPedido.EstadoEntrega
            'txtCidadeEntrega.Text = objPedido.CidadeEntrega
            'btnUFEntrega.Enabled = True

            '*****************************
            '****** Tab Produtos *********
            '*****************************
            If objPedido.Itens.Count > 0 AndAlso objPedido.Itens(0).Produto.Agrupar = "N" Then
                lnkAdicionarItem.Parent.Parent.Visible = False
            Else
                lnkAdicionarItem.Parent.Parent.Visible = True
            End If

            If objPedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS Then
                gridItensPedidoDeposito.DataSource = objPedido.Itens.ToArray
                If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    gridItensPedidoDeposito.Columns(10).Visible = False
                Else
                    gridItensPedidoDeposito.Columns(10).Visible = True
                End If
                gridItensPedidoDeposito.DataBind()
            ElseIf objPedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then
                gridItensPedidoAFixar.DataSource = objPedido.Itens.ToArray
                If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    gridItensPedidoAFixar.Columns(10).Visible = False
                    gridItensPedidoAFixar.Columns(13).Visible = False
                Else
                    gridItensPedidoAFixar.Columns(10).Visible = True
                    gridItensPedidoAFixar.Columns(13).Visible = True
                End If
                gridItensPedidoAFixar.DataBind()
            Else
                gridItensPedido.DataSource = objPedido.Itens.ToArray
                gridItensPedido.DataBind()
            End If

            AtualizarGridEncargos()

            '*****************************
            '******* Tab Depósitos *******
            '*****************************
            GridDepositos.DataSource = objPedido.Depositos.ToArray
            GridDepositos.DataBind()

            '******* Roteiros *******
            grdRoteiros.DataSource = objPedido.Roteiros
            grdRoteiros.DataBind()
            lnkLimparRoteiro.Parent.Visible = Not objPedido.IUD = "U"

            txtTotalQtdEmbarque.Text = objPedido.Roteiros.TotalQtdEmbarque.ToString("N2")
            txtTotalQtdDestino.Text = objPedido.Roteiros.TotalQtdDestino.ToString("N2")
            'txtTotalQtdPedido.Text = objPedido.Itens.QuantidadeTotal.ToString("N2")
            txtTotalQtdPedido.Text = QuantidadeTotalPedido.ToString("N2")

            '*****************************
            '***** Tab Transportes *******
            '*****************************
            gridTransportes.DataSource = objPedido.Transportadores.ToArray
            gridTransportes.DataBind()

            '*****************************
            '****** Tab Comissoes ********
            '*****************************
            GridRepresentantes.DataSource = objPedido.Representantes.ToArray
            GridRepresentantes.DataBind()

            '*****************************
            '**** Tab Vencimentos ********
            '*****************************
            If objPedido.SubOperacao.Financeiro Then
                TabVencimentosOld.Visible = objPedido.MomentoFinanceiro <> 9 AndAlso (Not FinanceiroVirtual AndAlso Not objPedido.Troca)
                TabParcelamento.Visible = (objPedido.MomentoFinanceiro = 9 Or objPedido.IUD = "I") And FinanceiroVirtual
            Else
                TabVencimentosOld.Visible = False
                TabParcelamento.Visible = False
            End If

            ddlMomentoFinanceiro.SelectedValue = objPedido.MomentoFinanceiro

            If objPedido.IUD = "U" Then
                ddlMomentoFinanceiro.Enabled = False
            End If

            If cmbCondicoes.Enabled Then
                If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                    cmbCondicoes.Enabled = (objPedido.SubOperacao.Financeiro And objPedido.Vencimentos.OficialMovimentado <= objPedido.Vencimentos.OficialAMovimentar)
                Else
                    cmbCondicoes.Enabled = (objPedido.SubOperacao.Financeiro And objPedido.Vencimentos.MoedaMovimentado <= objPedido.Vencimentos.MoedaAMovimentar)
                End If
            End If

            If objPedido.SubOperacao.Financeiro AndAlso objPedido.Vencimentos.Count > 0 Then
                If (objPedido IsNot Nothing AndAlso objPedido.CondicaoPagamento IsNot Nothing) Then
                    cmbCondicoes.SelectedValue = objPedido.CondicaoPagamento.Codigo
                End If

                If (objPedido IsNot Nothing AndAlso objPedido.Vencimentos IsNot Nothing) Then
                    Dim objVencimentos As New Hashtable
                    Dim i As Integer
                    For i = 0 To objPedido.Vencimentos.OrderBy(Function(s) s.Provisao).Count - 1

                        If objPedido.Vencimentos(i).Provisao = [Lib].Negocio.eProvisao.Provisao Then
                            objVencimentos.Add(i, objPedido.Vencimentos(i).Codigo)
                        End If

                        If objPedido.Vencimentos(i).Provisao = [Lib].Negocio.eProvisao.Baixa OrElse Not String.IsNullOrWhiteSpace(objPedido.Vencimentos(i).ContratoBancario) Then
                            cmdConsultarCliente.Enabled = False
                            btnConsultaPraca.Enabled = False
                            cmdConsultaDadosBancarios.Enabled = False
                            lnkExcluir.Parent.Visible = False

                            ddlMomentoFinanceiro.Enabled = False
                            cmbCondicoes.Enabled = False
                            lstCondicoesPgtoEntrega.Enabled = False
                            txtQuotaDeEntrega.Enabled = False
                            ddlPeriodicidadeEntrega.Enabled = False
                        End If
                        If Not String.IsNullOrWhiteSpace(objPedido.Vencimentos(i).ContratoBancario) Then cmbCondicoes.Enabled = False
                    Next
                    Session("objPedVencimentos" & HID.Value) = objVencimentos
                    GridCondicoes.DataSource = objPedido.Vencimentos.ToDataTablePedidos()
                    GridCondicoes.DataBind()

                    If objPedido.Vencimentos.Any(Function(s) s.Situacao = 101 OrElse s.Situacao = 102) Then
                        lnkAtualizar.Parent.Visible = False
                        lnkExcluir.Parent.Visible = False
                        cmbCondicoes.Enabled = False
                        btnOkVencimento.Enabled = False

                        MsgBox(Me.Page, "Pedido com Financeiro em Cobrança Bancária não pode ser Alterado, apenas para Consulta.", eTitulo.Info)
                    End If
                End If
            End If

            'txtTotalPacelado.Text = objPedido.Vencimentos.OficialAMovimentar.ToString("N2")
            txtTotalPacelado.Text = (objPedido.Vencimentos.OficialAMovimentar - objPedido.Vencimentos.OficialADevolver).ToString("N2")
            txtTotalPago.Text = (objPedido.Vencimentos.OficialMovimentado - objPedido.Vencimentos.OficialDevolvido).ToString("N2")
            'txtSaldoVencimentos.Text = (objPedido.Vencimentos.OficialAMovimentar - (objPedido.Vencimentos.OficialMovimentado - objPedido.Vencimentos.OficialDevolvido)).ToString("N2")
            txtSaldoVencimentos.Text = (objPedido.Vencimentos.OficialAMovimentar - objPedido.Vencimentos.OficialADevolver - (objPedido.Vencimentos.OficialMovimentado - objPedido.Vencimentos.OficialDevolvido)).ToString("N2")

            '*****************************
            '***** Tab Pedido Troca ******
            '*****************************
            If objPedido.Troca AndAlso objPedido.Operacao.CodigoClasse = eClassesOperacoes.COMPRAS.ToString Then
                ddl.Carregar(ddlContaAdto, CarregarDDL.Tabela.PlanoDeContas, " Adiantamento = 1 AND len(Conta_Id) > 6 AND left(Conta_Id,1) = '1' or conta_id ='" & objPedido.ContaAdiantamentoTroca & "'")
                ddlContaAdto.SelectedValue = objPedido.ContaAdiantamentoTroca
                TabPedidoTroca.Visible = True
            ElseIf objPedido.Troca AndAlso objPedido.Operacao.CodigoClasse = eClassesOperacoes.VENDAS.ToString Then
                ddl.Carregar(ddlContaAdto, CarregarDDL.Tabela.PlanoDeContas, " Adiantamento = 1 AND len(Conta_Id) > 6 AND left(Conta_Id,1) = '2' or conta_id ='" & objPedido.ContaAdiantamentoTroca & "'")
                ddlContaAdto.SelectedValue = objPedido.ContaAdiantamentoTroca
                TabPedidoTroca.Visible = True
            Else
                ddlContaAdto.Items.Clear()
                TabPedidoTroca.Visible = False
            End If

            If Not objPedido.PedidoTroca Is Nothing AndAlso objPedido.PedidoTroca.Troca Then
                Dim itemEmpresaTroca As ListItem = Funcoes.FormatarListItemCliente(objPedido.PedidoTroca.Empresa)
                lblEmpresaTroca.Text = itemEmpresaTroca.Text

                Dim itemClienteTroca As ListItem = Funcoes.FormatarListItemCliente(objPedido.PedidoTroca.Cliente)
                lblClienteTroca.Text = itemClienteTroca.Text

                lblPedidoTroca.Text = objPedido.CodigoPedidoTroca

                lblCNTroca.Text = objPedido.PedidoTroca.PedidoEfetivo

                chkTroca.Enabled = False
                LnkVincular.Parent.Visible = False

                If objPedido.Troca AndAlso objPedido.Operacao.CodigoClasse = eClassesOperacoes.COMPRAS.ToString Then
                    LnkDesvincular.Parent.Visible = True
                    txtVendaTroca.Text = objPedido.PedidoTroca.Itens.LiquidoOficial.ToString("N2")
                    txtCompraTroca.Text = objPedido.Itens.LiquidoOficial.ToString("N2")
                    txtSaldoTroca.Text = (objPedido.PedidoTroca.Itens.LiquidoOficial - objPedido.Itens.LiquidoOficial).ToString("N2")
                Else
                    txtVendaTroca.Text = objPedido.Itens.LiquidoOficial.ToString("N2")
                    txtCompraTroca.Text = objPedido.PedidoTroca.Itens.LiquidoOficial.ToString("N2")
                    txtSaldoTroca.Text = (objPedido.Itens.LiquidoOficial - objPedido.PedidoTroca.Itens.LiquidoOficial).ToString("N2")
                End If
            Else
                LnkVincular.Parent.Visible = objPedido.Operacao.CodigoClasse = eClassesOperacoes.COMPRAS.ToString
            End If

            '*****************************
            '***** Tab Observações *******
            '*****************************
            txtObservacao.Text = objPedido.Observacoes
            SessaoSalvaPedido()

            '*****************************
            '***** Tab Parcelas **********
            '*****************************
            If objPedido.CodigoCondicaoPagamento = 0 Then
                ddlCondPagPed.SelectedIndex = 0
            Else
                ddlCondPagPed.SelectedValue = objPedido.CodigoCondicaoPagamento
            End If

            txtDataCondPagParcela.Text = objPedido.DataPedido.ToString("dd/MM/yyyy")

            gridParcelas.DataSource = objPedido.Parcelas
            gridParcelas.DataBind()

            '(From x In objPedido.Financeiro Select New With {.Origem = x.Origem}).Distinct()
            Dim result = (From x In objPedido.Financeiro Select New With {.Origem = x.Origem}).Distinct().ToList()

            GridOrigemFinanceiro.DataSource = result
            GridOrigemFinanceiro.DataBind()

            txtAbertoCP.Text = objPedido.Parcelas.ApuracaoParcela.CP_Aberto.ToString("N2")
            txtBaixadoCP.Text = objPedido.Parcelas.ApuracaoParcela.CP_Baixado.ToString("N2")
            txtAdCP.Text = objPedido.Parcelas.ApuracaoParcela.CP_Adiantamento.ToString("N2")
            txtBxAdCP.Text = objPedido.Parcelas.ApuracaoParcela.CP_BxAdiantamento.ToString("N2")

            txtAbertoCR.Text = objPedido.Parcelas.ApuracaoParcela.CR_Aberto.ToString("N2")
            txtBaixadoCR.Text = objPedido.Parcelas.ApuracaoParcela.CR_Baixado.ToString("N2")
            txtAdCR.Text = objPedido.Parcelas.ApuracaoParcela.CR_Adiantamento.ToString("N2")
            txtBxAdCR.Text = objPedido.Parcelas.ApuracaoParcela.CR_BxAdiantamento.ToString("N2")

            txtResPedido.Text = objPedido.Parcelas.ApuracaoParcela.Liquido.ToString("N2")
            txtResRealizado.Text = (objPedido.Parcelas.ApuracaoParcela.Liquido - objPedido.Parcelas.ApuracaoParcela.SaldoAGerar).ToString("N2")
            txtResSaldo.Text = objPedido.Parcelas.ApuracaoParcela.SaldoAGerar.ToString("N2")


            '*****************************
            '*****************************
            '*****************************

            '*****************************
            '******** Vencimento  *******
            '*****************************

            If objPedido.MomentoFinanceiro = 1 AndAlso
                Left(objPedido.CodigoEmpresa, 8) = "04440724" AndAlso
                objPedido.SubOperacao.Financeiro AndAlso
                objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso
                objPedido.CondicaoPagamentoDaEntrega.Codigo > 0 Then

                divEntrega.Visible = True
                lstCondicoesPgtoEntrega.Items.IndexOf(lstCondicoesPgtoEntrega.Items.FindByValue(objPedido.CondicaoPagamentoDaEntrega.Codigo.ToString()))
                txtQuotaDeEntrega.Text = objPedido.QuotaEntrega
                ddlPeriodicidadeEntrega.SelectedValue = objPedido.PeriodicidadeEntrega
            End If

            If objPedido.Empresa.Empresa.PedidoBloqueado Then
                If objPedido.PedidoBloqueado Then
                    btnBloqueado.BackColor = Drawing.Color.Red
                    btnBloqueado.Text = "Bloqueado"
                    lnkAtualizar.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False
                    MsgBox(Me.Page, "Pedido bloqueado não pode ser alterado")
                Else
                    btnBloqueado.BackColor = Drawing.Color.Green
                    btnBloqueado.Text = "Liberado"
                End If
            End If

            '*****************************
            '********  Contrato    *******
            '*****************************
            If objPedido.Contratos IsNot Nothing AndAlso objPedido.Contratos.Count > 0 Then
                gridContratos.DataSource = objPedido.Contratos
                gridContratos.DataBind()

                Dim x As Integer = 0
                While x < gridContratos.Rows.Count
                    Dim pExtensao() As String = gridContratos.Rows(x).Cells(2).Text.Split(".")

                    If pExtensao(1).ToUpper.Equals("XML") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/xml16x16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("PDF") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icopdf16X16.jpg"
                    ElseIf pExtensao(1).ToUpper.Equals("XLS") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("XLSX") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("DOC") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("DOCX") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                    Else
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icoJpg.gif"
                    End If

                    CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ToolTip = String.Format("Download - {0}.{1}", pExtensao(0), pExtensao(1))
                    CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"

                    x += 1
                End While

            End If


            '*************************************************
            '***** Tab Observações de Controle Interno *******
            '*************************************************
            If objPedido.ObservacoesControleInterno.Length > 0 Then
                TabControleInterno.Visible = True

                txtObservacaoControleInterno.Text = objPedido.ObservacoesControleInterno
            Else
                TabControleInterno.Visible = False
            End If



            If Not objPedido.CodigoSituacao = [Lib].Negocio.eSituacao.Normal Then
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
            ElseIf objPedido.FiscalAberto = False AndAlso Not Funcoes.VerificaAcesso(objPedido.Empresa.Codigo, objPedido.Empresa.CodigoEndereco, IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text)), "PEDIDOS") Then
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
                lnkAdicionarItem.Parent.Parent.Visible = False
                Dim msg As String = "Movimento " & IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text)) & " já fechado para esta data!"
                MsgBox(Me.Page, msg)
            End If

            tcPedido.ActiveTabIndex = 0

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally

        End Try
    End Sub

    Private Sub ListarPedidos()
        SessaoRecuperaPedido()

        HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
        Dim parameters As New Dictionary(Of String, Object)
        If Not String.IsNullOrWhiteSpace(HID.Value) Then parameters.Add("tipo", HID.Value)
        If txtRegistro.Text.Length > 0 Then parameters.Add("pedido", txtRegistro.Text)
        If txtPedido.Text.Length > 0 Then parameters.Add("efetivo", txtPedido.Text)
        If cmbUnidadeNegocio.SelectedIndex > 0 Then
            parameters.Add("unidade", cmbUnidadeNegocio.SelectedValue)
        End If

        If cmbEmpresa.SelectedIndex > 0 Then
            Dim Empresa() As String = cmbEmpresa.SelectedValue.ToString.Split("-")
            parameters.Add("empresa", Empresa(0))
            parameters.Add("enderecoEmpresa", Empresa(1))
        End If

        Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
        If strCliente(0) <> "" Then parameters.Add("cliente", strCliente(0))
        If strCliente.Length > 1 Then parameters.Add("enderecoCliente", strCliente(1))

        'Comentado já que o campo situação está desabilitado, então quando um pedido consultado estiver, por exemplo, cancelado sua situação irá aparecer no campo.
        'If cmbSituacao.SelectedIndex > 0 Then parameters.Add("situacao", cmbSituacao.SelectedValue)
        If cmbOperacao.SelectedIndex > 0 Then parameters.Add("operacao", cmbOperacao.SelectedValue)
        If cmbSubOperacao.SelectedIndex > 0 Then parameters.Add("suboperacao", cmbSubOperacao.SelectedValue)
        If cmbSafra.SelectedIndex > 0 Then parameters.Add("safra", cmbSafra.SelectedValue)
        If txtNaviosXInvoice.Text.Length > 0 Then parameters.Add("InvoiceNavio", txtNaviosXInvoice.Text)

        If Left(Session("ssEmpresa").ToString, 8) = "24450490" Then
            Dim gruposPrd As String = "'10101','10102','20101','20102','20103','20104','20105','20106','30101','90203'"
            parameters.Add("grupoProduto", gruposPrd)
        Else
            parameters.Add("grupoProduto", "")
        End If

        If objPedido.Itens.Count > 0 Then
            Dim Produtos As String = String.Empty
            For Each row In objPedido.Itens
                Produtos &= IIf(String.IsNullOrWhiteSpace(Produtos), "", ",") & "'" & row.CodigoProduto & "'"
            Next
            parameters.Add("produto", Produtos)
        End If

        Session("ssTipoRetorno") = "objPedidosXItens" & HID.Value
        Dim numberRows As Integer = ucConsultaPedidos.BindGridView(parameters)
        If numberRows > 1 Then
            Popup.ConsultaDePedidos(Me.Page, "objPedidosXItens" & HID.Value)
        End If

        lnkEnviarEmail.Parent.Visible = True
    End Sub

    Public Function ValidarCamposPedido() As Boolean
        Dim Mensagem As String = ""

        If txtPedido.Text.Length = 0 Then txtPedido.Text = 0

        If objPedido.Cliente.CodigoEstado = "EX" AndAlso objPedido.CodigoLocalEmbarque.Length = 0 Then
            If objPedido.SubOperacao.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida Then
                objPedido.CodigoLocalEmbarque = objPedido.CodigoEmpresa
                objPedido.EndLocalEmbarque = objPedido.EnderecoEmpresa
            Else
                objPedido.CodigoLocalEmbarque = objPedido.CodigoCliente
                objPedido.EndLocalEmbarque = objPedido.EnderecoCliente
            End If
            SessaoSalvaPedido()
        End If

        'Transferencia só pode ser dentro do mesmo grupo(filiais) da Empresa - Furlan - 12/08/2024
        If objPedido.SubOperacao.Classe = eClassesOperacoes.TRANSFERENCIAS AndAlso Not Left(objPedido.CodigoEmpresa, 8) = Left(objPedido.CodigoCliente, 8) Then
            MsgBox(Me.Page, "Transferência só pode ser feita entre Filiais da mesma Empresa.")
            Return False
        End If

        'Verifica se tem Pedido Efetivo - Furlan - 24/07/2024
        If Left(objPedido.CodigoEmpresa, 8) = "24450490" AndAlso objPedido.IUD = "I" AndAlso objPedido.PedidoEfetivo.Length > 0 AndAlso objPedido.TemPedidoEfetivo Then
            MsgBox(Me.Page, "Pedido Efetivo já existe e não pode ser utilizado novamente.")
            Return False
        End If

        'Verifica se tem Pedido Efetivo - Furlan - 29/07/2024
        If Left(objPedido.CodigoEmpresa, 8) = "24450490" AndAlso objPedido.IUD = "I" AndAlso objPedido.Contrato.Length > 0 AndAlso objPedido.TemPedidoContrato Then
            MsgBox(Me.Page, "Pedido MIC já existe e não pode ser utilizado novamente.")
            Return False
        End If

        If objPedido.IUD = "I" And objPedido.SubOperacao.Financeiro And FinanceiroVirtual Then
            objPedido.MomentoFinanceiro = 9
        End If

        If objPedido.SubOperacao.Representante AndAlso objPedido.Representantes.Count = 0 Then
            MsgBox(Me.Page, "Tabela de Comissão do Representante não foi informado")
            Return False
        End If

        If cmbEmbalagem.SelectedValue.Length = 0 And Left(objPedido.CodigoEmpresa, 8) = "24450490" Then
            MsgBox(Me.Page, "A informação de embalagem é obrigadoria!")
            Return False
        End If

        If rbFOB.Checked And txtLocalDeEmbarque.Text.Length = 0 And Left(objPedido.CodigoEmpresa, 8) = "24450490" Then
            MsgBox(Me.Page, "Quando o frete é FOB é necessario informar o Local de Embarque!")
            Return False
        End If

        'REMOVIDO POR NÃO ESTAR SENDO USADO - FURLAN - 23/08/2024
        'For Each rep As [Lib].Negocio.PedidoXRepresentante In objPedido.Representantes
        '    If Not rep.PercentualFixo AndAlso Not rep.CodigoRepresentante = objPedido.CodigoEmpresa Then
        '        For Each comissao As [Lib].Negocio.PedidoXRepresentantesxTabelaDeComissao In rep.Comissoes.Where(Function(s) s.IUD <> "D")
        '            If comissao.CodigoTabela = 0 Then
        '                MsgBox(Me.Page, "Tabela de Comissão do Representante " & rep.Representante.Nome & " não foi selecionada")
        '                Return False
        '            End If
        '        Next
        '    End If
        'Next

        Dim numdep As Integer
        numdep = objPedido.Depositos.Where(Function(s) s.Tipo = "DE").Count
        If numdep = 0 Then
            MsgBox(Me.Page, "Informe o Deposito do Pedido")
            Return False
        End If

        numdep = objPedido.Depositos.Where(Function(s) s.Tipo = "OD").Count
        If numdep = 0 Then
            MsgBox(Me.Page, "Informe o Deposito de Origem Destino")
            Return False
        End If

        If objPedido.SubOperacao.ProprietarioDaMercadoria Then
            numdep = objPedido.Depositos.Where(Function(s) s.Tipo = "PM").Count
            If numdep = 0 Then
                MsgBox(Me.Page, "Informe o Proprietário da Mercadoria")
                Return False
            End If
        End If

        If objPedido.CodigoFinalidade = 3 AndAlso Not (objPedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS And objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada And objPedido.FreteCIFFOB = eTiposFrete.FOB) Then
            MsgBox(Me.Page, "Pedidos na finalidade 3 - 'DEPOSITO, ORIGEM VENDA, FRETE NOSSO' devem ser usado em operacoes da classe deposito de entrada e frete por conta da Nossa empresa.")
            Return False
        End If

        If objPedido.CodigoFinalidade = 4 AndAlso objPedido.Representantes IsNot Nothing AndAlso objPedido.Representantes.Where(Function(s) s.IUD <> "D").Count > 0 Then
            MsgBox(Me.Page, "Pedidos na finalidade 4 - 'EMPRESTIMO PRODUTO' nao podem ter representantes comissionados.")
            Return False
        End If

        ''DESABILITANDO A OBRIGAÇÃO DO ROTEIRO
        ''If objPedido.Empresa.Empresa.FretePedido _
        ''    AndAlso (objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida And objPedido.FreteCIFFOB = eTiposFrete.CIF _
        ''        OrElse objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada And objPedido.FreteCIFFOB = eTiposFrete.FOB) Then

        ''    'Saida - Cif
        ''    'Entrada - Fob

        ''    If objPedido.Roteiros Is Nothing OrElse objPedido.Roteiros.Count = 0 Then
        ''        MsgBox(Me.Page, "O Roteiro de Frete é obrigatório para Pedidos de Entrada com Frete FOB e Pedidos de Saída com Frete CIF.")
        ''        tcPedido.ActiveTab = TabDeposito
        ''        Return False
        ''    End If

        ''    If Not objPedido.Roteiros Is Nothing AndAlso objPedido.Roteiros.TotalQtd <> QuantidadeTotalPedido() Then
        ''        MsgBox(Me.Page, "A Quantidade total informada no roteiro " & objPedido.Roteiros.TotalQtd.ToString("N2") & " deve ser igual a quantidade total dos ítens do pedido " & QuantidadeTotalPedido.ToString("N2") & ".")
        ''        Return False
        ''    End If

        ''    If objPedido.Roteiros.Where(Function(r) r.QuantidadeAtual > 0 And r.Valor <= 0.0).Any Then
        ''        MsgBox(Me.Page, "Existe(m) roteiro(s) com quantidade porém sem o valor.")
        ''        Return False
        ''        'Verificar, pois nos pedidos a fixar o frete será lançado inicialmente sem quantidade definida
        ''        'ElseIf objPedido.Roteiros.Where(Function(r) r.QuantidadeAtual = 0 And r.Valor > 0.0).Any Then
        ''        '    MsgBox(Me.Page, "Existe(m) roteiro(s) com valor porém sem quantidade.")
        ''        '    Return False
        ''    End If

        ''End If

        'Validações para Trocas
        Mensagem = ValidaTroca(objPedido)
        If Not String.IsNullOrWhiteSpace(Mensagem) Then
            MsgBox(Me.Page, Mensagem)
            Return False
        End If

        'Validação de Representantes para RTGrãos conforme solicitação e-mail Matheus - Furlan - 26/08/2024
        If Left(objPedido.CodigoEmpresa, 8) = "24450490" AndAlso
            GridRepresentantes.Rows.Count = 0 AndAlso
            (objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20101") OrElse
             objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20102") OrElse
             objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20103") OrElse
             objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20104") OrElse
             objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20105") OrElse
             objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20106")) AndAlso
            (objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
             objPedido.SubOperacao.Classe = eClassesOperacoes.VENDASAORDEM OrElse
             objPedido.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES) Then
            MsgBox(Me.Page, "Representante é obrigatório")
            Return False
        End If

        If cmbUnidadeNegocio.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório")
            Return False
        ElseIf cmbOperacao.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Operação é obrigatório")
            Return False
        ElseIf cmbSubOperacao.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "SubOperaçao é obrigatório")
            Return False
        ElseIf cmbSituacao.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Situação é obrigatório")
            Return False
        ElseIf txtCliente.Text = "" Then
            MsgBox(Me.Page, "Cliente é obrigatório")
            Return False
        ElseIf txtPraca.Text = "" Then
            MsgBox(Me.Page, "Praça de pagamento é obrigatório")
            Return False
        ElseIf cmbSafra.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Safra é obrigatório")
            Return False
        ElseIf txtDataEntregaInicial.Text.Length = 0 Then
            MsgBox(Me.Page, "Data de entrega Inicial é obrigatório")
            Return False
        ElseIf txtDataEntregaFinal.Text.Length = 0 Then
            MsgBox(Me.Page, "Data de entrega Final é obrigatório")
            Return False
        ElseIf txtDataPedido.Text.Length = 0 Then
            MsgBox(Me.Page, "Data do pedido é obrigatório")
            Return False
        ElseIf objPedido.Itens.Count = 0 Then
            MsgBox(Me.Page, "Informe ao menos um item para adicionar o pedido")
            Return False
        ElseIf Not ValidarDepositos() Then
            Return False
        ElseIf cmbFinalidade.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Finalidade é obrigatório")
            Return False
        ElseIf objPedido.MomentoFinanceiro <> 9 AndAlso (Not objPedido.Troca Or Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "15204808") AndAlso objPedido.SubOperacao.Financeiro AndAlso objPedido.SubOperacao.PrecoFixo AndAlso objPedido.Moeda.Classificacao = eTiposMoeda.Oficial AndAlso objPedido.Itens.LiquidoOficial > 0 AndAlso objPedido.Vencimentos.Count = 0 Then
            MsgBox(Me.Page, "Condição de Pagamento não foi selecionada")
            Return False
        ElseIf (Not objPedido.Troca Or Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "15204808") AndAlso objPedido.SubOperacao.Financeiro AndAlso objPedido.SubOperacao.PrecoFixo AndAlso objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso objPedido.Itens.LiquidoMoeda > 0 AndAlso objPedido.Vencimentos.Count = 0 Then
            MsgBox(Me.Page, "Condição de Pagamento não foi selecionada")
            Return False
        ElseIf objPedido.MomentoFinanceiro <> 9 AndAlso (Not objPedido.Troca Or Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "15204808") AndAlso objPedido.SubOperacao.Financeiro AndAlso objPedido.SubOperacao.PrecoFixo AndAlso objPedido.Moeda.Classificacao = eTiposMoeda.Oficial AndAlso Math.Abs(objPedido.Itens.LiquidoOficial - Math.Round(objPedido.Vencimentos.OficialAMovimentar - objPedido.Vencimentos.OficialADevolver, 2, MidpointRounding.AwayFromZero)) > 0.01 Then
            MsgBox(Me.Page, "O Valor Financeiro nao Confere com o valor programado nos itens do pedido")
            Return False
        ElseIf objPedido.MomentoFinanceiro <> 9 AndAlso (Not objPedido.Troca Or Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "15204808") AndAlso objPedido.SubOperacao.Financeiro AndAlso objPedido.SubOperacao.PrecoFixo AndAlso objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso Math.Abs(objPedido.Itens.LiquidoMoeda - Math.Round(objPedido.Vencimentos.MoedaAMovimentar - objPedido.Vencimentos.MoedaADevolver, 2, MidpointRounding.AwayFromZero)) > 0.01 Then
            MsgBox(Me.Page, "O Valor Financeiro nao Confere com o valor programado nos itens do pedido")
            Return False
        ElseIf objPedido.MomentoFinanceiro <> 9 AndAlso objPedido.SubOperacao.Financeiro AndAlso Not objPedido.MomentoFinanceiro = 0 AndAlso objPedido.Vencimentos.Count = 0 AndAlso FinanceiroVirtual Then
            MsgBox(Me.Page, "Financeiro não foi selecionado")
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(txtRepresentante.Text) AndAlso GridRepresentantes.Rows.Count = 0 Then
            MsgBox(Me.Page, "Campo Representate está preenchido mas não está adicionado ao pedido!")
            Return False
        ElseIf Not rbCIF.Checked And Not rbFOB.Checked And Not rbTER.Checked And Not rbNenhum.Checked Then
            MsgBox(Me.Page, "O campo CIF, FOB, TERCEIRO ou Nenhum não foi selecionado!")
            Return False
            'ElseIf objPedido.Empresa.Empresa.ObrigaNavio AndAlso txtNaviosXInvoice.Text.Length = 0 Then  '''DESABILITADO ATÉ ESTABILIZAR A IMPLANTAÇÃO
            '    MsgBox(Me.Page, "Invoice não foi selecionada!")
            '    Return False
        ElseIf cmbMoeda.SelectedValue = 3 And cmbIndexador.SelectedValue = "99" And txtIndiceFixado.Text = 0 Then
            MsgBox(Me.Page, "Pedido em dolar e indexador 99 não pode ter cotação zero!")
            Return False
        Else
            Return True
        End If

    End Function

    Private Function ValidaTroca(ByRef Pedido As Pedido) As String
        Dim Retorno As String = String.Empty

        If Pedido.PedidoTroca IsNot Nothing AndAlso Pedido.Troca Then
            If objPedido.Operacao.CodigoClasse.Equals(eClassesOperacoes.COMPRAS.ToString) AndAlso
                         ((Pedido.CodigoMoeda = eTiposMoeda.Oficial AndAlso Math.Abs(objPedido.PedidoTroca.Itens.LiquidoOficial - objPedido.Itens.LiquidoOficial) > 0.01) OrElse
                         (Pedido.CodigoMoeda = eTiposMoeda.MoedaEstrangeira AndAlso Math.Abs(objPedido.PedidoTroca.Itens.LiquidoOficial - objPedido.Itens.LiquidoOficial) > 0.1)) Then
                If Pedido.Operacao.CodigoClasse.Equals(eClassesOperacoes.VENDAS.ToString) AndAlso
                                Pedido.Itens.LiquidoOficial - Pedido.PedidoTroca.Itens.LiquidoOficial < 0 AndAlso
                                Not Funcoes.VerificaPermissao("AJUSTARPEDIDOTROCA", "GRAVAR") Then
                    Retorno = "Valor da Soma das compras de troca é superior a Venda, apenas Usuários com Processo AJUSTARPEDIDOTROCA e permissão de GRAVAR pode fazer essa alteração."
                ElseIf Pedido.Operacao.CodigoClasse.Equals(eClassesOperacoes.COMPRAS.ToString) AndAlso
                               Pedido.PedidoTroca.Itens.LiquidoOficial - Pedido.Itens.LiquidoOficial < 0 AndAlso
                               Not Funcoes.VerificaPermissao("AJUSTARPEDIDOTROCA", "GRAVAR") Then
                    Retorno = "Valor da Soma das compras de troca é inferior a Venda, apenas Usuários com Processo AJUSTARPEDIDOTROCA e permissão de GRAVAR pode fazer essa alteração."
                End If
            End If

        ElseIf Pedido.Troca AndAlso Pedido.Operacao.CodigoClasse.Equals(eClassesOperacoes.COMPRAS.ToString) Then
            Retorno = "Pedido de troca não foi selecionado!"
        End If
        Return Retorno
    End Function

    Private Function QuantidadeTotalPedido() As Decimal
        Dim QtdTotalPedido As Decimal

        If objPedido Is Nothing Then SessaoRecuperaPedido()

        If objPedido.Itens.QuantidadeTotal > 0 Then
            QtdTotalPedido = objPedido.Itens.QuantidadeTotal
        Else
            QtdTotalPedido = objPedido.Itens.Sum(Function(s) s.SaldoItem.QtdeContratadoFiscal)
        End If
        Return QtdTotalPedido
    End Function
#End Region

#Region "Linha de Botoes"

    'Private Sub ImprimirPedido(ByVal pdf As Boolean)
    '    Dim ds As New DataSet
    '    Dim Sql As String

    '    If objPedido Is Nothing OrElse objPedido.Codigo = 0 Then
    '        MsgBox(Me.Page, "Pedido não foi informado.", eTitulo.Info)
    '        Exit Sub
    '    End If

    '    Sql = "SELECT Nome, Cliente_Id as Codigo, Inscricao, Endereco,complemento, Telefone, CEP, Cidade, Estado " & vbCrLf & _
    '          "  FROM Clientes" & vbCrLf & _
    '          " Where Cliente_id    ='" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
    '          "   and Endereco_id = " & objPedido.EnderecoEmpresa
    '    ds.Merge(Banco.ConsultaDataSet(Sql, "Empresa"))

    '    Sql = "SELECT Nome, Cliente_Id as Codigo, Inscricao, Endereco,complemento, Telefone, CEP, Cidade, Estado " & vbCrLf & _
    '          "  FROM Clientes" & vbCrLf & _
    '          " Where Cliente_id    ='" & objPedido.CodigoCliente & "'" & vbCrLf & _
    '          "   and Endereco_id = " & objPedido.EnderecoCliente
    '    ds.Merge(Banco.ConsultaDataSet(Sql, "Cliente"))

    '    Sql = "SELECT Nome, Cliente_Id as Codigo, Inscricao, Endereco,complemento, Telefone, CEP, Cidade, Estado " & vbCrLf & _
    '          "  FROM Clientes" & vbCrLf & _
    '          " Where Cliente_id    ='" & objPedido.CodigoPraca & "'" & vbCrLf & _
    '          "   and Endereco_id = " & objPedido.EnderecoPraca
    '    ds.Merge(Banco.ConsultaDataSet(Sql, "Praca"))

    '    Sql = "SELECT P.Pedido_Id as Codigo," & vbCrLf & _
    '          "       P.DataPedido," & vbCrLf & _
    '          "       P.Vendedor," & vbCrLf & _
    '          "       case When M.Classificacao = 'O' then PxI.TotalBrutoOficial   else PxI.TotalBrutoMoeda   end as TotalBruto," & vbCrLf & _
    '          "       case When M.Classificacao = 'O' then PxI.TotalLiquidoOficial else PxI.TotalLiquidoMoeda end as TotalLiquido," & vbCrLf & _
    '          "       P.Observacoes," & vbCrLf & _
    '          "       P.BancoCliente," & vbCrLf & _
    '          "       isnull(B.Descricao,'') as NomeBanco," & vbCrLf & _
    '          "       P.AgenciaCliente," & vbCrLf & _
    '          "       P.DigitoAgenciaCliente," & vbCrLf & _
    '          "       P.ContaCliente," & vbCrLf & _
    '          "       P.DigitoContaCliente," & vbCrLf & _
    '          "       SO.Classe," & vbCrLf & _
    '          "       cast(so.operacao_id as varchar) + '-' + cast(so.SubOperacoes_Id as varchar) + ' - ' + so.Descricao + ' - (' + so.Classe + ')' as Operacao," & vbCrLf & _
    '          "       P.Moeda, " & vbCrLf & _
    '          "       P.FreteCIFFOB" & vbCrLf & _
    '          "  FROM Pedidos P" & vbCrLf & _
    '          " Inner Join (Select Empresa_id," & vbCrLf & _
    '          "                    EndEmpresa_Id," & vbCrLf & _
    '          "                    Pedido_Id," & vbCrLf & _
    '          "                    sum(case When Encargo_Id = 'PRODUTO' Then valorOficial else 0 end) as TotalBrutoOficial," & vbCrLf & _
    '          "                    sum(case When Encargo_Id = 'PRODUTO' Then valorMoeda   else 0 end) as TotalBrutoMoeda," & vbCrLf & _
    '          "                    sum(case When Encargo_Id = 'LIQUIDO' Then valorOficial else 0 end) as TotalLiquidoOficial," & vbCrLf & _
    '          "                    sum(case When Encargo_Id = 'LIQUIDO' Then valorMoeda   else 0 end) as TotalLiquidoMoeda" & vbCrLf & _
    '          "               From pedidosxencargos" & vbCrLf & _
    '          "              GRoup by Empresa_id, EndEmpresa_Id, Pedido_Id " & vbCrLf & _
    '          "             ) PxI" & vbCrLf & _
    '          "    on P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
    '          "   and P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf & _
    '          "   and P.Pedido_id     = PxI.Pedido_Id" & vbCrLf & _
    '          "  Left Join Bancos B" & vbCrLf & _
    '          "    on P.BancoCliente = B.Banco_Id" & vbCrLf & _
    '          " Inner Join SubOperacoes SO" & vbCrLf & _
    '          "    on SO.Operacao_Id     = P.Operacao" & vbCrLf & _
    '          "   and SO.SubOperacoes_Id = P.SubOperacao" & vbCrLf & _
    '          " Inner join Moedas M" & vbCrLf & _
    '          "    on M.Moeda_Id = P.Moeda" & vbCrLf & _
    '          " Where P.Empresa_Id    ='" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
    '          "   and P.EndEmpresa_Id = " & objPedido.EnderecoEmpresa & vbCrLf & _
    '          "   and P.Pedido_Id     = " & objPedido.Codigo
    '    ds.Merge(Banco.ConsultaDataSet(Sql, "Pedido"))

    '    Sql = "   SELECT		co.Representante_Id + '-' + CAST(co.EndRepresentante_Id AS varchar) + ' - ' + c.Nome AS Representante,                         " & vbCrLf & _
    '        "   			CONVERT(DECIMAL(18, 10), co.Percentual) AS Percentual, co.ValorComissao, co.Principal,                                         " & vbCrLf & _
    '        "   			ISNULL(co.PercentualFixo, CASE WHEN co.valorcomissao > 0 THEN 1 ELSE 0 END) AS PercentualFixo, co.Pedido_Id  " & vbCrLf & _
    '        "   FROM         Comissoes AS co                                                                                                               " & vbCrLf & _
    '        "   		INNER JOIN Clientes AS c                                                                                                           " & vbCrLf & _
    '        "   			ON c.Cliente_Id = co.Representante_Id                                                                                                " & vbCrLf & _
    '        "   			AND c.Endereco_Id = co.EndRepresentante_Id                                                                                           " & vbCrLf & _
    '        "   WHERE   (co.Empresa_Id = '" & objPedido.CodigoEmpresa & "') " & vbCrLf & _
    '        "           AND (co.EndEmpresa_Id = " & objPedido.EnderecoEmpresa & ")" & vbCrLf & _
    '        "           AND (co.Pedido_Id = " & objPedido.Codigo & ")                                      " & vbCrLf

    '    ds.Merge(Banco.ConsultaDataSet(Sql, "Representantes"))

    '    Sql = "SELECT co.Transportador_Id + '-' + cast(co.EndTransportador_Id as varchar) + ' - ' + c.Nome as Transportador, co.Quantidade, co.QuotaDiaria, co.Redespacho, co.DataFrete_Id, isnull(co.UnitarioFrete,0) AS UnitarioFrete " & vbCrLf & _
    '              "  FROM PedidosXTransportadores co" & vbCrLf & _
    '              "   	INNER JOIN Clientes AS c                                                                                                           " & vbCrLf & _
    '              "   		ON c.Cliente_Id = co.Transportador_Id                                                                                                " & vbCrLf & _
    '              "		    AND c.Endereco_Id = co.EndTransportador_Id                                                                                           " & vbCrLf & _
    '              "     WHERE   co.Empresa_Id    ='" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
    '              "         AND co.EndEmpresa_Id = " & objPedido.EnderecoEmpresa & vbCrLf & _
    '              "         AND co.Pedido_Id     = " & objPedido.Codigo & vbCrLf
    '    ds.Merge(Banco.ConsultaDataSet(Sql, "Transportadores"))

    '    Sql = "SELECT PxI.Pedido_Id as Pedido," & vbCrLf & _
    '          "       sum(case" & vbCrLf & _
    '          "             When PxI.TipodeLancamento = 'E'" & vbCrLf & _
    '          "               then PxI.Quantidade * -1" & vbCrLf & _
    '          "               else PxI.Quantidade" & vbCrLf & _
    '          "           end) as Quantidade," & vbCrLf & _
    '          "       Prd.Unidade," & vbCrLf

    '    If Left(objPedido.CodigoEmpresa, 8) = "03189063" Then
    '        Sql &= "       Prd.Nome + '-' + Prd.Descricao as DescricaoProduto," & vbCrLf
    '    Else
    '        Sql &= "       Prd.Nome as DescricaoProduto," & vbCrLf
    '    End If

    '    Sql &= "       sum(case" & vbCrLf & _
    '          "            when PxI.TipodeLancamento = 'E'" & vbCrLf & _
    '          "             then" & vbCrLf & _
    '          "               case" & vbCrLf & _
    '          "                 When M.classificacao = 'O'" & vbCrLf & _
    '          "                   then TotalOficial * -1" & vbCrLf & _
    '          "                   else TotalMoeda * - 1" & vbCrLf & _
    '          "               end" & vbCrLf & _
    '          "             else" & vbCrLf & _
    '          "               case" & vbCrLf & _
    '          "                 When M.classificacao = 'O'" & vbCrLf & _
    '          "                   then TotalOficial " & vbCrLf & _
    '          "                   else TotalMoeda " & vbCrLf & _
    '          "               end           end" & vbCrLf & _
    '          "          ) as Total," & vbCrLf & _
    '          "       PxI.Produto_Id as CodigoProduto" & vbCrLf & _
    '          "  INTO #Temp " & vbCrLf & _
    '          "  FROM Pedidos P" & vbCrLf & _
    '          " INNER JOIN PedidoXItemXLancamento PxI" & vbCrLf & _
    '          "    ON P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
    '          "   AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf & _
    '          "   AND P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf & _
    '          " Inner Join Produtos Prd" & vbCrLf & _
    '          "    on PxI.Produto_id = Prd.Produto_id" & vbCrLf & _
    '          " Inner Join Moedas M" & vbCrLf & _
    '          "    on M.Moeda_Id = P.Moeda" & vbCrLf & _
    '          " Where P.Empresa_Id    ='" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
    '          "   and P.EndEmpresa_Id = " & objPedido.EnderecoEmpresa & vbCrLf & _
    '          "   and P.Pedido_Id     = " & objPedido.Codigo & vbCrLf & _
    '          " group by PxI.Pedido_Id," & vbCrLf & _
    '          "          Prd.Unidade," & vbCrLf

    '    If Left(objPedido.CodigoEmpresa, 8) = "03189063" Then
    '        Sql &= "       Prd.Nome + '-' + Prd.Descricao," & vbCrLf
    '    Else
    '        Sql &= "          Prd.Nome," & vbCrLf
    '    End If

    '    Sql &= "		  PxI.Produto_Id" & vbCrLf & _
    '          " --Composicao do Unitario " & vbCrLf & _
    '          "SELECT Pedido, CodigoProduto,  DescricaoProduto, Unidade, Quantidade, " & vbCrLf & _
    '          "       CASE " & vbCrLf & _
    '          "           WHEN Quantidade>0 " & vbCrLf & _
    '          "			       THEN Total / Quantidade " & vbCrLf & _
    '          "				   ELSE 0 " & vbCrLf & _
    '          "		  END Unitario,  " & vbCrLf & _
    '          "       Total" & vbCrLf & _
    '          "  FROM #TEMP" & vbCrLf

    '    ds.Merge(Banco.ConsultaDataSet(Sql, "PedidosXItens"))

    '    Sql = "SELECT P.Pedido_id as Pedido," & vbCrLf & _
    '          "       PxE.Encargo_Id as Encargo," & vbCrLf & _
    '          "       sum(case" & vbCrLf & _
    '          "             when M.Classificacao = 'O'" & vbCrLf & _
    '          "               Then PxE.ValorOficial" & vbCrLf & _
    '          "               Else PxE.ValorMoeda" & vbCrLf & _
    '          "           end" & vbCrLf & _
    '          "           ) as Total" & vbCrLf & _
    '          "  FROM Pedidos P" & vbCrLf & _
    '          " INNER JOIN PedidosXEncargos PxE" & vbCrLf & _
    '          "    ON P.Empresa_Id    = PxE.Empresa_Id" & vbCrLf & _
    '          "   AND P.EndEmpresa_Id = PxE.EndEmpresa_Id" & vbCrLf & _
    '          "   AND P.Pedido_Id     = PxE.Pedido_Id" & vbCrLf & _
    '          " INNER JOIN Moedas M" & vbCrLf & _
    '          "    ON P.Moeda = M.Moeda_Id" & vbCrLf & _
    '          " Where P.Empresa_Id    ='" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
    '          "   and P.EndEmpresa_Id = " & objPedido.EnderecoEmpresa & vbCrLf & _
    '          "   and P.Pedido_Id     = " & objPedido.Codigo & vbCrLf & _
    '          " Group By P.Pedido_Id, PxE.Encargo_Id" & vbCrLf & _
    '          " having sum(case when M.Classificacao = 'O' Then PxE.ValorOficial Else PxE.ValorMoeda end) > 0" & vbCrLf & _
    '          " order by Case" & vbCrLf & _
    '          "            when PxE.Encargo_Id = 'PRODUTO' THEN 1" & vbCrLf & _
    '          "            When PxE.Encargo_Id = 'LIQUIDO' THEN 3" & vbCrLf & _
    '          "            else 2" & vbCrLf & _
    '          "          end " & vbCrLf

    '    ds.Merge(Banco.ConsultaDataSet(Sql, "PedidosXEncargos"))

    '    Sql = "Select P.Pedido_id as Pedido," & vbCrLf & _
    '          "       Sb.Prorrogacao AS Vencimento," & vbCrLf & _
    '          "       case" & vbCrLf & _
    '          "          When M.Classificacao = 'O'" & vbCrLf & _
    '          "            then sb.ValorLiquido" & vbCrLf & _
    '          "            else sb.MoedaValorLiquido" & vbCrLf & _
    '          "       end Valor" & vbCrLf & _
    '          "  From Pedidos P" & vbCrLf & _
    '          " Inner Join" & vbCrLf & _
    '          "      (Select cr.Empresa, cr.EndEmpresa, cr.Registro_id, cr.Provisao, cr.Pedido, cr.Prorrogacao, cr.ValorLiquido, cr.MoedaValorLiquido" & vbCrLf & _
    '          "	   from ContasaReceber cr" & vbCrLf & _
    '          "			inner join ComprasXProdutos cXp" & vbCrLf & _
    '          "					on cXp.Produto_Id = cr.Carteira " & vbCrLf & _
    '          "	   WHERE cr.Situacao  = 1 and not (cXp.Adiantamento = 'S' and cXp.BaixaAdiantamento = 0)  " & vbCrLf & _
    '          "          union all" & vbCrLf & _
    '          "       Select cr.Empresa, cr.EndEmpresa, cr.Registro_id, cr.Provisao, cr.Pedido, cr.Prorrogacao, cr.ValorLiquido, cr.MoedaValorLiquido" & vbCrLf & _
    '          "	   from ContasaPagar cr" & vbCrLf & _
    '          "			inner join ComprasXProdutos cXp" & vbCrLf & _
    '          "					on cXp.Produto_Id = cr.Carteira " & vbCrLf & _
    '          "	   WHERE cr.Situacao  = 1 and not (cXp.Adiantamento = 'S' and cXp.BaixaAdiantamento = 0)  " & vbCrLf & _
    '          "       ) sb" & vbCrLf & _
    '          "    on sb.Empresa    = P.Empresa_id" & vbCrLf & _
    '          "   and Sb.EndEmpresa = P.EndEmpresa_id" & vbCrLf & _
    '          "   and sb.Pedido     = P.Pedido_id" & vbCrLf & _
    '          " Inner join Moedas M" & vbCrLf & _
    '          "    on M.Moeda_Id = P.Moeda" & vbCrLf & _
    '          " Where P.Empresa_Id    ='" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
    '          "   and P.EndEmpresa_Id = " & objPedido.EnderecoEmpresa & vbCrLf & _
    '          "   and P.Pedido_Id     = " & objPedido.Codigo & vbCrLf & _
    '          " ORDER BY Sb.Provisao, Vencimento" & vbCrLf


    '    ds.Merge(Banco.ConsultaDataSet(Sql, "PedidoXParcelas"))

    '    'Imagem
    '    Dim dtImagem As DataTable = ds.Tables.Add("Images")
    '    dtImagem.Columns.Add("path", GetType(String))
    '    dtImagem.Columns.Add("image", GetType(System.Byte()))

    '    Dim drImagem As DataRow = dtImagem.NewRow()
    '    Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & objPedido.Empresa.Imagem)

    '    drImagem("path") = strCaminhoImagem
    '    drImagem("image") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
    '    dtImagem.Rows.Add(drImagem)

    '    Dim param As New Dictionary(Of String, Object)
    '    param.Add("Rep", IIf(ds.Tables("Representantes").Rows.Count > 0, True, False))
    '    param.Add("Transp", IIf(ds.Tables("Transportadores").Rows.Count > 0, True, False))

    '    param.Add("UsuarioInclusao", "Usuário Inclusao: " & objPedido.UsuarioInclusao)

    '    param.Add("CliouFor", IIf(objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "Fornecedor:", "Cliente:"))

    '    Dim crptRelatorio As New ReportDocument()

    '    Try
    '        Dim strCaminho As String = Server.MapPath("~/Reports/Cr_Pedidos.rpt")
    '        crptRelatorio.Load(strCaminho)

    '        Dim strNomeArquivo As String = String.Empty
    '        If pdf Then
    '            strNomeArquivo = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
    '        Else
    '            strNomeArquivo = "Files/" & Funcoes.GeraNomeArquivo & ".XLS"
    '        End If

    '        Dim strArquivo As String = HttpContext.Current.Server.MapPath(strNomeArquivo)

    '        crptRelatorio.SetDataSource(ds)
    '        Funcoes.BindParameters(crptRelatorio, param)

    '        If Dir(strArquivo).Length > 0 Then Kill(strArquivo)

    '        If pdf Then
    '            crptRelatorio.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, strArquivo)
    '        Else
    '            crptRelatorio.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Excel, strArquivo)
    '        End If

    '        If IO.File.Exists(strArquivo) Then
    '            If pdf Then
    '                ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & strNomeArquivo & "');", True)
    '            Else
    '                ScriptManager.RegisterClientScriptBlock(Me, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.location = '" & strNomeArquivo & "';", True)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    Finally
    '        crptRelatorio.Close()
    '        crptRelatorio.Dispose()
    '    End Try
    'End Sub

    Private Sub LimparPedido(ByVal tipo As String)
        '*************************************************************************
        '************************  Remove Sessoes  *******************************
        '*************************************************************************
        Session.Remove("ssMessage")
        Session.Remove("objPedido" & HID.Value)
        Session.Remove("objPedVencimentos" & HID.Value)
        Session.Remove("objPedidoSelecionado" & HID.Value)
        Session.Remove("objClientePXI" & HID.Value)
        Session.Remove("objPracaPXI" & HID.Value)
        Session.Remove("objContaBancariaPXI" & HID.Value)
        Session.Remove("Fixacoes" & HID.Value)
        Session.Remove("FixacaoXNotaFiscal" & HID.Value)
        Session.Remove("objNavioXInvoice" & HID.Value)
        Session.Remove("ssNomeArquivoPedido" & HID.Value)

        '*************************************************************************
        '*********************  Recria pedido e HID  *****************************
        '*************************************************************************
        HID.Value = Guid.NewGuid().ToString
        objPedido = New [Lib].Negocio.Pedido()
        objPedido.IUD = "I"

        'Usuario de Inclusao / alteacao / delecao
        ddlUsuarios.Items.Clear()
        ddlUsuarios.Items.Add("Inc.- " & Session("ssNomeUsuario"))
        'Fiscal e financeiro Aberto ou Fechado
        MudarSituacao("FISCAL", False)
        MudarSituacao("FINANCEIRO", False)

        '*************************************************************************
        '************************* Barra de Botoes *******************************
        '*************************************************************************
        lnkNovoPedido.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkEnviarEmail.Parent.Visible = False
        lnkBloquear.Parent.Visible = False

        '*************************************************************************
        '**************** Primeira Coluna da pagina mestre ***********************
        '*************************************************************************
        txtRegistro.Text = ""
        txtRegistro.Enabled = True
        imgExtratoPedido.Visible = False
        txtPedido.Text = ""
        txtContrato.Text = ""
        txtXPedNFe.Text = ""
        txtItemXPedNFe.Text = ""

        txtDataEntregaInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtDataEntregaFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtDataPedido.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtDataVencimento.Text = DateTime.Now.ToString("dd/MM/yyyy")

        objPedido.DataEntregaInicial = DateTime.Now
        objPedido.DataEntregaFinal = DateTime.Now
        objPedido.DataPedido = DateTime.Now
        objPedido.DataVencimentoPedido = DateTime.Now

        chkPeso.Checked = False
        objPedido.OrigemDestino = "0"

        ddlComercializacao.SelectedValue = 1
        objPedido.Comercializacao = eComercializacao.Comercio

        '*************************************************************************
        '**************** Segunda Coluna da pagina mestre ************************
        '*************************************************************************
        cmbSituacao.SelectedValue = 1
        objPedido.CodigoSituacao = 1

        cmbEmbalagem.SelectedIndex = 0
        rbPagAnterior.Checked = False
        rbPagPosterior.Checked = False
        rbPtaxAnterior.Checked = False
        rbPtaxPosterior.Checked = False
        txtLocalDeEmbarque.Text = ""

        cmbFinalidade.SelectedIndex = 0
        cmbFinalidade.Enabled = True

        cmbSafra.SelectedIndex = 0
        cmbSafra.Enabled = True

        cmbMoeda.SelectedIndex = 0
        chkTemVariacao.Checked = False
        chkTemVariacao.Visible = False
        objPedido.CodigoMoeda = 0
        cmbMoeda.Enabled = True

        cmbIndexador.Visible = True
        txtIndiceFixado.Visible = True
        imgAjudaFixarDolar1.Visible = True

        cmbIndexador.Items.Clear()
        cmbIndexador.Enabled = True

        objPedido.IndiceFixado = Funcoes.PegarValorConversao(2, objPedido.DataPedido)
        objPedido.IndiceFixado = 0
        txtIndiceFixado.Text = objPedido.IndiceFixado.ToString("N8")
        txtIndiceFixado.Enabled = False
        chkIndexadorFixo.Checked = False
        chkIndexadorFixo.Enabled = True
        objPedido.IndexadorFixo = chkIndexadorFixo.Checked
        txtTaxa.Text = "0,00"


        '*************************************************************************
        '**************** Terceira Coluna da pagina mestre ***********************
        '*************************************************************************
        If String.IsNullOrWhiteSpace(tipo) Then
            cmbUnidadeNegocio.SelectedIndex = 0
            'Desabilitado Empresa - Furlan - 23/08/2022
            'cmbUnidadeNegocio.Enabled = True

            cmbEmpresa.Items.Clear()
            'Desabilitado Empresa - Furlan - 23/08/2022
            'cmbEmpresa.Enabled = True

            VerificaUnidade()
        Else
            objPedido.CodigoUnidadeNegocio = cmbUnidadeNegocio.SelectedValue
            Dim Emp As String() = cmbEmpresa.SelectedValue.ToString.Split("-")
            objPedido.CodigoEmpresa = Emp(0)
            objPedido.EnderecoEmpresa = Emp(1)
        End If

        objPedido.Empresa = New [Lib].Negocio.Cliente(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa)
        If objPedido.Empresa.Empresa.PedidoBloqueado Then
            btnBloqueado.Parent.Visible = True
            btnBloqueado.BackColor = Drawing.Color.Green
            btnBloqueado.Text = "Liberado"
        Else
            btnBloqueado.Parent.Visible = False
        End If

        If objPedido.Empresa.Empresa.FretePedido Then
            'pnlRoteiro.Visible = True
            TabTransportes.Visible = True
        Else
            pnlRoteiro.Visible = False
            TabTransportes.Visible = False
        End If

        If objPedido.Empresa.Empresa.ObrigaNavio Then
            divNaviosXInvoice.Visible = True
        Else
            divNaviosXInvoice.Visible = False
        End If
        txtNaviosXInvoice.Text = String.Empty

        'Desabilitado Empresa - Furlan - 23/08/2022
        'cmbUnidadeNegocio.Enabled = True
        'cmbEmpresa.Enabled = True

        cmbOperacao.SelectedIndex = 0
        cmbOperacao.Enabled = True

        cmbSubOperacao.Items.Clear()
        cmbSubOperacao.Enabled = True

        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        cmdConsultarCliente.Enabled = True

        txtPraca.Text = ""
        txtCodigoPraca.Value = ""
        btnConsultaPraca.Enabled = True

        txtDadosBancarios.Text = ""
        cmdConsultaDadosBancarios.Enabled = True

        lblCifFob.Text = "Frete:"
        rbCIF.Text = "Não (CIF)"
        rbFOB.Text = "Sim (FOB)"
        rbTER.Text = "Terceiro"
        rbNenhum.Text = "Nenhum"
        rbCIF.Checked = False
        rbFOB.Checked = False
        rbNenhum.Checked = False

        'OUTROS
        chkAgruparFinanceiro.Checked = False

        'Troca e Antecipada
        chkAntecipada.Checked = False
        chkRecompra.Checked = False
        chkTroca.Checked = False
        chkTroca.Enabled = True

        'Div de geração de contrato de compra e venda baseado no pedido.
        divContratoArquivo.Visible = False

        '*************************************************************************
        '*************************** Tab Produto *********************************
        '*************************************************************************
        lnkAdicionarItem.Parent.Parent.Visible = True
        gridItensPedido.DataSource = Nothing
        gridItensPedido.DataBind()

        gridItensPedidoDeposito.DataSource = Nothing
        gridItensPedidoDeposito.DataBind()

        gridItensPedidoAFixar.DataSource = Nothing
        gridItensPedidoAFixar.DataBind()

        gridEncargosGerais.DataSource = Nothing
        gridEncargosGerais.DataBind()

        '*************************************************************************
        '*************************** Tab Deposito ********************************
        '*************************************************************************
        GridDepositos.SelectedIndex = -1
        GridDepositos.DataSource = Nothing
        GridDepositos.DataBind()
        LimparDepositos()

        grdRoteiros.DataSource = Nothing
        grdRoteiros.DataBind()

        '*************************************************************************
        '************************* Tab Transportador *****************************
        '*************************************************************************
        gridTransportes.SelectedIndex = -1
        gridTransportes.DataSource = Nothing
        gridTransportes.DataBind()
        LimparTransportes()

        '*************************************************************************
        '******************* Tab Representante / Comissoes ***********************
        '*************************************************************************
        GridRepresentantes.SelectedIndex = -1
        GridRepresentantes.DataSource = Nothing
        GridRepresentantes.DataBind()
        LimparComissoes()

        gridTabelaDeComissao.DataSource = Nothing
        gridTabelaDeComissao.DataBind()

        '*************************************************************************
        '************************* Tab Vencimentos *******************************
        '*************************************************************************
        LimparCondicoes(True)
        btnOkVencimento.Enabled = True
        cmbCondicoes.Enabled = True
        If lstCondicoesPgtoEntrega.Items.Count > 0 Then
            lstCondicoesPgtoEntrega.SelectedIndex = 0
        End If
        txtQuotaDeEntrega.Text = "0"
        ddlPeriodicidadeEntrega.SelectedIndex = 0
        divEntrega.Visible = False

        ddlMomentoFinanceiro.SelectedValue = 0
        ddlMomentoFinanceiro.Enabled = True

        TabVencimentosOld.Visible = False
        TabParcelamento.Visible = False
        '*************************************************************************
        '************************* Tab Observacao ********************************
        '*************************************************************************
        txtAddObservacao.Text = ""
        txtObservacao.Text = ""

        '*************************************************************************
        '**************************** Tab Troca **********************************
        '*************************************************************************
        lblEmpresaTroca.Text = ""
        lblClienteTroca.Text = ""
        lblPedidoTroca.Text = ""
        lblCNTroca.Text = ""
        txtVendaTroca.Text = ""
        txtCompraTroca.Text = ""
        txtSaldoTroca.Text = ""
        ddlContaAdto.Items.Clear()
        LnkVincular.Parent.Visible = False
        LnkDesvincular.Parent.Visible = False
        TabPedidoTroca.Visible = False
        divContaAdiantamentoTroca.Visible = True


        '*************************************************************************
        '**************************** Tab Parcelas *******************************
        '*************************************************************************
        ddlCondPagPed.SelectedIndex = 0
        txtDataCondPagParcela.Text = ""
        txtPedidoTotal.Text = "0"
        txtPedidoTotalPago.Text = "0"
        txtPedidoSaldo.Text = "0"
        gridParcelas.DataSource = Nothing
        gridParcelas.DataBind()

        txtCodigoParcela.Text = ""
        txtDataVencParcela.Text = ""
        txtValorParcela.Text = ""

        GridOrigemFinanceiro.DataSource = Nothing
        GridOrigemFinanceiro.DataBind()

        gridFinanceiro.DataSource = Nothing
        gridFinanceiro.DataBind()

        txtAbertoCP.Text = "0,00"
        txtBaixadoCP.Text = "0,00"
        txtAdCP.Text = "0,00"
        txtBxAdCP.Text = "0,00"

        txtAbertoCR.Text = "0,00"
        txtBaixadoCR.Text = "0,00"
        txtAdCR.Text = "0,00"
        txtBxAdCR.Text = "0,00"

        txtResPedido.Text = "0,00"
        txtResRealizado.Text = "0,00"
        txtResSaldo.Text = "0,00"


        '*************************************************************************
        '**************************** Tab Contratos ******************************
        '*************************************************************************
        txtDescricaoContrato.Text = String.Empty
        txtNomeDoArquivo.Text = String.Empty
        gridContratos.DataSource = objPedido.Contratos
        gridContratos.DataBind()


        '*************************************************************************
        '***************** Tab Observações de Controle Interno *******************
        '*************************************************************************
        txtObservacaoControleInterno.Text = String.Empty
        TabControleInterno.Visible = False


        '*************************************************************************
        '*************************************************************************


        SessaoSalvaPedido()

        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaDadosBancarios.SetarHID(HID.Value)
        ucConsultaEstados.SetarHID(HID.Value)
        ucConsultaCodMunicipios.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)
        ucConsultaPedidoDeTroca.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        ucPedidoXLancamento.SetarHID(HID.Value)
        ucPedidoxEncargo.SetarHID(HID.Value)
        ucPedidoxFixacao.SetarHID(HID.Value)
        ucConsultaEncargosPlanoDeContas.SetarHID(HID.Value)
        ucConsultarNaviosXInvoice.SetarHID(HID.Value)
        ucEmailNFePedido.SetarHID(HID.Value)
        Session.Remove("_MainUserControl")

        tcPedido.ActiveTab = TabProduto
        txtRegistro.Focus()

    End Sub

    Protected Sub lnkNovoPedido_Click(sender As Object, e As EventArgs) Handles lnkNovoPedido.Click
        Try
            If Funcoes.VerificaPermissao("PedidosXItens", "GRAVAR") Then
                SessaoRecuperaPedido()

                If objPedido.FiscalAberto OrElse (Not objPedido.FiscalAberto AndAlso Funcoes.VerificaAcesso(objPedido.Empresa.Codigo, objPedido.Empresa.CodigoEndereco, IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text)), "PEDIDOS")) Then
                    If ValidarCamposPedido() Then
                        If String.IsNullOrWhiteSpace(txtRegistro.Text) Then
                            Try
                                Dim novoCodigo As Integer = 0

                                If FinanceiroVirtual AndAlso Not String.IsNullOrWhiteSpace(objPedido.Empresa.Empresa.Servidor) AndAlso UsuarioServidor.NomeServidor <> objPedido.Empresa.Empresa.Servidor Then
                                    novoCodigo = [Lib].Negocio.Numerador.PegarNumeroRemoto(2, objPedido.Empresa.Empresa.Servidor, [Lib].Negocio.eTiposNumerador.Pedido)
                                Else
                                    novoCodigo = [Lib].Negocio.Numerador.PegarNumero(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, [Lib].Negocio.eTiposNumerador.Pedido)
                                End If

                                If novoCodigo = -1 Then
                                    MsgBox(Me.Page, "Numerador não cadastrado!")
                                    Exit Sub
                                ElseIf novoCodigo = -2 Then
                                    MsgBox(Me.Page, [Lib].Negocio.Numerador.Erro.Message)
                                    Exit Sub
                                End If

                                If objPedido.IUD = "I" Then
                                    If Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "05272759" Then
                                        If objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "30101") Then
                                            If objPedido.Empresa.Empresa.PedidoBloqueado AndAlso
                                               (objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES) Then
                                                objPedido.PedidoBloqueado = True
                                            Else
                                                objPedido.PedidoBloqueado = False
                                            End If
                                        End If

                                    ElseIf Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "05366261" OrElse
                                        Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "38198213" OrElse
                                        Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "40938762" OrElse
                                        Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "49673784" Then

                                        If objPedido.Empresa.Empresa.PedidoBloqueado AndAlso (objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRAS OrElse
                                                                                              objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRASAORDEM OrElse
                                                                                              objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
                                                                                              objPedido.SubOperacao.Classe = eClassesOperacoes.VENDASAORDEM OrElse
                                                                                              objPedido.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES) Then
                                            objPedido.PedidoBloqueado = True
                                        Else
                                            objPedido.PedidoBloqueado = False
                                        End If

                                        'Acrescentado Libera Carregamento Pedido Osmar - Furlan - 11/04/2024
                                        If objPedido.Empresa.Empresa.LiberaCarregamento Then objPedido.LiberaCarregamento = False

                                    Else
                                        If Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "04854422" Then
                                            If objPedido.Empresa.Empresa.PedidoBloqueado AndAlso ([Enum].Parse(GetType(eClassesOperacoes), objPedido.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.VENDAS) Then
                                                objPedido.PedidoBloqueado = True
                                            Else
                                                objPedido.PedidoBloqueado = False
                                            End If

                                        ElseIf Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "44979506" Then 'VERDE

                                            If objPedido.Empresa.Empresa.PedidoBloqueado AndAlso
                                                 (objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRAS OrElse
                                                 objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
                                                 objPedido.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES OrElse
                                                 objPedido.SubOperacao.Classe = eClassesOperacoes.VENDASAORDEM) AndAlso
                                                 objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                                                objPedido.PedidoBloqueado = True
                                            Else
                                                objPedido.PedidoBloqueado = False
                                            End If

                                        ElseIf Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "24450490" Then 'RT GRÃOS
                                            If objPedido.Empresa.Empresa.PedidoBloqueado AndAlso
                                                 (objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRAS OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.IMPORTACOES OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.VENDASAORDEM) AndAlso
                                                (objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20101") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20102") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20103") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20104") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20105") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20106")) Then
                                                objPedido.PedidoBloqueado = True
                                            Else
                                                objPedido.PedidoBloqueado = False
                                            End If

                                        ElseIf Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "03961253" Then
                                            If objPedido.Empresa.Empresa.PedidoBloqueado AndAlso
                                                objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso
                                                 (objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.TRANSFERENCIAS OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.LUCROREAL OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.VENDASAORDEM OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.GLOBAL) AndAlso
                                                (objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10101") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10102") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10103") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10104") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10105") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10220") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10301") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10401") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20101") OrElse
                                                 objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "70101")) Then
                                                objPedido.PedidoBloqueado = True
                                            Else
                                                objPedido.PedidoBloqueado = False
                                            End If
                                        ElseIf objPedido.Itens.Exists(Function(s) s.Produto.Nome.Contains("SOJA") Or s.Produto.Nome.Contains("MILHO") Or s.Produto.Nome.Contains("LECITINA") And Not s.Produto.Nome.Contains("TAMBOR")) Then
                                            If objPedido.Empresa.Empresa.PedidoBloqueado AndAlso
                                               (objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRAS OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.TRANSFERENCIAS OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.LUCROREAL OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.VENDASAORDEM OrElse
                                                objPedido.SubOperacao.Classe = eClassesOperacoes.GLOBAL) Then
                                                objPedido.PedidoBloqueado = True
                                            Else
                                                objPedido.PedidoBloqueado = False
                                            End If
                                        End If
                                    End If
                                End If

                                objPedido.Codigo = novoCodigo
                                objPedido.UsuarioInclusao = Session("ssNomeUsuario")

                                If Not cmbEmbalagem.SelectedValue Is Nothing AndAlso cmbEmbalagem.SelectedValue.Length > 0 Then
                                    objPedido.Embalagem = cmbEmbalagem.SelectedValue
                                End If

                                'Pagamento anterior ou posterior ao carregamento
                                'AC - Pagamento anterior ao carregamento - (AC) - Anterior Carregamento
                                'PC - Pagamento posterior ao carregamento - (PC) - Posterior Carregamento
                                If rbPagAnterior.Checked Then
                                    objPedido.TipoCondicaoEntrega = "AC"
                                ElseIf rbPagPosterior.Checked Then
                                    objPedido.TipoCondicaoEntrega = "PC"
                                End If

                                'Ptax anterior ou posterior ao carregamento
                                'XA - Ptax anterior ao carregamento - (XA) - Ptax Anterior
                                'XP - Ptax posterior ao pagamento - (XP) - Ptax Posterior
                                If rbPtaxAnterior.Checked Then
                                    objPedido.TipoPagamentoPtax = "XP"
                                ElseIf rbPtaxPosterior.Checked Then
                                    objPedido.TipoPagamentoPtax = "XA"
                                End If

                                If FinanceiroVirtual AndAlso Not objPedido.Bloquear() Then
                                    MsgBox(Me.Page, "O pedido " & objPedido.Codigo & " foi bloqueado por outro usuário, por favor recarregue o registro!")
                                    Exit Sub
                                End If

                                If objPedido.Salvar() Then
                                    MsgBox(Me.Page, "O pedido " & novoCodigo & " foi incluído com Sucesso.", eTitulo.Sucess)
                                    'ImprimirPedido(True)

                                    objPedido = New Pedido(cmbEmpresa.SelectedValue.Split("-")(0), cmbEmpresa.SelectedValue.Split("-")(1), novoCodigo)

                                    objPedido.ImprimirPedido(Me.Page, True)

                                    LimparPedido("I")
                                Else
                                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                                End If
                            Catch ex As Exception
                                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                            End Try
                        Else
                            MsgBox(Me.Page, "Pedido já gravado!")
                        End If
                    End If
                Else
                    MsgBox(Me.Page, "Movimento " & (IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text))) & " já fechado para esta data!")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar pedido!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            SessaoRecuperaPedido()
            If Funcoes.VerificaPermissao("PedidosXItens", "ALTERAR") Then
                If ValidarCamposPedido() Then
                    Try
                        objPedido.UsuarioAlteracao = Session("ssNomeUsuario")
                        objPedido.DataAlteracao = DateTime.Now

                        If FinanceiroVirtual AndAlso Not objPedido.Bloquear() Then
                            MsgBox(Me.Page, "O pedido " & objPedido.Codigo & " foi bloqueado por outro usuário, por favor recarregue o registro!")
                            Exit Sub
                        End If

                        If Not cmbEmbalagem.SelectedValue Is Nothing AndAlso cmbEmbalagem.SelectedValue.Length > 0 Then
                            objPedido.Embalagem = cmbEmbalagem.SelectedValue
                        End If

                        objPedido.LocalDeEmbarque = txtLocalDeEmbarque.Text

                        'Pagamento anterior ou posterior ao carregamento
                        'AC - Pagamento anterior ao carregamento - (AC) - Anterior Carregamento
                        'PC - Pagamento posterior ao carregamento - (PC) - Posterior Carregamento
                        If rbPagAnterior.Checked Then
                            objPedido.TipoCondicaoEntrega = "AC"
                        ElseIf rbPagPosterior.Checked Then
                            objPedido.TipoCondicaoEntrega = "PC"
                        End If

                        'Ptax anterior ou posterior ao carregamento
                        'XA - Ptax anterior ao carregamento - (XA) - Ptax Anterior
                        'XP - Ptax posterior ao pagamento - (XP) - Ptax Posterior
                        If rbPtaxAnterior.Checked Then
                            objPedido.TipoPagamentoPtax = "XP"
                        ElseIf rbPtaxPosterior.Checked Then
                            objPedido.TipoPagamentoPtax = "XA"
                        End If

                        If objPedido.Salvar() Then
                            MsgBox(Me.Page, "O pedido " & objPedido.Codigo.ToString() & " foi atualizado com Sucesso.", eTitulo.Sucess)
                            'ImprimirPedido(True)

                            objPedido = New Pedido(cmbEmpresa.SelectedValue.Split("-")(0), cmbEmpresa.SelectedValue.Split("-")(1), txtRegistro.Text)

                            objPedido.ImprimirPedido(Me.Page, True)

                            LimparPedido("U")
                        Else
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        End If
                    Catch ex As Exception
                        MsgBox(Me.Page, ex.Message & ". " & ex.StackTrace)
                    End Try
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar pedido!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("PedidosXItens", "EXCLUIR") Then
                SessaoRecuperaPedido()

                If Not objPedido.FiscalAberto Then
                    MsgBox(Me.Page, "Pedido com Fiscal Fechado não pode ser excluído.")
                    Exit Sub
                ElseIf Not objPedido.FinanceiroAberto Then
                    MsgBox(Me.Page, "Pedido com Finaneiro Fechado não pode ser excluído.")
                    Exit Sub
                End If

                Dim strMensagem As String = ""

                If objPedido.Troca And objPedido.Operacao.CodigoClasse = eClassesOperacoes.VENDAS.ToString AndAlso Not objPedido.PedidoTroca Is Nothing Then
                    MsgBox(Me.Page, "O Pedido contem vinculo com outro pedido e não pode ser deletado.")
                    Exit Sub
                End If

                If Not objPedido.PedidoTroca Is Nothing AndAlso objPedido.PedidoTroca.Codigo > 0 Then
                    MsgBox(Me.Page, "Pedidos Vinculados não podem ser cancelados. Desvincule-o e então o cancele.")
                    Exit Sub
                End If

                If objPedido.SubOperacao.Financeiro AndAlso Not objPedido.Vencimentos Is Nothing AndAlso objPedido.Vencimentos.Count > 0 Then
                    For x As Integer = 0 To objPedido.Vencimentos.OrderBy(Function(s) s.Provisao).Count - 1
                        If Not String.IsNullOrWhiteSpace(objPedido.Vencimentos(x).ContratoBancario) Then
                            MsgBox(Me.Page, "Título(s) com Contrato Bancário. Pedido não pode ser excluido.")
                            Exit Sub
                        End If
                    Next

                    'Verificar se tem alguma baixa no financeiro deste pedido
                    If PedidoXParcelas.Existe(objPedido.Codigo, objPedido.SubOperacao.EntradaSaida, eProvisao.Baixa, objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa) Then
                        MsgBox(Me.Page, "O Pedido tem titulos Baixados.")
                        Exit Sub
                    End If
                End If

                'Verifica Pesagem
                If objPedido.TemPesagem Then
                    MsgBox(Me.Page, "Pedido com Pesagem não pode ser excluído.")
                    Exit Sub
                End If

                'Verifica Nota Fiscal
                For Each item As [Lib].Negocio.PedidoXItem In objPedido.Itens
                    If item.TemNota Then
                        MsgBox(Me.Page, "Produto " & item.CodigoProduto & "-" & item.Produto.Nome & " com Nota Fiscal, Pedido não pode ser excluído.")
                        Exit Sub
                    End If
                Next

                ''Verificar se tem alguma Cessão de Crédito com este pedido
                'If ListProcuracao.Existe(objPedido.Codigo) Then
                '    MsgBox(Me.Page, "O Pedido esta vinculado a Cessão de Crédito.")
                '    Exit Sub
                'End If

                Dim PedidoReferencia As Integer
                PedidoReferencia = objPedido.TemReferenciaEmOutroPedidoDeTroca
                If PedidoReferencia > 0 Then
                    MsgBox(Me.Page, "Existe outro Pedido (" & PedidoReferencia & ") usando este como Troca. Desmarque a opção troca do outro para possibilitar a exclusão deste.")
                    Exit Sub
                End If

                objPedido.IUD = "C"
                objPedido.UsuarioCancelamento = Session("ssNomeUsuario")
                objPedido.DataCancelamento = DateTime.Now

                If objPedido.Salvar Then
                    MsgBox(Me.Page, "O pedido " & objPedido.Codigo.ToString() & " foi excluido com Sucesso.", eTitulo.Sucess)
                    LimparPedido("D")
                Else
                    MsgBox(Me.Page, "Erro ao excluir o Pedido: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissão para excluir pedido!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("PedidosXItens", "LEITURA") Then
                ListarPedidos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar pedido!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparPedido("L")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            'SessaoRecuperaPedido()
            'ImprimirPedido(True)

            objPedido = New Pedido(cmbEmpresa.SelectedValue.Split("-")(0), cmbEmpresa.SelectedValue.Split("-")(1), txtRegistro.Text)

            objPedido.ImprimirPedido(Me.Page, True)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            'SessaoRecuperaPedido()
            'ImprimirPedido(False)

            objPedido = New Pedido(cmbEmpresa.SelectedValue.Split("-")(0), cmbEmpresa.SelectedValue.Split("-")(1), txtRegistro.Text)

            objPedido.ImprimirPedido(Me.Page, False)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkContrato_Click(sender As Object, e As EventArgs) Handles lnkContrato.Click
        Try
            Dim var As Dictionary(Of String, Object) = New Dictionary(Of String, Object)
            Dim ConteudoDaTabela As Dictionary(Of String, Object) = Nothing

            If objPedido Is Nothing Then SessaoRecuperaPedido()

            If objPedido.Itens.Count = 0 Then
                MsgBox(Me.Page, "Consulte um Pedido.")
                Exit Sub
            ElseIf Not objPedido Is Nothing AndAlso objPedido.Vencimentos Is Nothing AndAlso objPedido.Vencimentos.Count = 0 Then
                MsgBox(Me.Page, "Pedido sem Financeiro.")
                Exit Sub
            End If

            var.Add("#COD_PEDIDO#", objPedido.Codigo)
            '#Nutri - Pedido
            var.Add("#023#", objPedido.Codigo)

            'CLIENTE
            var.Add("#NOME_CLIENTE#", objPedido.Cliente.Nome)
            var.Add("#CPF_CLIENTE#", [Lib].Negocio.Funcoes.FormatarCpfCnpj(objPedido.CodigoCliente))
            var.Add("#IE_CLIENTE#", IIf(objPedido.Cliente.InscricaoEstadual.Length > 0, [Lib].Negocio.Funcoes.FormatarInscricaoEstadual(objPedido.Cliente.InscricaoEstadual), ""))

            Dim enderecoCliente As String = objPedido.Cliente.Endereco
            If objPedido.Cliente.Numero > 0 Then enderecoCliente += "," & objPedido.Cliente.Numero
            If objPedido.Cliente.Complemento.Length > 0 Then enderecoCliente += " - " & objPedido.Cliente.Complemento
            If objPedido.Cliente.Bairro.Length > 0 Then enderecoCliente += " - " & objPedido.Cliente.Bairro
            var.Add("#END_CLIENTE#", enderecoCliente)
            var.Add("#CIDADE_CLIENTE#", objPedido.Cliente.Cidade)
            var.Add("#UF_CLIENTE#", objPedido.Cliente.CodigoEstado)

            Dim cepCliente As String = String.Empty
            If objPedido.Cliente.CEP.Contains("-") Then
                cepCliente = objPedido.Cliente.CEP
            Else
                cepCliente = Left(objPedido.Cliente.CEP, 5) & "-" & Mid(objPedido.Cliente.CEP, 6, 3)
            End If
            var.Add("#CEP_CLIENTE#", cepCliente)

            '#Nutri - Cliente/Fornecedor
            var.Add("#009#", objPedido.Cliente.Nome)
            var.Add("#012#", [Lib].Negocio.Funcoes.FormatarCpfCnpj(objPedido.CodigoCliente))
            var.Add("#105#", IIf(objPedido.Cliente.InscricaoEstadual.Length > 0, [Lib].Negocio.Funcoes.FormatarInscricaoEstadual(objPedido.Cliente.InscricaoEstadual), ""))
            var.Add("#106#", IIf(objPedido.Cliente.InscricaoEstadual.Length > 0, [Lib].Negocio.Funcoes.FormatarInscricaoEstadual(objPedido.Cliente.InscricaoEstadual), ""))

            var.Add("#104#", objPedido.Cliente.Endereco)

            Dim sNumeroCliente As String = ""
            If objPedido.Cliente.Numero > 0 Then sNumeroCliente += "," & objPedido.Cliente.Numero
            var.Add("#107#", sNumeroCliente)

            var.Add("#014#", [Lib].Negocio.Funcoes.FormatarTelefone(objPedido.Cliente.Telefone))
            var.Add("#109#", objPedido.Cliente.Cidade)
            var.Add("#110#", objPedido.Cliente.CodigoEstado)

            var.Add("#115#", objPedido.Cliente.Cidade)
            var.Add("#116#", objPedido.Cliente.CodigoEstado)

            If String.IsNullOrWhiteSpace(objPedido.Cliente.EmailNFE) Then
                var.Add("#154#", objPedido.Cliente.Email)
            Else
                var.Add("#154#", objPedido.Cliente.EmailNFE)
            End If

            '#verde - dados
            If objPedido.Cliente.Sexo = "M" Then
                var.Add("#NACIONALIDADE#", "BRASILEIRO")
            Else
                var.Add("#NACIONALIDADE#", "BRASILEIRA")
            End If

            If Not String.IsNullOrWhiteSpace(objPedido.Cliente.EstadoCivil) Then
                Dim estadoCivel As String = objPedido.Cliente.EstadoCivil
                estadoCivel = estadoCivel.Substring(0, estadoCivel.Length - 1)

                If objPedido.Cliente.Sexo = "M" Then
                    var.Add("#CIVIL#", estadoCivel.ToUpper & "O")
                Else
                    var.Add("#CIVIL#", estadoCivel.ToUpper & "A")
                End If
            End If

            var.Add("#RG#", objPedido.Cliente.RG)

            If objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS Then

                If Not objPedido.Empresa.ContasBancarias Is Nothing AndAlso objPedido.Empresa.ContasBancarias.Where(Function(x) x.Ativo).Count > 0 Then
                    Dim contaBancaria As New ClienteXContaBancaria
                    contaBancaria = objPedido.Empresa.ContasBancarias.FirstOrDefault()

                    If contaBancaria.TipoConta = "C" Then
                        var.Add("#CONTA#", "CORRENTE")
                    ElseIf contaBancaria.TipoConta = "P" Then
                        var.Add("#CONTA#", "POUPANÇA")
                    Else
                        var.Add("#CONTA#", "")
                    End If
                Else
                    var.Add("#CONTA#", "")
                End If
            Else
                If Not objPedido.ContaBancariaSelecionada Is Nothing AndAlso objPedido.ContaBancariaSelecionada.TipoConta = "C" Then
                    var.Add("#CONTA#", "CORRENTE")
                ElseIf Not objPedido.ContaBancariaSelecionada Is Nothing AndAlso objPedido.ContaBancariaSelecionada.TipoConta = "P" Then
                    var.Add("#CONTA#", "POUPANÇA")
                Else
                    var.Add("#CONTA#", "")
                End If
            End If

            var.Add("#PROFISSAO#", objPedido.Cliente.Categoria.Descricao)

            'EMPRESA
            var.Add("#NOME_EMPRESA#", objPedido.Empresa.Nome)
            var.Add("#CNPJ_EMPRESA#", [Lib].Negocio.Funcoes.FormatarCpfCnpj(objPedido.CodigoEmpresa))
            var.Add("#IE_EMPRESA#", ([Lib].Negocio.Funcoes.FormatarInscricaoEstadual(objPedido.Empresa.InscricaoEstadual)))

            Dim enderecoEmpresa As String = objPedido.Empresa.Endereco
            If objPedido.Empresa.Numero > 0 Then enderecoEmpresa += "," & objPedido.Empresa.Numero
            If objPedido.Empresa.Complemento.Length > 0 Then enderecoEmpresa += " - " & objPedido.Empresa.Complemento
            If objPedido.Empresa.Bairro.Length > 0 Then enderecoEmpresa += " - " & objPedido.Empresa.Bairro
            var.Add("#END_EMPRESA#", enderecoEmpresa)
            var.Add("#CIDADE_EMPRESA#", objPedido.Empresa.Cidade)
            var.Add("#UF_EMPRESA#", objPedido.Empresa.CodigoEstado)

            Dim cepEmpresa As String = String.Empty
            If objPedido.Empresa.CEP.Contains("-") Then
                cepEmpresa = objPedido.Empresa.CEP
            Else
                cepEmpresa = Left(objPedido.Empresa.CEP, 5) & "-" & Mid(objPedido.Empresa.CEP, 6, 3)
            End If
            var.Add("#CEP_EMPRESA#", cepEmpresa)

            '#Nutri - Empresa
            var.Add("#001#", objPedido.Empresa.Nome)
            var.Add("#004#", [Lib].Negocio.Funcoes.FormatarCpfCnpj(objPedido.CodigoEmpresa))
            var.Add("#005#", ([Lib].Negocio.Funcoes.FormatarInscricaoEstadual(objPedido.Empresa.InscricaoEstadual)))
            var.Add("#002#", objPedido.Empresa.Endereco)

            Dim sNumeroEmpresa As String = ""
            If objPedido.Empresa.Numero > 0 Then sNumeroEmpresa += "," & objPedido.Empresa.Numero
            var.Add("#007#", sNumeroEmpresa)
            var.Add("#006#", [Lib].Negocio.Funcoes.FormatarTelefone(objPedido.Empresa.Telefone))
            var.Add("#003#", objPedido.Empresa.Cidade)
            var.Add("#094#", objPedido.Empresa.CodigoEstado)

            '#verde - dados
            '44979506-10101.docx
            If String.IsNullOrWhiteSpace(objPedido.Empresa.EmailNFE) Then
                var.Add("#EMAIL#", objPedido.Empresa.Email)
            Else
                var.Add("#EMAIL#", objPedido.Empresa.EmailNFE)
            End If

            var.Add("#ESPECIAL#", objPedido.Observacoes)

            If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                var.Add("#DESIGNACAO_CLIENTE#", "COMPRADOR")
                var.Add("#DESIGNACAO_EMPRESA#", "VENDEDORA")
            Else
                var.Add("#DESIGNACAO_CLIENTE#", "VENDEDOR")
                var.Add("#DESIGNACAO_EMPRESA#", "COMPRADORA")
            End If

            If objPedido.Vencimentos IsNot Nothing AndAlso objPedido.Vencimentos.Count > 0 Then
                var.Add("#VENCIMENTO#", objPedido.Vencimentos(0).DataProrrogacao.ToString("dd/MM/yyyy"))


                If objPedido.CondicaoPagamento.Parcelas = 1 AndAlso objPedido.CondicaoPagamento.Periodo(0) > 1 Then
                    var.Add("#074#", "Pedido será pago em " & objPedido.CondicaoPagamento.Periodo(0) & " dias após cada Faturamento")
                Else

                    If (Left(objPedido.CodigoEmpresa, 8) = "05366261" OrElse Left(objPedido.CodigoEmpresa, 8) = "44979506") And objPedido.CondicaoPagamento.Parcelas > 1 Then

                        Dim sParcelas As String = String.Format("Pedido será pago em {0} parcelas: {1}", objPedido.CondicaoPagamento.Parcelas, vbCrLf)

                        Dim nParcela As Integer = 1
                        For Each vencimento In objPedido.Vencimentos
                            sParcelas += String.Format("Parcela {0} será pago em {1} no valor de R$ {2} ({3})", nParcela, vencimento.DataProrrogacao.ToString("dd/MM/yyyy"), vencimento.ValorLiquidoOficial.ToString("N2"), Funcoes.Extenso(vencimento.ValorLiquidoOficial, "real", "reais"))
                            If nParcela < objPedido.CondicaoPagamento.Parcelas Then
                                sParcelas += ", "
                            End If
                            nParcela += 1
                        Next
                        var.Add("#074#", sParcelas)

                    Else

                        var.Add("#074#", "Pedido será pago em " & objPedido.Vencimentos(0).DataProrrogacao.ToString("dd/MM/yyyy") & " no valor de R$ " & objPedido.Itens.LiquidoOficial.ToString("N2") & " (" & Funcoes.Extenso(objPedido.Itens.LiquidoOficial, "real", "reais") & ")")

                    End If

                    'objPedido.Vencimentos(0).ValorDocumentoOficial

                End If

                var.Add("#076#", objPedido.CondicaoPagamento.Descricao)

            End If

            var.Add("#SAFRA#", objPedido.CodigoSafra)

            If objPedido.FreteCIFFOB = eTiposFrete.CIF Then
                var.Add("#E_CIF#", "X")
                var.Add("#E_FOB#", " ")
                var.Add("#058#", "CIF")  '#Nutri - Modalidade de Frete
            ElseIf objPedido.FreteCIFFOB = eTiposFrete.FOB Then
                var.Add("#E_CIF#", " ")
                var.Add("#E_FOB#", "X")
                var.Add("#058#", "FOB")
            Else
                var.Add("#E_FOB#", " ")
                var.Add("#E_CIF#", " ")
                var.Add("#058#", "")
            End If

            If txtLocalDeEmbarque.Text.Length > 0 Then
                var.Add("#LOCALENTREGA#", txtLocalDeEmbarque.Text.ToUpper)
            Else
                If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                    var.Add("#LOCALENTREGA#", objPedido.Empresa.Endereco & "," & objPedido.Empresa.Numero)
                Else
                    var.Add("#LOCALENTREGA#", objPedido.Cliente.Endereco & "," & objPedido.Cliente.Numero)
                End If
            End If

            var.Add("#MOEDA#", objPedido.Moeda.Simbolo)

            If objPedido.Vencimentos.Count > 0 Then
                var.Add("#VALOR#", objPedido.Vencimentos(0).ValorLiquidoOficial.ToString("N2"))
                var.Add("#VALOR_POR_EXTENSO#", [Lib].Negocio.Funcoes.Extenso(objPedido.Vencimentos(0).ValorLiquidoOficial, IIf(objPedido.Moeda.Classificacao = eTiposMoeda.Oficial, "Real", "Dólar"), IIf(objPedido.Moeda.Classificacao = eTiposMoeda.Oficial, "Reais", "Dólares")))
            End If

            var.Add("#DIA#", Day(Today))
            var.Add("#MES_POR_EXTENSO#", [Lib].Negocio.Funcoes.MesPorExtenso(Month(Today)))
            var.Add("#ANO#", Year(Today))

            Dim objUsuario As [Lib].Negocio.Usuario = New [Lib].Negocio.Usuario(IIf(String.IsNullOrWhiteSpace(objPedido.UsuarioAlteracao), objPedido.UsuarioInclusao, objPedido.UsuarioAlteracao))

            If Not objUsuario Is Nothing AndAlso String.IsNullOrWhiteSpace(objUsuario.CodigoCliente) Then
                MsgBox(Me.Page, "Vincule o usuário de alteração ou de inclusão Do pedido a um cliente na tela de usuários.")
            Else

                var.Add("#NOME_SEGUNDA_TESTEMUNHA#", objUsuario.Cliente.Nome)
                var.Add("#CPF_SEGUNDA_TESTEMUNHA#", [Lib].Negocio.Funcoes.FormatarCpfCnpj(objUsuario.CodigoCliente))

                var.Add("#PRODUTO#", objPedido.Itens(0).Descricao)
                var.Add("#QTDE_PRODUTO#", objPedido.Itens(0).Lancamentos(0).QuantidadeFaturamento)
                Dim extenso As String = UCase(Funcoes.Extenso(objPedido.Itens(0).Lancamentos(0).QuantidadeFaturamento, "", ""))
                var.Add("#QTDE_EXTENSO#", extenso)

                Dim unitario As Decimal = Math.Round((objPedido.Itens(0).Lancamentos(0).UnitarioFaturamento * 60), 2, MidpointRounding.AwayFromZero)
                var.Add("#UNI_PRODUTO#", unitario)
                Dim extensoUnitario As String = UCase(Funcoes.Extenso(unitario, "REAL", "REAIS"))
                var.Add("#UNI_EXTENSO#", extensoUnitario)

                var.Add("#PRAZO_PGTO#", objPedido.CondicaoPagamento.Descricao)

                Dim objDeposito As [Lib].Negocio.PedidoXDeposito = objPedido.Depositos.Where(Function(s) s.Tipo = "OD").FirstOrDefault

                Dim objDep As Cliente = New Cliente(objDeposito.Codigo, objDeposito.CodigoEndereco)
                Dim sDeposito As String = objDep.Endereco & IIf(objDep.Numero > 0, "," & objDep.Numero, "") & " - " & objDep.Cidade & "/" & objDep.CodigoEstado

                var.Add("#LOCAL_ENTREGA#", sDeposito)

                '#Nutri - Produto
                var.Add("#041#", objPedido.Itens(0).Descricao)
                var.Add("#042#", objPedido.Itens(0).Descricao)
                var.Add("#020#", objPedido.CodigoSafra)
                var.Add("#048#", objPedido.Itens(0).Lancamentos(0).QuantidadeFaturamento.ToString("N0"))
                var.Add("#119#", [Lib].Negocio.Funcoes.Extenso(objPedido.Itens(0).Lancamentos(0).QuantidadeFaturamento, "Quilo", "Quilos"))

                If objPedido.Itens(0).Produto.CodigoGrupo = "10101" OrElse objPedido.Itens(0).Produto.CodigoGrupo = "10102" OrElse objPedido.Itens(0).Produto.CodigoGrupo = "10103" OrElse objPedido.Itens(0).Produto.CodigoGrupo = "10201" OrElse objPedido.Itens(0).Produto.CodigoGrupo = "10401" OrElse objPedido.Itens(0).Produto.CodigoGrupo = "20101" OrElse objPedido.Itens(0).Produto.CodigoGrupo = "30101" Then
                    var.Add("#149#", Math.Round((objPedido.Itens(0).Lancamentos(0).QuantidadeFaturamento / 60), 0, MidpointRounding.AwayFromZero).ToString("N0"))

                    If Left(objPedido.CodigoEmpresa, 8) = "05366261" OrElse Left(objPedido.CodigoEmpresa, 8) = "49673784" OrElse Left(objPedido.CodigoEmpresa, 8) = "40938762" Then
                        var.Add("#150#", [Lib].Negocio.Funcoes.ExtensoGeneroFeminino(Math.Round((objPedido.Itens(0).Lancamentos(0).QuantidadeFaturamento / 60), 0, MidpointRounding.AwayFromZero), "saca", "sacas"))
                    Else
                        var.Add("#150#", [Lib].Negocio.Funcoes.Extenso(Math.Round((objPedido.Itens(0).Lancamentos(0).QuantidadeFaturamento / 60), 0, MidpointRounding.AwayFromZero), "saca", "sacas"))
                    End If

                    var.Add("#153#", "60")
                    Dim vUnitario As Decimal = Math.Round((objPedido.Itens(0).Lancamentos(0).UnitarioOficial * 60), 2, MidpointRounding.AwayFromZero)

                    If Left(objPedido.CodigoEmpresa, 8) = "44979506" AndAlso objPedido.Itens(0).Produto.CodigoGrupo = "10201" Then
                        vUnitario = objPedido.Itens(0).Lancamentos(0).UnitarioOficial
                    End If

                    var.Add("#151#", vUnitario.ToString("N2"))
                    var.Add("#152#", [Lib].Negocio.Funcoes.Extenso(vUnitario, "real", "reais"))
                Else
                    var.Add("#149#", "")
                    var.Add("#150#", "")
                    var.Add("#151#", "")
                    var.Add("#152#", "")
                    var.Add("#153#", "")
                End If


                var.Add("#050#", objPedido.Itens.TotalOficial.ToString("N2"))
                var.Add("#121#", [Lib].Negocio.Funcoes.Extenso((objPedido.Itens.TotalOficial), "real", "reais"))

                Dim vSenar As Decimal = 0
                Dim VFunrural As Decimal = 0

                For Each item In objPedido.Itens.Where(Function(x) x.IUD <> "D")
                    vSenar += item.Encargos.Where(Function(s) s.CodigoEncargo = "SENAR").Sum(Function(t) t.Valor)
                    VFunrural += item.Encargos.Where(Function(s) s.CodigoEncargo = "FUNRURAL").Sum(Function(t) t.Valor)
                Next

                var.Add("#135#", vSenar)
                var.Add("#136#", VFunrural)

                var.Add("#143#", objPedido.Itens.LiquidoOficial.ToString("N2"))
                var.Add("#144#", [Lib].Negocio.Funcoes.Extenso((objPedido.Itens.LiquidoOficial), "reais", "reais"))

                If Left(objPedido.CodigoEmpresa, 8) = "05366261" Or Left(objPedido.CodigoEmpresa, 8) = "40938762" Or Left(objPedido.CodigoEmpresa, 8) = "44979506" Or Left(objPedido.CodigoEmpresa, 8) = "27153202" Then
                    If (objPedido.Itens.TotalOficial - objPedido.Itens.LiquidoOficial) = 0 Then
                        var.Add("#145#", "")
                    Else
                        var.Add("#145#", String.Format("R$ {0}, a ser corrigido pela UPF/PR vigente, conforme item 5.1;", (objPedido.Itens.TotalOficial - objPedido.Itens.LiquidoOficial).ToString("N2")))
                    End If
                Else

                    If (objPedido.Itens.TotalOficial - objPedido.Itens.LiquidoOficial) = 0 Then
                        var.Add("#145#", "0,00")
                    Else
                        var.Add("#145#", (objPedido.Itens.TotalOficial - objPedido.Itens.LiquidoOficial).ToString("N2"))
                    End If

                End If

                If Left(objPedido.CodigoEmpresa, 8) = "05366261" Or Left(objPedido.CodigoEmpresa, 8) = "40938762" Or Left(objPedido.CodigoEmpresa, 8) = "44979506" Or Left(objPedido.CodigoEmpresa, 8) = "27153202" Then
                    var.Add("#155#", txtLocalDeEmbarque.Text)
                Else
                    var.Add("#155#", "")
                End If

                var.Add("#026#", objPedido.DataEntregaInicial.ToString("dd/MM/yyyy"))
                var.Add("#027#", objPedido.DataEntregaFinal.ToString("dd/MM/yyyy"))

                If objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS Then

                    If Not objPedido.Empresa.ContasBancarias Is Nothing AndAlso objPedido.Empresa.ContasBancarias.Where(Function(x) x.Ativo).Count > 0 Then
                        Dim contaBancaria As New ClienteXContaBancaria
                        contaBancaria = objPedido.Empresa.ContasBancarias.FirstOrDefault()
                        var.Add("#063#", contaBancaria.CodigoBanco)
                        var.Add("#064#", contaBancaria.Banco.Descricao)
                        var.Add("#090#", contaBancaria.CodigoAgencia.ToString().PadLeft(4, "0"c) & IIf(contaBancaria.DigitoAgencia.Length > 0, "-" & contaBancaria.DigitoAgencia, ""))
                        var.Add("#091#", contaBancaria.ContaCorrente & IIf(contaBancaria.DigitoConta.Length > 0, "-" & contaBancaria.DigitoConta, ""))
                    Else
                        var.Add("#063#", "")
                        var.Add("#064#", "")
                        var.Add("#090#", "")
                        var.Add("#091#", "")
                    End If

                Else

                    If Not objPedido.ContaBancariaSelecionada Is Nothing AndAlso objPedido.ContaBancariaSelecionada.CodigoBanco > 0 Then
                        var.Add("#063#", objPedido.ContaBancariaSelecionada.CodigoBanco)
                        var.Add("#064#", objPedido.ContaBancariaSelecionada.Banco.Descricao)
                        var.Add("#090#", objPedido.ContaBancariaSelecionada.CodigoAgencia & IIf(objPedido.ContaBancariaSelecionada.DigitoAgencia.Length > 0, "-" & objPedido.ContaBancariaSelecionada.DigitoAgencia, ""))
                        var.Add("#091#", objPedido.ContaBancariaSelecionada.ContaCorrente & IIf(objPedido.ContaBancariaSelecionada.DigitoConta.Length > 0, "-" & objPedido.ContaBancariaSelecionada.DigitoConta, ""))
                    Else
                        var.Add("#063#", "")
                        var.Add("#064#", "")
                        var.Add("#090#", "")
                        var.Add("#091#", "")
                    End If

                End If

                Dim ds As DataSet = New DataSet()

                Dim dtMercadoria As DataTable = ds.Tables.Add("Mercadoria")
                dtMercadoria.Columns.Add("id", GetType(Integer)).Unique = True
                dtMercadoria.Columns.Add("Descricao", GetType(String))
                dtMercadoria.Columns.Add("Unidade", GetType(String))
                dtMercadoria.Columns.Add("Valor", GetType(Decimal))

                Dim dtrow As DataRow
                For i As Integer = 1 To objPedido.Itens.Count
                    dtrow = dtMercadoria.NewRow()
                    dtrow("id") = i
                    dtrow("Descricao") = objPedido.Itens(i - 1).Descricao
                    dtrow("Unidade") = objPedido.Itens(i - 1).Produto.Unidade
                    dtrow("Valor") = objPedido.Itens(i - 1).Lancamentos.TotalOficialPrd.ToString("N2")
                    dtMercadoria.Rows.Add(dtrow)
                Next

                var.Add("#TABELA_DE_MERCADORIAS#", ds)
                Dim CaminhoDoArquivoGerado As String = String.Empty

                Dim arquivo As String = "ModeloDeContrato"

                If objPedido.Itens(0).Produto.CodigoGrupo = "10101" Then
                    arquivo = Left(objPedido.CodigoEmpresa, 8) & "-10101"
                ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10102" Then
                    arquivo = Left(objPedido.CodigoEmpresa, 8) & "-10102"
                End If

                'Baxi e Nutri usam o mesmo modelo de contrato
                If Left(objPedido.CodigoEmpresa, 8) = "05366261" Then
                    '(objPedido.Itens(0).Produto.CodigoGrupo = "10101" OrElse objPedido.Itens(0).Produto.CodigoGrupo = "10102") AndAlso
                    '(objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRAS OrElse
                    '    objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRASAORDEM OrElse
                    '    objPedido.SubOperacao.Classe = eClassesOperacoes.IMPORTACOES OrElse
                    '    objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
                    '    objPedido.SubOperacao.Classe = eClassesOperacoes.VENDASAORDEM OrElse
                    '    objPedido.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES OrElse
                    '    objPedido.SubOperacao.Classe = eClassesOperacoes.LUCROREAL) Then

                    If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then

                        If objPedido.Itens(0).Produto.CodigoGrupo = "10401" Then
                            arquivo = "05366261 10401 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10101" Then
                            arquivo = "05366261 10101 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10102" Then
                            arquivo = "05366261 10102 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10201" Then
                            arquivo = "05366261 10201"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "20101" Then
                            arquivo = "05366261 20101 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "30101" Then
                            arquivo = "05366261 30101 Compra"
                        Else
                            arquivo = "contrato compra nutri"
                        End If

                    Else

                        If objPedido.Itens(0).Produto.CodigoGrupo = "10401" Then
                            arquivo = "05366261 10401 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10101" Then
                            arquivo = "05366261 10101 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10102" Then
                            arquivo = "05366261 10102 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10201" Then
                            arquivo = "05366261 10201"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "20101" Then
                            arquivo = "05366261 20101 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "30101" Then
                            arquivo = "05366261 30101 Venda"
                        Else
                            arquivo = "contrato venda nutri"
                        End If

                    End If

                ElseIf Left(objPedido.CodigoEmpresa, 8) = "40938762" Then


                    If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then

                        If objPedido.Itens(0).Produto.CodigoGrupo = "10401" Then
                            arquivo = "40938762 10401 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10101" Then
                            arquivo = "40938762 10101 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10102" Then
                            arquivo = "40938762 10102 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10201" Then
                            arquivo = "40938762 10201"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "20101" Then
                            arquivo = "40938762 20101 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "30101" Then
                            arquivo = "40938762 30101 Compra"
                        Else
                            arquivo = "contrato compra baxi foods"
                        End If

                    Else

                        If objPedido.Itens(0).Produto.CodigoGrupo = "10401" Then
                            arquivo = "40938762 10401 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10101" Then
                            arquivo = "40938762 10101 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10102" Then
                            arquivo = "40938762 10102 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10201" Then
                            arquivo = "40938762 10201"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "20101" Then
                            arquivo = "40938762 20101 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "30101" Then
                            arquivo = "40938762 30101 Venda"
                        Else
                            arquivo = "contrato venda baxi foods"
                        End If

                    End If

                ElseIf Left(objPedido.CodigoEmpresa, 8) = "49673784" Then

                    If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then

                        If objPedido.Itens(0).Produto.CodigoGrupo = "10401" Then
                            arquivo = "49673784 10401 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10101" Then
                            arquivo = "49673784 10101 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10102" Then
                            arquivo = "49673784 10102 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10201" Then
                            arquivo = "49673784 10201"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "20101" Then
                            arquivo = "49673784 20101 Compra"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "30101" Then
                            arquivo = "49673784 30101 Compra"
                        Else
                            arquivo = "contrato compra baxi distribuidora"
                        End If

                    Else

                        If objPedido.Itens(0).Produto.CodigoGrupo = "10401" Then
                            arquivo = "49673784 10401 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10101" Then
                            arquivo = "49673784 10101 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10102" Then
                            arquivo = "49673784 10102 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10201" Then
                            arquivo = "49673784 10201"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "20101" Then
                            arquivo = "49673784 20101 Venda"
                        ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "30101" Then
                            arquivo = "49673784 30101 Venda"
                        Else
                            arquivo = "contrato venda baxi distribuidora"
                        End If

                    End If

                End If

                If (Left(objPedido.CodigoEmpresa, 8) = "44979506" OrElse Left(objPedido.CodigoEmpresa, 8) = "27153202") AndAlso
                        (objPedido.Itens(0).Produto.CodigoGrupo = "10101" OrElse
                         objPedido.Itens(0).Produto.CodigoGrupo = "10102" OrElse
                         objPedido.Itens(0).Produto.CodigoGrupo = "10103" OrElse
                        objPedido.Itens(0).Produto.CodigoGrupo = "10201") AndAlso
                        (objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRAS OrElse
                            objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRASAORDEM OrElse
                            objPedido.SubOperacao.Classe = eClassesOperacoes.IMPORTACOES OrElse
                            objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
                            objPedido.SubOperacao.Classe = eClassesOperacoes.VENDASAORDEM OrElse
                            objPedido.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES) Then

                    'REPRESENTANTE DO CLIENTE/FORNECEDOR
                    var.Add("#nomeRepre1#", "")
                    var.Add("#cpfRepre1#", "")

                    'REPRESENTANTE DA EMPRESA
                    If Left(objPedido.CodigoEmpresa, 8) = "44979506" Then
                        var.Add("#nomeRepre2#", "João Paulo Barbieri")
                        var.Add("#cpfRepre2#", "061.219.909-61")
                    Else
                        var.Add("#nomeRepre2#", "Neri Galvan")
                        var.Add("#cpfRepre2#", "123.456.789-00")
                    End If

                    var.Add("#UMIDADE#", "14,00")
                    var.Add("#IMPUREZA#", "1,00")

                    If objPedido.Itens(0).Produto.CodigoGrupo = "10101" Then
                        var.Add("#AVARIADO#", "8,00")

                        If Len(objPedido.Cliente.Codigo) = 11 Then
                            arquivo = "44979506-PF-10101"
                        Else
                            arquivo = "44979506-PJ-10101"
                        End If
                    ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10102" Then
                        var.Add("#AVARIADO#", "5,00")

                        If Len(objPedido.Cliente.Codigo) = 11 Then
                            arquivo = "44979506-PF-10102"
                        Else
                            arquivo = "44979506-PJ-10102"
                        End If
                    ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10103" Then
                        var.Add("#AVARIADO#", "78,00")

                        If Len(objPedido.Cliente.Codigo) = 11 Then
                            arquivo = "44979506-PF-10103"
                        Else
                            arquivo = "44979506-PJ-10103"
                        End If
                    ElseIf objPedido.Itens(0).Produto.CodigoGrupo = "10201" Then
                        If Len(objPedido.Cliente.Codigo) = 11 Then
                            arquivo = "44979506-PF-10201"
                        Else
                            arquivo = "44979506-PJ-10201"
                        End If
                    End If
                End If

                '#Nutri - Data
                var.Add("#126#", Now.ToString("dd/MM/yyyy"))

                'Funcoes.BindWordOffice(Me.Page, "ModeloDeContrato", var, CaminhoDoArquivoGerado)
                Funcoes.BindWordOffice(Me.Page, arquivo, var, CaminhoDoArquivoGerado)

                txtArquivoDeSaida.Text = CaminhoDoArquivoGerado
                divContratoArquivo.Visible = Not String.IsNullOrWhiteSpace(CaminhoDoArquivoGerado)

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEnviarEmail_Click(sender As Object, e As EventArgs) Handles lnkEnviarEmail.Click
        Try
            SessaoRecuperaPedido()
            If Not String.IsNullOrWhiteSpace(objPedido.Cliente.EmailNFE) Then
                ucEmailNFePedido.Limpar()

                Dim parameters As New Dictionary(Of String, Object)
                parameters("Empresa") = objPedido.CodigoEmpresa
                parameters("EndEmpresa") = objPedido.EnderecoEmpresa
                parameters("Pedido") = objPedido.Codigo

                parameters("EmailNFEPedido") = objPedido.Cliente.EmailNFE
                ucEmailNFePedido.EmailNFEPedido(parameters)

                Dim txtDestinatario As TextBox = CType(ucEmailNFePedido.FindControlRecursive("txtDestinatario"), TextBox)
                Popup.ConsultaDeEmailNFePedido(Me.Page, "objPedidosXItens" & HID.Value, txtDestinatario.ClientID, 100)
            Else
                MsgBox(Me.Page, "Cliente sem e-mail cadastrado. Por favor cadeastre no E-mail NFe!")
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imdDownload_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imdDownload.Click
        Try
            'Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" '"text/plain"
            Response.ContentType = "application/msword"
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & txtArquivoDeSaida.Text)
            Const bufferLength As Integer = 10000
            Dim buffer As Byte() = New Byte(bufferLength - 1) {}
            Dim length As Integer = 0
            Dim download As Stream = Nothing
            Try
                download = New FileStream(txtArquivoDeSaida.Text, FileMode.Open, FileAccess.Read)
                Do
                    If Response.IsClientConnected Then
                        length = download.Read(buffer, 0, bufferLength)
                        Response.OutputStream.Write(buffer, 0, length)
                        buffer = New Byte(bufferLength - 1) {}
                    Else
                        length = -1
                    End If
                Loop While length > 0
                Response.Flush()
                Response.End()
            Finally
                If download IsNot Nothing Then
                    download.Close()
                End If
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnFiscal_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaPedido()

            If objPedido.FiscalAberto Then

                Dim fecharPedido As Boolean = True

                If objPedido.SubOperacao.EstoqueFisico Then
                    If objPedido.SubOperacao.Classe = eClassesOperacoes.GLOBAL Then
                        For Each item In objPedido.SaldoItensPedido
                            If Not item.QtdeContratadoFisico = item.QtdeEntregueFisicoGlobal Then
                                fecharPedido = False
                            End If
                        Next
                    End If

                    If objPedido.SubOperacao.Classe = eClassesOperacoes.REMESSAS Then
                        For Each item In objPedido.SaldoItensPedido
                            If Not item.QtdeContratadoFisico = item.QtdeEntregueFisicoRemessa Then
                                fecharPedido = False
                            End If
                        Next
                    End If

                    If (objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRAS OrElse
                       objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRASAORDEM OrElse
                       objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
                       objPedido.SubOperacao.Classe = eClassesOperacoes.VENDASAORDEM OrElse
                       objPedido.SubOperacao.Classe = eClassesOperacoes.CONTAEORDEM OrElse
                       objPedido.SubOperacao.Classe = eClassesOperacoes.LUCROREAL) Then

                        For Each item In objPedido.SaldoItensPedido
                            If Not item.QtdeContratadoFisico = item.QtdeEntregueFisicoDireta Then
                                fecharPedido = False
                            End If
                        Next
                    End If

                    If Not fecharPedido Then
                        MsgBox(Me.Page, "Fiscal Do Pedido não pode ser Fechado, verifique o Físico no Extrato Do Pedido.")
                        Exit Sub
                    End If
                End If

                If objPedido.SubOperacao.EstoqueFiscal Then
                    If (objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRAS OrElse
                       objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRASAORDEM OrElse
                       objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
                       objPedido.SubOperacao.Classe = eClassesOperacoes.VENDASAORDEM OrElse
                       objPedido.SubOperacao.Classe = eClassesOperacoes.CONTAEORDEM OrElse
                       objPedido.SubOperacao.Classe = eClassesOperacoes.LUCROREAL) Then

                        For Each item In objPedido.SaldoItensPedido
                            If Not item.QtdeContratadoFiscal = item.QtdeEntregueFiscalDireta Then
                                fecharPedido = False
                            End If
                        Next
                    End If

                    If objPedido.SubOperacao.Classe = eClassesOperacoes.GLOBAL Then
                        For Each item In objPedido.SaldoItensPedido
                            If Not item.QtdeContratadoFisico = item.QtdeEntregueFiscalGlobal Then
                                fecharPedido = False
                            End If
                        Next
                    End If

                    If objPedido.SubOperacao.Classe = eClassesOperacoes.REMESSAS Then
                        For Each item In objPedido.SaldoItensPedido
                            If Not item.QtdeContratadoFisico = item.QtdeEntregueFiscalRemessa Then
                                fecharPedido = False
                            End If
                        Next
                    End If

                    If Not fecharPedido Then
                        MsgBox(Me.Page, "Fiscal Do Pedido não pode ser Fechado, verifique o Fiscal no Extrato Do Pedido.")
                        Exit Sub
                    End If
                End If
            End If

            MudarSituacao("FISCAL", True)
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnFinanceiro_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFinanceiro.Click
        Try
            SessaoRecuperaPedido()

            If objPedido.SubOperacao.Financeiro AndAlso objPedido.FinanceiroAberto Then
                Dim valorPedido As Decimal = 0

                For Each item In objPedido.Itens.Where(Function(x) x.IUD <> "D")
                    valorPedido += item.Encargos.Where(Function(s) s.CodigoEncargo = "LIQUIDO").Sum(Function(t) t.Valor)
                Next

                If valorPedido <> (objPedido.Vencimentos.OficialMovimentado - objPedido.Vencimentos.OficialDevolvido) Then
                    MsgBox(Me.Page, "Movimento Financeiro Do Pedido não pode ser Fechado, verifique o Extrato Do Pedido.")
                    Exit Sub
                End If
            End If

            MudarSituacao("FINANCEIRO", True)
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub MudarSituacao(ByVal Origem As String, ByVal Mudar As Boolean)
        If Origem = "FISCAL" Then
            If objPedido.FiscalAberto Then
                If Mudar Then
                    If Not Funcoes.VerificaPermissao("PEDIDOFISCALFECHAR", "ALTERAR") Then
                        MsgBox(Me.Page, "Usuario Sem Permissão para Fechar o movimento Fiscal Do Pedido")
                        Exit Sub
                    End If
                    lnkAtualizar.Parent.Visible = True
                    objPedido.FiscalAberto = False
                    btnFiscal.BackColor = Drawing.Color.Red
                    btnFiscal.Text = "Fiscal Fechado"
                Else
                    btnFiscal.BackColor = Drawing.Color.Green
                    btnFiscal.Text = "Fiscal Aberto"
                End If
            Else
                If objPedido.Empresa.Empresa.PedidoBloqueado Then
                    If Not objPedido.IUD = "I" AndAlso
                       (objPedido.SubOperacao.Classe = eClassesOperacoes.COMPRAS OrElse
                        objPedido.SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse
                        objPedido.SubOperacao.Classe = eClassesOperacoes.LUCROREAL OrElse
                        objPedido.SubOperacao.Classe = eClassesOperacoes.GLOBAL) Then
                        btnOkVencimento.Enabled = False
                        cmbCondicoes.Enabled = False
                    End If
                Else
                    btnOkVencimento.Enabled = False
                    cmbCondicoes.Enabled = False
                End If
                If Mudar Then
                    If Not Funcoes.VerificaPermissao("PEDIDOFISCALABRIR", "ALTERAR") Then
                        MsgBox(Me.Page, "Usuario Sem Permissão para Liberar o movimento Fiscal Do Pedido")
                        Exit Sub
                    End If
                    objPedido.FiscalAberto = True
                    btnFiscal.BackColor = Drawing.Color.Green
                    btnFiscal.Text = "Fiscal Aberto"
                    btnOkVencimento.Enabled = True
                    cmbCondicoes.Enabled = True
                    lnkAtualizar.Parent.Visible = True
                Else
                    btnFiscal.BackColor = Drawing.Color.Red
                    btnFiscal.Text = "Fiscal Fechado"
                End If
            End If
        ElseIf Origem = "FINANCEIRO" Then
            If objPedido.FinanceiroAberto Then
                If Mudar Then
                    If Not Funcoes.VerificaPermissao("PEDIDOFINANCEIROFECHAR", "ALTERAR") Then
                        MsgBox(Me.Page, "Usuario Sem Permissão para Fechar o movimento Financeiro Do Pedido")
                        Exit Sub
                    End If
                    lnkAtualizar.Parent.Visible = True
                    objPedido.FinanceiroAberto = False
                    btnFinanceiro.BackColor = Drawing.Color.Red
                    btnFinanceiro.Text = "Financ.Fechado"
                Else
                    btnFinanceiro.BackColor = Drawing.Color.Green
                    btnFinanceiro.Text = "Financ.Aberto"
                End If
            Else
                btnOkVencimento.Enabled = False
                cmbCondicoes.Enabled = False

                If Mudar Then
                    If Not Funcoes.VerificaPermissao("PEDIDOFINANCEIROABRIR", "ALTERAR") Then
                        MsgBox(Me.Page, "Usuario Sem Permissão para Liberar o movimento Financeiro Do Pedido")
                        Exit Sub
                    End If
                    lnkAtualizar.Parent.Visible = True
                    objPedido.FinanceiroAberto = True
                    btnFinanceiro.BackColor = Drawing.Color.Green
                    btnFinanceiro.Text = "Financ.Aberto"
                    btnOkVencimento.Enabled = True
                    cmbCondicoes.Enabled = True
                Else
                    btnFinanceiro.BackColor = Drawing.Color.Red
                    btnFinanceiro.Text = "Financ.Fechado"
                End If
            End If
        End If
    End Sub

#End Region

#Region "Principal Cabeçalho"
    '****************************************
    '******  Primeira Coluna ****************
    '****************************************
    Protected Sub imgExtratoPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgExtratoPedido.Click
        Try
            If txtRegistro.Text.Length = 0 Then
                MsgBox(Me.Page, "Consulte o Registro para visualização Do Extrato")
            ElseIf cmbEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa Do Registro não encontrada")
            ElseIf txtCodigoCliente.Value.ToString.Length = 0 Then
                MsgBox(Me.Page, "Cliente Do Registro não encontrado")
            ElseIf txtRegistro.Text.Length = 0 OrElse txtRegistro.Text = "0" Then
                MsgBox(Me.Page, "Registro sem Pedido não pode ser visualizado")
            Else
                SessaoRecuperaPedido()
                Extrato.Emitir(Me.Page, False, objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, "T", objPedido.Codigo)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtPedido_TextChanged(sender As Object, e As EventArgs) Handles txtPedido.TextChanged
        Try
            SessaoRecuperaPedido()
            objPedido.PedidoEfetivo = txtPedido.Text
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtContrato_TextChanged(sender As Object, e As EventArgs) Handles txtContrato.TextChanged
        Try
            SessaoRecuperaPedido()
            objPedido.Contrato = txtContrato.Text
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtXPedNFe_TextChanged(sender As Object, e As EventArgs) Handles txtXPedNFe.TextChanged
        Try
            SessaoRecuperaPedido()
            objPedido.XPedNFe = txtXPedNFe.Text
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtItemXPedNFe_TextChanged(sender As Object, e As EventArgs) Handles txtItemXPedNFe.TextChanged
        Try
            SessaoRecuperaPedido()
            objPedido.ItemXPedNFe = txtItemXPedNFe.Text
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtDataPedido_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If IsDate(txtDataPedido.Text) Then
                If IsDate(txtDataVencimentoPedido.Text) AndAlso CDate(txtDataPedido.Text) > CDate(txtDataVencimentoPedido.Text) Then
                    MsgBox(Me.Page, "Data Do Pedido nao pode ser maior que a Data Do Vencimento Do Pedido")
                    txtDataPedido.Text = txtDataVencimentoPedido.Text
                End If

                SessaoRecuperaPedido()
                objPedido.DataPedido = CDate(txtDataPedido.Text)

                If cmbIndexador.SelectedIndex > 0 AndAlso Not cmbIndexador.SelectedValue = 99 Then
                    txtIndiceFixado.Text = Funcoes.PegarValorConversao(cmbIndexador.SelectedValue, objPedido.DataPedido).ToString("N8")
                Else
                    txtIndiceFixado.Text = "0,00000000"
                End If

                If objPedido.IndiceFixado <> CDec(txtIndiceFixado.Text) Then
                    objPedido.IndiceFixado = CDec(txtIndiceFixado.Text)
                    'If objPedido.Itens.Count > 0 AndAlso Not objPedido.AlteracaoIndiceFixado() Then
                    '    MsgBox(Me.Page, "Erro ao Atualizar os lancamentos para o novo indice")
                    '    SessaoRecuperaPedido()
                    '    txtDataPedido.Text = objPedido.DataPedido.ToString("dd/MM/yyyy")
                    '    txtIndiceFixado.Text = objPedido.IndiceFixado
                    '    Exit Sub
                    'End If
                End If

                SessaoSalvaPedido()
            Else
                MsgBox(Me.Page, "Data Do Pedido não foi informada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtDataEntregaInicial_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If IsDate(txtDataEntregaInicial.Text) Then
                If CDate(txtDataEntregaInicial.Text) > CDate(txtDataEntregaFinal.Text) Then
                    MsgBox(Me.Page, "Data de Entrega Inicial não pode ser maior que Data de Entrega Final")
                    txtDataEntregaInicial.Text = txtDataEntregaFinal.Text
                End If

                SessaoRecuperaPedido()
                objPedido.DataEntregaInicial = CDate(txtDataEntregaInicial.Text)
                SessaoSalvaPedido()
            Else
                MsgBox(Me.Page, "Data de Entrega Inicial não foi informada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtDataEntregaFinal_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If IsDate(txtDataEntregaFinal.Text) Then
                If CDate(txtDataEntregaFinal.Text) < CDate(txtDataEntregaInicial.Text) Then
                    MsgBox(Me.Page, "Data de Entrega Final não pode ser menor que Data de Entrega Inicial")
                    txtDataEntregaFinal.Text = txtDataEntregaInicial.Text
                End If

                SessaoRecuperaPedido()
                objPedido.DataEntregaFinal = CDate(txtDataEntregaFinal.Text)
                SessaoSalvaPedido()
            Else
                MsgBox(Me.Page, "Data de Entrega Final não foi informada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtDataVencimentoPedido_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If IsDate(txtDataVencimentoPedido.Text) Then
                If CDate(txtDataVencimentoPedido.Text) < CDate(txtDataPedido.Text) Then
                    MsgBox(Me.Page, "Data Do Vencimento Do Pedido não pode ser menor que a Data Do Pedido")
                    txtDataVencimentoPedido.Text = txtDataPedido.Text
                End If

                SessaoRecuperaPedido()
                objPedido.DataVencimentoPedido = CDate(txtDataVencimentoPedido.Text)
                SessaoSalvaPedido()
            Else
                MsgBox(Me.Page, "Data Do Vencimento da Troca não foi informada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkPeso_CheckedChanged(sender As Object, e As EventArgs) Handles chkPeso.CheckedChanged
        Try
            SessaoRecuperaPedido()
            objPedido.OrigemDestino = IIf(chkPeso.Checked, "1", "0")
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    '****************************************
    '*******  Segunda Coluna ****************
    '****************************************
    Protected Sub cmbFinalidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFinalidade.SelectedIndexChanged
        Try
            If cmbFinalidade.SelectedIndex > 0 Then
                SessaoRecuperaPedido()
                objPedido.CodigoFinalidade = cmbFinalidade.SelectedValue
                SessaoSalvaPedido()
            Else
                MsgBox(Me.Page, "Finalidade deve ser selecionada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbSafra_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSafra.SelectedIndexChanged
        Try
            If cmbSafra.SelectedIndex > 0 Then
                SessaoRecuperaPedido()

                objPedido.CodigoSafra = cmbSafra.SelectedValue

                '****** Tabelas de Comissão ********
                For Each Rep In objPedido.Representantes
                    For Each Tabela In Rep.Comissoes.Where(Function(x) x.IUD <> "D")
                        If Tabela.CodigoTabela > 0 Then
                            Tabela.IUD = "D"
                        Else
                            Tabela = Nothing
                        End If
                    Next
                    Rep.Comissoes.AtualizarLista()
                Next
                LimparComissoes()
                '*************************************

                Dim S As New [Lib].Negocio.Safra(cmbSafra.SelectedValue)
                objPedido.Taxa = S.Taxa
                objPedido.DataVencimentoPedido = S.Vencimento
                txtTaxa.Text = S.Taxa.ToString("N2")
                txtDataVencimentoPedido.Text = S.Vencimento.ToString("dd/MM/yyyy")
                If objPedido.IUD = "I" Then
                    objPedido.DataEntregaFinal = S.Vencimento
                    txtDataEntregaFinal.Text = S.Vencimento.ToString("dd/MM/yyyy")
                End If
                SessaoSalvaPedido()
            Else
                MsgBox(Me.Page, "Informe Cliente e Safra antes de selecionar o pedido de Troca")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbMoeda_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbMoeda.SelectedIndexChanged
        Try
            If cmbMoeda.SelectedIndex = 0 Then
                txtIndiceFixado.Text = 0
                txtIndiceFixado.Enabled = False
            Else
                SessaoRecuperaPedido()
                objPedido.CodigoMoeda = cmbMoeda.SelectedValue

                ddl.Carregar(cmbIndexador, CarregarDDL.Tabela.Indexador, "indexador_id = 99 Or isnull(Moeda,0) = " & objPedido.CodigoMoeda, False)

                If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    cmbIndexador.Visible = False
                    txtIndiceFixado.Visible = False
                    imgAjudaFixarDolar1.Visible = False

                    chkTemVariacao.Checked = False
                    chkTemVariacao.Visible = False
                    objPedido.TemVariacao = False

                    objPedido.CodigoIndexador = 99

                    txtIndiceFixado.Text = "0.00"
                Else

                    Dim objCotacao As New [Lib].Negocio.Cotacao(cmbIndexador.SelectedValue, DateTime.Now)
                    If Not objCotacao.Realizado Then
                        lnkNovoPedido.Parent.Visible = False
                        MsgBox(Me.Page, "Verifique a cotação Do dólar, o mesmo ainda não foi realizado. Qualquer dúvida entre em contato com o suporte!")
                        Exit Sub
                    End If

                    cmbIndexador.Visible = True
                    txtIndiceFixado.Visible = True
                    imgAjudaFixarDolar1.Visible = True

                    chkTemVariacao.Visible = True
                    objPedido.CodigoIndexador = cmbIndexador.SelectedValue
                    txtIndiceFixado.Enabled = True

                    If Not objPedido.CodigoIndexador = 99 Then
                        txtIndiceFixado.Text = Funcoes.PegarValorConversao(objPedido.CodigoIndexador, objPedido.DataPedido).ToString("N8")
                    Else
                        txtIndiceFixado.Text = "0,00000000"
                    End If
                    objPedido.IndiceFixado = CDec(txtIndiceFixado.Text)
                End If
                SessaoSalvaPedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbIndexador_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbIndexador.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()
            objPedido.CodigoIndexador = cmbIndexador.SelectedValue

            If Not objPedido.CodigoIndexador = 99 Then
                txtIndiceFixado.Text = Funcoes.PegarValorConversao(objPedido.CodigoIndexador, objPedido.DataPedido).ToString("N8")
            Else
                txtIndiceFixado.Text = "0,00000000"
            End If

            If objPedido.IndiceFixado <> CDec(txtIndiceFixado.Text) Then
                objPedido.IndiceFixado = CDec(txtIndiceFixado.Text)
            End If

            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function LiberarIndiceFixado() As Boolean
        If Not objPedido.FiscalAberto Then
            Return False
        ElseIf Not objPedido.FinanceiroAberto Then
            Return False
        ElseIf objPedido.PedidoBloqueado Then
            Return False
        ElseIf objPedido.TemFaturamento Then
            Return False
        Else
            If objPedido.TemFinanceiro(1) Then Return False
        End If

        lnkAtualizar.Enabled = True
        Return True
    End Function

    Protected Sub txtIndiceFixado_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtIndiceFixado.TextChanged
        Try
            SessaoRecuperaPedido()

            If objPedido.IUD <> "I" AndAlso (objPedido.TemFaturamento Or objPedido.TemFinanceiro(1)) Then
                'Processo para Liberar ajuste do índice do dólar fixado - furlan - 09/10/2024
                If Not Funcoes.VerificaPermissao("LIBERAINDICEFIXADO", "GRAVAR") Then
                    MsgBox(Me.Page, "Pedidos com notas emitidas nao podem alterar o indice Do dolar")
                    txtIndiceFixado.Text = objPedido.IndiceFixado
                    Exit Sub
                End If
            End If

            If Not IsNumeric(txtIndiceFixado.Text) OrElse CDec(txtIndiceFixado.Text) <= 0 Then Exit Sub

            objPedido.IndiceFixado = txtIndiceFixado.Text

            If objPedido.IUD = "U" AndAlso objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                For Each item In objPedido.Itens
                    If item.QuantidadePedidoFaturamento > 0 Then
                        If objPedido.IndexadorFixo Then
                            For Each lcto In item.Lancamentos
                                If lcto.QuantidadeFaturamento > 0 Then
                                    item.IUD = "U"
                                    lcto.IUD = "U"
                                    If objPedido.IndexadorFixo Then
                                        lcto.UnitarioOficial = Funcoes.ConverteMoeda(lcto.UnitarioMoeda, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 10)
                                        lcto.TotalOficial = Funcoes.ConverteMoeda(lcto.TotalMoeda, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 10)
                                    Else
                                        lcto.UnitarioOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(lcto.UnitarioMoeda, objPedido.Indexador.Codigo, lcto.Movimento), 10, MidpointRounding.AwayFromZero)
                                        lcto.TotalOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(lcto.TotalMoeda, objPedido.Indexador.Codigo, lcto.Movimento), 2, MidpointRounding.AwayFromZero)
                                    End If
                                End If
                            Next

                            For Each enc In item.Encargos
                                enc.IUD = "U"
                                Dim baseOficial As Decimal = Math.Round(enc.BaseMoeda * objPedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                If baseOficial <> enc.ValorOficial Then enc.BaseOficial = baseOficial

                                Dim vlrOficial As Decimal = Math.Round(enc.ValorMoeda * objPedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                If vlrOficial <> enc.ValorOficial Then enc.ValorOficial = vlrOficial
                            Next
                        End If
                    End If

                    item.Encargos.AtualizaLiquido()
                Next
            End If

            SessaoSalvaPedido()

            If objPedido.Itens.Count > 0 Then AtualizarResumoItens()
            If cmbCondicoes.SelectedIndex > 0 Then CalcularParcelamento(objPedido.Itens.Count > 1) Else LimparCondicoes(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkIndexadorFixo_CheckedChanged(sender As Object, e As EventArgs) Handles chkIndexadorFixo.CheckedChanged
        Try
            SessaoRecuperaPedido()

            objPedido.IndexadorFixo = chkIndexadorFixo.Checked

            If objPedido.IUD = "U" AndAlso objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                For Each item In objPedido.Itens
                    If item.QuantidadePedidoFaturamento > 0 Then
                        If objPedido.IndexadorFixo Then
                            For Each lcto In item.Lancamentos
                                If lcto.QuantidadeFaturamento > 0 Then
                                    item.IUD = "U"
                                    lcto.IUD = "U"
                                    If objPedido.IndexadorFixo Then
                                        lcto.UnitarioOficial = Funcoes.ConverteMoeda(lcto.UnitarioMoeda, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 10)
                                        lcto.TotalOficial = Funcoes.ConverteMoeda(lcto.TotalMoeda, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 10)
                                    Else
                                        lcto.UnitarioOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(lcto.UnitarioMoeda, objPedido.Indexador.Codigo, lcto.Movimento), 10, MidpointRounding.AwayFromZero)
                                        lcto.TotalOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(lcto.TotalMoeda, objPedido.Indexador.Codigo, lcto.Movimento), 2, MidpointRounding.AwayFromZero)
                                    End If
                                End If
                            Next

                            For Each enc In item.Encargos
                                enc.IUD = "U"
                                Dim baseOficial As Decimal = Math.Round(enc.BaseMoeda * objPedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                If baseOficial <> enc.ValorOficial Then enc.BaseOficial = baseOficial

                                Dim vlrOficial As Decimal = Math.Round(enc.ValorMoeda * objPedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                If vlrOficial <> enc.ValorOficial Then enc.ValorOficial = vlrOficial
                            Next
                        End If
                    End If
                    item.Encargos.AtualizaLiquido()
                Next
            End If

            SessaoSalvaPedido()

            If objPedido.Itens.Count > 0 Then AtualizarResumoItens()
            If cmbCondicoes.SelectedIndex > 0 Then CalcularParcelamento(objPedido.Itens.Count > 1) Else LimparCondicoes(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtTaxa_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaPedido()
            objPedido.Taxa = IIf(IsNumeric(txtTaxa.Text), txtTaxa.Text, 0)
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkAgruparFinanceiro_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaPedido()
            objPedido.AgruparFinanceiro = chkAgruparFinanceiro.Checked
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlComercializacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlComercializacao.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()
            objPedido.Comercializacao = ddlComercializacao.SelectedValue
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    '*****************************************
    '*******  Terceira Coluna ****************
    '*****************************************
    Sub VerificaUnidade()
        Dim Sql As String = ""
        Sql = "Select isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf &
              "       isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf &
              "       isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf &
              "  from Usuarios" & vbCrLf &
              " where Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            cmbUnidadeNegocio.SelectedValue = Dr("AcessoUnidade")
            ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.Empresas, Dr("AcessoUnidade"), True)
            cmbEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")

            objPedido.CodigoUnidadeNegocio = Dr("AcessoUnidade")
            objPedido.CodigoEmpresa = Dr("AcessoEmpresa")
            objPedido.EnderecoEmpresa = Dr("AcessoEndEmpresa")
        Next
    End Sub

    Protected Sub cmbUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbUnidadeNegocio.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()
            ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.Empresas, cmbUnidadeNegocio.SelectedValue, True)
            objPedido.CodigoUnidadeNegocio = cmbUnidadeNegocio.SelectedValue
            objPedido.CodigoEmpresa = ""
            objPedido.EnderecoEmpresa = 0
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbEmpresa.SelectedIndexChanged
        Try
            If cmbEmpresa.SelectedIndex > 0 Then
                SessaoRecuperaPedido()
                Dim Emp As String() = cmbEmpresa.SelectedValue.ToString.Split("-")
                objPedido.CodigoEmpresa = Emp(0)
                objPedido.EnderecoEmpresa = Emp(1)

                If Not objPedido.Vencimentos Is Nothing AndAlso objPedido.Vencimentos.Count > 0 Then
                    For Each vencimento In objPedido.Vencimentos
                        vencimento.CodigoEmpresa = Emp(0)
                        vencimento.EnderecoEmpresa = Emp(1)

                        vencimento.CodigoEmpresaPedido = Emp(0)
                        vencimento.EnderecoEmpresaPedido = Emp(1)
                    Next
                End If

                txtContrato.Text = ""
                txtXPedNFe.Text = ""
                txtItemXPedNFe.Text = ""

                objPedido.Empresa = New [Lib].Negocio.Cliente(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa)
                If objPedido.Empresa.Empresa.PedidoBloqueado Then
                    btnBloqueado.Parent.Visible = True
                    btnBloqueado.BackColor = Drawing.Color.Green
                    btnBloqueado.Text = "Liberado"
                Else
                    btnBloqueado.Parent.Visible = False
                End If
                SessaoSalvaPedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOperacao.SelectedIndexChanged
        Try
            If cmbOperacao.SelectedIndex > 0 Then
                Dim Where As String = String.Empty
                SessaoRecuperaPedido()
                objPedido.CodigoOperacao = cmbOperacao.SelectedValue
                SessaoSalvaPedido()

                'Usado para listar apenas suboperações com situação 1 - Normal na inclusão de pedidos.
                If objPedido.IUD = "I" Then
                    Where = " So.Operacao_Id = " & objPedido.CodigoOperacao & " and So.Pedido = 1 AND Situacao = 1 "
                Else
                    Where = " So.Operacao_Id = " & objPedido.CodigoOperacao & " and So.Pedido = 1"
                End If
                ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.SubOperacaoSemOperacao, Where)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbSubOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSubOperacao.SelectedIndexChanged
        Try
            If Not cmbSubOperacao.SelectedIndex > 0 Then Exit Sub

            SessaoRecuperaPedido()
            objPedido.CodigoSubOperacao = cmbSubOperacao.SelectedValue

            If objPedido.SubOperacao.Financeiro Then
                TabVencimentosOld.Visible = objPedido.MomentoFinanceiro <> 9 And (objPedido.IUD = "I" And Not FinanceiroVirtual)
                TabParcelamento.Visible = (objPedido.MomentoFinanceiro = 9 Or objPedido.IUD = "I") And FinanceiroVirtual
            Else
                TabVencimentosOld.Visible = False
                TabParcelamento.Visible = False
            End If

            divEntrega.Visible = False

            If FinanceiroVirtual Then
                objPedido.MomentoFinanceiro = 9
                ddlMomentoFinanceiro.SelectedValue = 9
            ElseIf Not objPedido.SubOperacao.Financeiro Then
                objPedido.MomentoFinanceiro = 0
                ddlMomentoFinanceiro.SelectedValue = 0
            ElseIf objPedido.Troca Then
                objPedido.MomentoFinanceiro = 4
                ddlMomentoFinanceiro.SelectedValue = 4
                ddlMomentoFinanceiro.Enabled = False
            ElseIf Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "04854422" Then
                objPedido.MomentoFinanceiro = 2
                ddlMomentoFinanceiro.SelectedValue = 2
                ddlMomentoFinanceiro.Enabled = False
            ElseIf Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "03961253" Then
                objPedido.MomentoFinanceiro = 2
                ddlMomentoFinanceiro.SelectedValue = 2
                ddlMomentoFinanceiro.Enabled = True
            Else
                objPedido.MomentoFinanceiro = 3
                ddlMomentoFinanceiro.SelectedValue = 3
                ddlMomentoFinanceiro.Enabled = False
                If Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "04440724" Then
                    ddlMomentoFinanceiro.Enabled = True
                    divEntrega.Visible = True
                End If
            End If

            lblCifFob.Text = "Frete pela Empresa:"
            If objPedido.SubOperacao.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida Then
                rbCIF.Text = "Sim (CIF)"
                rbFOB.Text = "Não (FOB)"
            Else
                rbCIF.Text = "Não (CIF)"
                rbFOB.Text = "Sim (FOB)"
            End If

            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultarCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultarCliente.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("4,5")
            Popup.ConsultaDeClientes(Me, "objClientePXI" & HID.Value.ToString, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultarNaviosXInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkConsultarNaviosXInvoice.Click
        Try
            SessaoRecuperaPedido()

            If objPedido.Itens.Count = 0 Then
                MsgBox(Me.Page, "Produto não foi selecionado.")
            Else
                ucConsultarNaviosXInvoice.Limpar()
                ucConsultarNaviosXInvoice.carregarNavios(objPedido.Itens(0).CodigoProduto)
                Popup.ConsultarNaviosXInvoice(Me.Page, "objNavioXInvoice" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaPraca_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("4,5")
            Popup.ConsultaDeClientes(Me, "objPracaPXI" & HID.Value.ToString, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaDadosBancarios_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If txtCodigoCliente.Value.Length > 0 Or txtCodigoPraca.Value.Length > 0 Then
                Dim strCliente As String()
                If txtCodigoPraca.Value.Length > 0 Then
                    strCliente = txtCodigoPraca.Value.Split("-")
                Else
                    strCliente = txtCodigoCliente.Value.Split("-")
                End If
                ucConsultaDadosBancarios.CarregaGrid(strCliente(0), strCliente(1))
                Popup.ConsultaDeDadosBancarios(Me.Page, "objContaBancariaPXI" & HID.Value.ToString)
            Else
                MsgBox(Me.Page, "Selecione primeiro um cliente ou pr pagto!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    '*********************************************
    '*******  Rodape do Cabeçalho ****************
    '*********************************************
    Protected Sub rbCIF_CheckedChanged(sender As Object, e As EventArgs) Handles rbCIF.CheckedChanged
        Try
            SessaoRecuperaPedido()
            objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.CIF
            SessaoSalvaPedido()
            'LinhaValorFrete.Visible = EnabledValorFrete(ddlTipoDeposito.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rbFOB_CheckedChanged(sender As Object, e As EventArgs) Handles rbFOB.CheckedChanged
        Try
            SessaoRecuperaPedido()
            objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.FOB
            SessaoSalvaPedido()

            txtLocalDeEmbarque.Enabled = True
            'LinhaValorFrete.Visible = EnabledValorFrete(ddlTipoDeposito.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rbTER_CheckedChanged(sender As Object, e As EventArgs) Handles rbTER.CheckedChanged
        Try
            SessaoRecuperaPedido()
            objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.TER
            SessaoSalvaPedido()
            'LinhaValorFrete.Visible = EnabledValorFrete(ddlTipoDeposito.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rbNenhum_CheckedChanged(sender As Object, e As EventArgs) Handles rbNenhum.CheckedChanged
        Try
            SessaoRecuperaPedido()
            objPedido.FreteCIFFOB = [Lib].Negocio.eTiposFrete.NEN
            SessaoSalvaPedido()
            'LinhaValorFrete.Visible = EnabledValorFrete(ddlTipoDeposito.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnUFEntrega_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Popup.ConsultaDeEstados(Me.Page, "objEstadoPDXI" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkTroca_CheckedChanged(sender As Object, e As EventArgs) Handles chkTroca.CheckedChanged
        Try
            SessaoRecuperaPedido()
            TabPedidoTroca.Visible = chkTroca.Checked
            objPedido.Troca = chkTroca.Checked

            If objPedido.Troca Then
                If objPedido.PedidoTroca IsNot Nothing AndAlso objPedido.PedidoTroca.Codigo > 0 Then
                    LnkVincular.Parent.Visible = False
                    LnkDesvincular.Parent.Visible = True
                ElseIf objPedido.Operacao IsNot Nothing AndAlso objPedido.Operacao.CodigoClasse = eClassesOperacoes.COMPRAS.ToString Then
                    LnkVincular.Parent.Visible = True
                    LnkDesvincular.Parent.Visible = False
                Else
                    LnkVincular.Parent.Visible = False
                    LnkDesvincular.Parent.Visible = False
                End If

                If Not FinanceiroNovo Then
                    'LIBERAR FEX PARA GERAR FINANCEIRO NA TROCA - FURLAN - 08-12-2015
                    If Not Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "15204808" Then
                        TabVencimentosOld.Visible = Not objPedido.Troca
                    End If
                End If
            Else
                If objPedido.SubOperacao.Financeiro Then
                    TabVencimentosOld.Visible = objPedido.MomentoFinanceiro <> 9
                    TabParcelamento.Visible = (objPedido.MomentoFinanceiro = 9 Or objPedido.IUD = "I") And FinanceiroVirtual
                Else
                    TabVencimentosOld.Visible = False
                    TabParcelamento.Visible = False
                End If

                objPedido.CodigoEmpresaTroca = String.Empty
                objPedido.EnderecoEmpresaTroca = 0
                objPedido.CodigoPedidoTroca = 0
            End If

            If objPedido.IUD = "I" And objPedido.Troca Then
                'LIBERAR FEX PARA GERAR FINANCEIRO NA TROCA - FURLAN - 08-12-2015
                If Not Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "15204808" Then
                    objPedido.Vencimentos = Nothing
                    GridCondicoes.DataSource = Nothing
                    GridCondicoes.DataBind()
                End If
            End If

            If objPedido.IUD = "U" And objPedido.Troca And objPedido.Vencimentos.Count > 0 Then
                If objPedido.Vencimentos.Where(Function(s) s.Provisao = eProvisao.Baixa).Count > 0 Then
                    'LIBERAR FEX PARA GERAR FINANCEIRO NA TROCA - FURLAN - 08-12-2015
                    If Not Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "15204808" Then
                        objPedido.Troca = False
                        chkTroca.Checked = False
                        TabVencimentosOld.Visible = True
                        MsgBox(Me.Page, "Pedidos com Financeiro baixado nao pode se tornar uma troca")
                        Exit Sub
                    End If
                End If
            End If

            SessaoSalvaPedido()
            If chkTroca.Checked Then
                tcPedido.ActiveTab = TabPedidoTroca
            ElseIf objPedido.SubOperacao.Financeiro Then
                tcPedido.ActiveTab = TabVencimentosOld
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkAntecipada_CheckedChanged(sender As Object, e As EventArgs) Handles chkAntecipada.CheckedChanged
        Try
            SessaoRecuperaPedido()
            objPedido.Antecipada = chkAntecipada.Checked
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkRecompra_CheckedChanged(sender As Object, e As EventArgs) Handles chkRecompra.CheckedChanged
        Try
            SessaoRecuperaPedido()
            objPedido.Recompra = chkRecompra.Checked
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Tab Produto"
    Protected Sub lnkAdicionarItem_Click(sender As Object, e As EventArgs) Handles lnkAdicionarItem.Click
        Try
            If String.IsNullOrWhiteSpace(cmbFinalidade.SelectedValue) Then
                MsgBox(Me.Page, "É necessário selecionar a Finalidade!")
                Exit Sub
            End If
            If String.IsNullOrWhiteSpace(cmbSafra.SelectedValue) Then
                MsgBox(Me.Page, "É necessário selecionar a Safra!")
                Exit Sub
            End If
            If String.IsNullOrWhiteSpace(cmbMoeda.SelectedValue) Then
                MsgBox(Me.Page, "É necessário selecionar a Moeda!")
                Exit Sub
            End If
            If String.IsNullOrWhiteSpace(cmbSituacao.SelectedValue) Then
                MsgBox(Me.Page, "É necessário selecionar a Situação!")
                Exit Sub
            End If
            If Not cmbMoeda.SelectedValue = 1 AndAlso (Not IsNumeric(txtIndiceFixado.Text) OrElse CDec(txtIndiceFixado.Text) <= 0) Then
                MsgBox(Me.Page, "É necessário informar uma cotação para o indexador selecionado!")
                Exit Sub
            End If
            If String.IsNullOrWhiteSpace(cmbEmpresa.SelectedValue) Then
                MsgBox(Me.Page, "É necessário selecionar a Empresa!")
                Exit Sub
            End If
            If String.IsNullOrWhiteSpace(cmbOperacao.SelectedValue) OrElse String.IsNullOrWhiteSpace(cmbSubOperacao.SelectedValue) Then
                MsgBox(Me.Page, "É necessário selecionar a operação e a sub-operação!")
                Exit Sub
            End If
            If String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "É necessário selecionar a Cliente!")
                Exit Sub
            End If

            ucPedidoXLancamento.SetarParametros(-1)
            ucPedidoXLancamento.Limpar()
            Popup.ConsultaLancamentosPedido(Me.Page, "objLancamentosPedido" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    ''' <summary>
    ''' Volta do User control de itens de lancamento Complemento / Estorno
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub AtualizarResumoItens(Optional ByVal atualizaFinan As Boolean = True)
        If objPedido Is Nothing Then SessaoRecuperaPedido()

        If objPedido Is Nothing Then
            Exit Sub
        End If

        If objPedido.Itens.Count = 0 Then
            lnkAdicionarItem.Parent.Parent.Visible = True
        ElseIf objPedido.Itens(0).Produto.Agrupar.Equals("S") Then
            lnkAdicionarItem.Parent.Parent.Visible = True
        Else
            lnkAdicionarItem.Parent.Parent.Visible = False
        End If

        If objPedido.Itens.Count > 0 Then
            cmbSafra.Enabled = False
            cmbIndexador.Enabled = False
            cmbMoeda.Enabled = False
            cmbOperacao.Enabled = False
            cmbSubOperacao.Enabled = False
            cmbSafra.Enabled = False

            cmbUnidadeNegocio.Enabled = False
            cmbEmpresa.Enabled = False
            'txtIndiceFixado.Enabled = False
        Else
            cmbSafra.Enabled = True
            cmbMoeda.Enabled = True
            cmbOperacao.Enabled = True
            cmbIndexador.Enabled = True
            cmbSubOperacao.Enabled = True
            cmbSafra.Enabled = True

            'Desabilitado Empresa - Furlan - 23/08/2022
            'cmbUnidadeNegocio.Enabled = True
            'cmbEmpresa.Enabled = True

            'txtIndiceFixado.Enabled = True
        End If

        If Not objPedido.PedidoTroca Is Nothing Then
            If objPedido.Troca AndAlso objPedido.Operacao.CodigoClasse = eClassesOperacoes.COMPRAS.ToString Then
                txtVendaTroca.Text = objPedido.PedidoTroca.Itens.LiquidoOficial
                txtCompraTroca.Text = objPedido.Itens.LiquidoOficial
                txtSaldoTroca.Text = objPedido.PedidoTroca.Itens.LiquidoOficial - objPedido.Itens.LiquidoOficial
            Else
                txtVendaTroca.Text = objPedido.Itens.LiquidoOficial
                txtCompraTroca.Text = objPedido.PedidoTroca.Itens.LiquidoOficial
                txtSaldoTroca.Text = objPedido.Itens.LiquidoOficial - objPedido.PedidoTroca.Itens.LiquidoOficial
            End If
        End If

        If objPedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS Then
            gridItensPedidoDeposito.DataSource = objPedido.Itens.ToArray
            If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                gridItensPedidoDeposito.Columns(10).Visible = False
            Else
                gridItensPedidoDeposito.Columns(10).Visible = True
            End If
            gridItensPedidoDeposito.DataBind()
        ElseIf objPedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then
            gridItensPedidoAFixar.DataSource = objPedido.Itens.ToArray
            If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                gridItensPedidoAFixar.Columns(10).Visible = False
                gridItensPedidoAFixar.Columns(13).Visible = False
            Else
                gridItensPedidoAFixar.Columns(10).Visible = True
                gridItensPedidoAFixar.Columns(13).Visible = True
            End If
            gridItensPedidoAFixar.DataBind()
        Else
            gridItensPedido.DataSource = objPedido.Itens.ToArray
            gridItensPedido.DataBind()
        End If

        'txtTotalQtdPedido.Text = objPedido.Itens.QuantidadeTotal.ToString("N2")
        txtTotalQtdPedido.Text = QuantidadeTotalPedido.ToString("N2")

        AtualizarGridEncargos()

        GridRepresentantes.DataSource = objPedido.Representantes.ToArray
        GridRepresentantes.DataBind()

        If Not objPedido.SubOperacao Is Nothing AndAlso objPedido.SubOperacao.Financeiro AndAlso atualizaFinan Then
            If cmbCondicoes.SelectedIndex > 0 Then CalcularParcelamento(False) Else LimparCondicoes(True)
        End If
    End Sub

    Public Sub AtualizarGridEncargos()
        gridEncargosGerais.DataSource = From p In objPedido.Itens.SelectMany(Function(s) s.Encargos)
                                        Group By p.CodigoEncargo
                                         Into ValorMoeda = Sum(p.ValorMoeda), ValorOficial = Sum(p.ValorOficial)
                                        Order By IIf(CodigoEncargo = "PRODUTO", 1, IIf(CodigoEncargo = "LIQUIDO", 3, 2))
                                        Select CodigoEncargo, ValorMoeda, ValorOficial

        If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
            gridEncargosGerais.Columns(2).Visible = False
        Else
            gridEncargosGerais.Columns(2).Visible = True
        End If
        gridEncargosGerais.DataBind()

        Dim cla As eTiposMoeda = objPedido.Moeda.Classificacao

        txtPedidoTotal.Text = objPedido.Itens.Liquido.ToString("N2")
        If objPedido.Itens.Liquido = objPedido.Parcelas.Total Then txtPedidoTotalPago.Text = (objPedido.Itens.Liquido - objPedido.Parcelas.ApuracaoParcela.SaldoAGerar).ToString("N2")
        If objPedido.IUD = "I" Then
            txtPedidoSaldo.Text = objPedido.Itens.Liquido.ToString("N2")
        Else
            txtPedidoSaldo.Text = objPedido.Parcelas.ApuracaoParcela.SaldoAGerar.ToString("N2")
        End If
    End Sub

    Protected Sub gridItensPedido_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridItensPedido.RowDataBound
        Try
            'If e.Row.RowType = DataControlRowType.Header Then
            '    If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
            '        e.Row.Cells(7).Text = "Vlr Entregue"
            '        e.Row.Cells(9).Text = "Vlr Saldo"
            '    Else
            '        e.Row.Cells(7).Text = "Vlr Entregue"
            '        e.Row.Cells(9).Text = "Vlr Saldo"
            '    End If
            'End If

            If e.Row.RowType = DataControlRowType.DataRow Then
                Dim ExibeMsgRetencao As Boolean
                If objPedido.Cliente.CodigoEstado = "EX" Then
                    ExibeMsgRetencao = False
                ElseIf objPedido.CodigoCliente.Length = 14 Then
                    ExibeMsgRetencao = objPedido.Itens(e.Row.RowIndex).OperacaoxEstado.Encargos.DescRetencaoPJ.Length > 0
                Else
                    ExibeMsgRetencao = objPedido.Itens(e.Row.RowIndex).OperacaoxEstado.Encargos.DescRetencaoPF.Length > 0
                End If

                If Not ExibeMsgRetencao Then
                    CType(e.Row.FindControl("LRT"), Label).Visible = False
                Else
                    If objPedido.Itens(e.Row.RowIndex).Retencao Then
                        CType(e.Row.FindControl("LRT"), Label).Text = "Com Retenção"
                        CType(e.Row.FindControl("LRT"), Label).ForeColor = Drawing.Color.Red
                    Else
                        CType(e.Row.FindControl("LRT"), Label).Text = "Sem Retenção"
                        CType(e.Row.FindControl("LRT"), Label).ForeColor = Drawing.Color.Blue
                    End If
                End If


                If Not objPedido.Itens(e.Row.RowIndex).Produto.ControlarDecimais Then
                    CType(e.Row.FindControl("lblQtdeFat"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("lblQtdeFat"), Literal).Text))
                    CType(e.Row.FindControl("lblQtdeCom"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("lblQtdeCom"), Literal).Text))
                    CType(e.Row.FindControl("L17F"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L17F"), Literal).Text))
                    CType(e.Row.FindControl("L18F"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L18F"), Literal).Text))
                    CType(e.Row.FindControl("L19F"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L19F"), Literal).Text))
                    CType(e.Row.FindControl("L17C"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L17C"), Literal).Text))
                    CType(e.Row.FindControl("L18C"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L18C"), Literal).Text))
                    CType(e.Row.FindControl("L19C"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L19C"), Literal).Text))

                    CType(e.Row.FindControl("L20F"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L20F"), Literal).Text))
                    CType(e.Row.FindControl("L21F"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L21F"), Literal).Text))
                    CType(e.Row.FindControl("L22F"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L22F"), Literal).Text))
                    CType(e.Row.FindControl("L20C"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L20C"), Literal).Text))
                    CType(e.Row.FindControl("L21C"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L21C"), Literal).Text))
                    CType(e.Row.FindControl("L22C"), Literal).Text = String.Format("{0:N0}", CDec(CType(e.Row.FindControl("L22C"), Literal).Text))
                Else
                    CType(e.Row.FindControl("lblQtdeFat"), Literal).Text = String.Format("{0:N4}", CDec(CType(e.Row.FindControl("lblQtdeFat"), Literal).Text))
                    CType(e.Row.FindControl("lblQtdeCom"), Literal).Text = String.Format("{0:N4}", CDec(CType(e.Row.FindControl("lblQtdeCom"), Literal).Text))
                    CType(e.Row.FindControl("L17F"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L17F"), Literal).Text))
                    CType(e.Row.FindControl("L18F"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L18F"), Literal).Text))
                    CType(e.Row.FindControl("L19F"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L19F"), Literal).Text))
                    CType(e.Row.FindControl("L17C"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L17C"), Literal).Text))
                    CType(e.Row.FindControl("L18C"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L18C"), Literal).Text))
                    CType(e.Row.FindControl("L19C"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L19C"), Literal).Text))

                    CType(e.Row.FindControl("L20F"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L20F"), Literal).Text))
                    CType(e.Row.FindControl("L21F"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L21F"), Literal).Text))
                    CType(e.Row.FindControl("L22F"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L22F"), Literal).Text))
                    CType(e.Row.FindControl("L20C"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L20C"), Literal).Text))
                    CType(e.Row.FindControl("L21C"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L21C"), Literal).Text))
                    CType(e.Row.FindControl("L22C"), Literal).Text = String.Format("{0:N2}", CDec(CType(e.Row.FindControl("L22C"), Literal).Text))
                End If


                e.Row.Cells(6).BackColor = Drawing.Color.LightSkyBlue
                e.Row.Cells(7).BackColor = Drawing.Color.LightSkyBlue
                e.Row.Cells(8).BackColor = Drawing.Color.LightYellow
                e.Row.Cells(9).BackColor = Drawing.Color.LightYellow

                CType(e.Row.FindControl("QEC"), Literal).Parent.Visible = Not objPedido.Itens(e.Row.RowIndex).UnidadeFaturamento = objPedido.Itens(e.Row.RowIndex).CodigoUnidadeComercializacao
                CType(e.Row.FindControl("QSC"), Literal).Parent.Visible = Not objPedido.Itens(e.Row.RowIndex).UnidadeFaturamento = objPedido.Itens(e.Row.RowIndex).CodigoUnidadeComercializacao

                CType(e.Row.FindControl("lblQtdeCom"), Literal).Visible = Not objPedido.Itens(e.Row.RowIndex).UnidadeFaturamento = objPedido.Itens(e.Row.RowIndex).CodigoUnidadeComercializacao
                CType(e.Row.FindControl("lblUnCom"), Literal).Visible = Not objPedido.Itens(e.Row.RowIndex).UnidadeFaturamento = objPedido.Itens(e.Row.RowIndex).CodigoUnidadeComercializacao
                CType(e.Row.FindControl("lblUnMdCom"), Literal).Visible = Not objPedido.Itens(e.Row.RowIndex).UnidadeFaturamento = objPedido.Itens(e.Row.RowIndex).CodigoUnidadeComercializacao

                If objPedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    CType(e.Row.FindControl("L08C"), Literal).Parent.Visible = False
                    CType(e.Row.FindControl("L14C"), Literal).Parent.Visible = False
                    CType(e.Row.FindControl("PVM"), Literal).Visible = False
                    CType(e.Row.FindControl("PVMC"), Literal).Visible = False
                Else
                    CType(e.Row.FindControl("L08C"), Literal).Parent.Visible = True
                    CType(e.Row.FindControl("L14C"), Literal).Parent.Visible = True
                    CType(e.Row.FindControl("PVM"), Literal).Visible = True
                    CType(e.Row.FindControl("PVMC"), Literal).Visible = True
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkComplementoEstorno_Click(sender As Object, e As EventArgs)
        Try
            Dim linha As Integer = CType(CType(sender, LinkButton).NamingContainer, GridViewRow).RowIndex
            SessaoRecuperaPedido()
            If objPedido.Itens.Any(Function(l) l.Lancamentos.Where(Function(x) x.IUD = "I").Count >= 1) Then
                Throw New Exception("Tem um lançamento complemento/estorno pedente para salvar")
            ElseIf objPedido.Itens(linha).Produto.CodigoSituacao <> 1 Then
                MsgBox(Me.Page, "O produto " & objPedido.Itens(linha).Produto.Codigo & " - " & objPedido.Itens(linha).Produto.Descricao & " está " & objPedido.Itens(linha).Produto.DescSituacao)
            Else
                ucPedidoXLancamento.SetarParametros(linha)
                ucPedidoXLancamento.desabilitaProdutoNovo()
                'Comentei porque no SetarParamentros acima já executa o Limpar - Furlan - 22/03/2016 (Estava limpando o preço unitário médio do produto)
                'ucPedidoXLancamento.Limpar()
                Popup.ConsultaLancamentosPedido(Me.Page, "objLancamentosPedido" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEncargos_Click(sender As Object, e As EventArgs)
        Try
            Dim linha As Integer = CType(CType(sender, LinkButton).NamingContainer, GridViewRow).RowIndex
            ucPedidoxEncargo.Limpar()
            ucPedidoxEncargo.SetarParametros(linha)
            Popup.ConsultaEncargosPedido(Me.Page, "objEncargosPedido" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Tab Depósitos"
    Private Sub LimparDepositos()
        ddlTipoDeposito.SelectedIndex = 0
        chkDepPrincipal.Checked = False
        If cmbDepositos.Items.Count > 0 Then cmbDepositos.SelectedIndex = 0
        GridDepositos.SelectedIndex = -1
    End Sub

    Protected Sub lnkAdicionarDeposito_Click(sender As Object, e As EventArgs) Handles lnkAdicionarDeposito.Click
        Try
            SessaoRecuperaPedido()

            If Not Funcoes.VerificaPermissao("PedidosXItens", IIf(objPedido.IUD = "I", "GRAVAR", "ALTERAR")) Then
                MsgBox(Me.Page, "Sem Permissão para " & IIf(objPedido.IUD = "I", "Adicionar", "Alterar") & " Depósito")
                Exit Sub
            End If

            If Not rbFOB.Checked And Not rbCIF.Checked And Not rbTER.Checked And Not rbNenhum.Checked Then
                MsgBox(Me.Page, "Para Adicionar o Deposito antes marque se o frete é CIF/FOB/TERCEIRO/NEN")
                Exit Sub
            End If

            If cmbDepositos.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Depósito não foi Selecionado")
                cmbDepositos.Focus()
                Exit Sub
            End If

            If objPedido.Itens.Count = 0 Then
                MsgBox(Me.Page, "Para Adicionar o Deposito antes Inclua 1 Produto")
                Exit Sub
            End If

            Dim dep() As String = cmbDepositos.SelectedValue.ToString.Split("-")

            If VerificarDepositos(dep(0), dep(1)) Then
                Dim objNovoDeposito As [Lib].Negocio.PedidoXDeposito
                objNovoDeposito = objPedido.Depositos.Where(Function(s) s.Codigo = dep(0) And s.CodigoEndereco = dep(1) And s.Tipo = ddlTipoDeposito.SelectedValue).FirstOrDefault

                If objNovoDeposito Is Nothing Then
                    objNovoDeposito = New [Lib].Negocio.PedidoXDeposito(objPedido)
                    objNovoDeposito.IUD = "I"
                    objNovoDeposito.Codigo = dep(0)
                    objNovoDeposito.CodigoEndereco = dep(1)
                    objNovoDeposito.Tipo = ddlTipoDeposito.SelectedValue
                    'objNovoDeposito.Principal = IIf(ddlTipoDeposito.SelectedValue = "OD", True, chkDepPrincipal.Checked)
                    objNovoDeposito.Principal = chkDepPrincipal.Checked

                    'If LinhaValorFrete.Visible AndAlso txtValorDoFrete.Text > 0 Then
                    '    objNovoDeposito.Quantidade = txtValorDoFrete.Text
                    'Else
                    '    objNovoDeposito.Quantidade = 0
                    'End If

                    'If ddlTipoDeposito.SelectedValue = "OD" Then
                    '    Dim DepDE As [Lib].Negocio.PedidoXDeposito
                    '    DepDE = objPedido.Depositos.Where(Function(s) s.Tipo = "OD" And s.IUD <> "D").FirstOrDefault
                    '    If DepDE Is Nothing Then
                    '        objPedido.Depositos.Add(objNovoDeposito)
                    '    Else
                    '        MsgBox(Me.Page, "Já existe um Deposito de Origem/Destino, por pedido só pode conter 1.")
                    '        Exit Sub
                    '    End If
                    'Else
                    objPedido.Depositos.Add(objNovoDeposito)
                    'End If

                Else
                    If objNovoDeposito.IUD <> "I" Then objNovoDeposito.IUD = "U"
                End If

                If objNovoDeposito.Principal Then
                    Dim objAntigoPrincipal As [Lib].Negocio.PedidoXDeposito = objPedido.Depositos.Where(Function(s) s.Tipo = objNovoDeposito.Tipo And s.Principal = True And (s.Codigo <> objNovoDeposito.Codigo Or s.CodigoEndereco <> objNovoDeposito.CodigoEndereco)).FirstOrDefault
                    If Not objAntigoPrincipal Is Nothing Then
                        If objAntigoPrincipal.IUD <> "I" Then objAntigoPrincipal.IUD = "U"
                        objAntigoPrincipal.Principal = False
                    End If
                Else
                    If objNovoDeposito.Tipo = "DE" OrElse objNovoDeposito.Tipo = "OD" OrElse objNovoDeposito.Tipo = "LE" OrElse objNovoDeposito.Tipo = "PM" Then
                        Dim i As Integer = objPedido.Depositos.FindIndex(Function(s) s.Tipo = objNovoDeposito.Tipo And s.Principal = True)
                        objNovoDeposito.Principal = (i = -1)
                    End If
                End If

                LimparDepositos()

                GridDepositos.DataSource = objPedido.Depositos.Where(Function(s) s.IUD <> "D").ToArray
                GridDepositos.DataBind()

                SessaoSalvaPedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlTipoDeposito_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlTipoDeposito.SelectedIndexChanged
        Try
            If Not String.IsNullOrWhiteSpace(ddlTipoDeposito.SelectedValue) Then
                'Habilita ou Nao o valor do Frete no Pedido
                'LinhaValorFrete.Visible = EnabledValorFrete(ddlTipoDeposito.SelectedValue)

                SessaoRecuperaPedido()
                If ddlTipoDeposito.SelectedValue = "DE" Then
                    ddl.Carregar(cmbDepositos, CarregarDDL.Tabela.Depositos, "", True)
                    cmbDepositos.SelectedValue = objPedido.CodigoEmpresa & "-" & CStr(objPedido.EnderecoEmpresa)
                ElseIf ddlTipoDeposito.SelectedValue = "OD" Then

                    If objPedido.CodigoEmpresa = objPedido.CodigoCliente AndAlso objPedido.EnderecoEmpresa = objPedido.EnderecoCliente Then
                        BuscarDepositos()
                    Else
                        cmbDepositos.Items.Clear()
                        cmbDepositos.DataTextField = "Descricao"
                        cmbDepositos.DataValueField = "Codigo"

                        cmbDepositos.Items.Add(New ListItem(Funcoes.AlinharEsquerda(objPedido.Cliente.Nome, 28, ".") &
                                                            " - " & Funcoes.AlinharEsquerda(objPedido.Cliente.Cidade, 20, ".") &
                                                            " " & objPedido.Cliente.CodigoEstado &
                                                            " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(objPedido.Cliente.Codigo), 18, ".") &
                                                            "-" & CStr(objPedido.Cliente.CodigoEndereco) & "-" & objPedido.Cliente.Reduzido,
                                                            objPedido.Cliente.Codigo & "-" & CStr(objPedido.Cliente.CodigoEndereco)))

                        Funcoes.InserirLinhaEmBranco(cmbDepositos)

                        cmbDepositos.SelectedValue = objPedido.CodigoCliente & "-" & CStr(objPedido.EnderecoCliente)
                    End If
                Else
                    BuscarDepositos()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Function EnabledValorFrete(TipoDeposito As String) As Boolean
        If objPedido.Empresa.Empresa.FretePedido Then
            Return True
        Else
            Return False
        End If
        'If Not Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "04854422" Then Return False
        If TipoDeposito <> "DE" And TipoDeposito <> "LE" Then Return False
        If Not rbFOB.Checked And Not rbCIF.Checked And Not rbTER.Checked And Not rbNenhum.Checked Then Return False
        If objPedido Is Nothing Then SessaoRecuperaPedido()
        If objPedido.SubOperacao Is Nothing Then Return False

        If (objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada And objPedido.FreteCIFFOB = eTiposFrete.FOB) _
            Or
            (objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida And objPedido.FreteCIFFOB = eTiposFrete.CIF) Then
            Return True
        Else
            Return False
        End If
    End Function

    Protected Function ValidaRoteiro() As Boolean
        If objPedido Is Nothing Then
            SessaoRecuperaPedido()
        End If

        If Not rbFOB.Checked And Not rbCIF.Checked And Not rbTER.Checked And Not rbNenhum.Checked Then
            MsgBox(Me.Page, "Marque pelo menos uma opção de Frete (FOB, CIF)")
            Return False
        ElseIf objPedido.SubOperacao Is Nothing Then
            MsgBox(Me.Page, "Indique a SubOperação para o Pedido.")
            Return False
        ElseIf (objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada And objPedido.FreteCIFFOB = eTiposFrete.FOB) _
                        Or
                        (objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida And objPedido.FreteCIFFOB = eTiposFrete.CIF) Then
            Return True
        Else
            MsgBox(Me.Page, "Para Gerar o Roteiro do pedido de Entrada o Frete deve ser (FOB) e para pedido de Saída o Frete deve ser (CIF).")
            Return False
        End If
    End Function

    Protected Sub cmbDepositos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If cmbDepositos.SelectedIndex > 0 Then
                SessaoRecuperaPedido()
                Dim dep() As String = cmbDepositos.SelectedValue.ToString.Split("-")
                If Not VerificarDepositos(dep(0), dep(1)) Then cmbDepositos.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BuscarDepositos()
        ucConsultaClientes.Limpar()

        If String.IsNullOrWhiteSpace(ddlTipoDeposito.SelectedValue) Then
            MsgBox(Me.Page, "Tipo do Depósito não foi selecionado.", eTitulo.Info)
            Exit Sub
        ElseIf ddlTipoDeposito.SelectedValue = "DE" Then
            ucConsultaClientes.SetarTipoCliente(eTipoCliente.Depositos)
        Else
            ucConsultaClientes.SetarTipoCliente(eTipoCliente.Clientes)
        End If

        Popup.ConsultaDeClientes(Me.Page, "objClientePXDEP" & HID.Value)
    End Sub

    Protected Sub btnBuscarDeposito_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        BuscarDepositos()
    End Sub

    Protected Sub GridDepositos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridDepositos.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()

            Dim objDeposito As [Lib].Negocio.PedidoXDeposito = objPedido.Depositos(GridDepositos.SelectedIndex)

            If objDeposito.Tipo = "DE" Then
                ddl.Carregar(cmbDepositos, CarregarDDL.Tabela.Depositos, "", True)
            Else
                cmbDepositos.Items.Clear()
                cmbDepositos.DataTextField = "Descricao"
                cmbDepositos.DataValueField = "Codigo"

                cmbDepositos.Items.Add(New ListItem(Funcoes.AlinharEsquerda(objDeposito.Deposito.Nome, 28, ".") &
                                                    " - " & Funcoes.AlinharEsquerda(objDeposito.Deposito.Cidade, 20, ".") &
                                                    " " & objDeposito.Deposito.CodigoEstado &
                                                    " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(objDeposito.Codigo), 18, ".") &
                                                    "-" & CStr(objDeposito.CodigoEndereco) & "-" & objDeposito.Deposito.Reduzido,
                                                    objDeposito.Codigo & "-" & CStr(objDeposito.CodigoEndereco)))

                Funcoes.InserirLinhaEmBranco(cmbDepositos)
            End If

            ddlTipoDeposito.SelectedValue = objDeposito.Tipo
            chkDepPrincipal.Checked = objDeposito.Principal

            cmbDepositos.SelectedValue = objDeposito.Codigo & "-" & objDeposito.CodigoEndereco
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Private Function ValidaRoteiroDeposito(linha As Integer) As Boolean
        SessaoRecuperaPedido()
        If Not objPedido.Roteiros Is Nothing AndAlso objPedido.Roteiros.Count > 0 Then

            If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                For Each roteiro In objPedido.Roteiros
                    If objPedido.Depositos(linha).Deposito.Equals(roteiro.Destino) Then
                        Return False
                    End If
                Next
            Else

                For Each roteiro As PedidoXRoteiro In objPedido.Roteiros
                    If objPedido.Depositos(linha).Codigo = roteiro.CodigoOrigem AndAlso objPedido.Depositos(linha).CodigoEndereco = roteiro.EnderecoOrigem Then
                        Return False
                    End If
                Next
            End If
        End If
        Return True
    End Function

    Protected Sub imgExcluirDeposito_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim linha As Integer = CType(CType(sender, ImageButton).NamingContainer, GridViewRow).RowIndex

            SessaoRecuperaPedido()

            Dim Tp As String
            Dim msg As String = String.Empty

            Select Case objPedido.Depositos(linha).Tipo
                Case "DE"
                    Tp = "Depósito"
                Case "LE"
                    Tp = "Local de Embarque"
                Case "TR"
                    Tp = "Transbordo"
                Case "PM"
                    Tp = "Proprietário da Mercadoria"
                Case Else
                    Tp = "Origem/Destino"
            End Select

            If Not objPedido.Roteiros Is Nothing AndAlso objPedido.Roteiros.Count > 0 Then
                If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                    If Not objPedido.Roteiros.Where(Function(s) s.CodigoDestino = objPedido.Depositos(linha).Codigo AndAlso s.EnderecoDestino = objPedido.Depositos(linha).CodigoEndereco) Is Nothing Then
                        MsgBox(Me.Page, "Já existe Roteiro definido para o " & Tp & " " & objPedido.Depositos(linha).Deposito.Nome & " - " & objPedido.Depositos(linha).Deposito.Endereco)
                        Exit Sub
                    End If
                Else
                    If Not objPedido.Roteiros.Where(Function(s) s.CodigoOrigem = objPedido.Depositos(linha).Codigo AndAlso s.EnderecoOrigem = objPedido.Depositos(linha).CodigoEndereco) Is Nothing Then
                        MsgBox(Me.Page, "Já existe Roteiro definido para o " & Tp & " " & objPedido.Depositos(linha).Deposito.Nome & " - " & objPedido.Depositos(linha).Deposito.Endereco)
                        Exit Sub
                    End If
                End If
            End If

            If objPedido.TemFaturamentoComDeposito(linha) Then
                MsgBox(Me.Page, "Já existe Faturamento desse pedido e com esse " & Tp & " " & objPedido.Depositos(linha).Deposito.Nome & " - " & objPedido.Depositos(linha).Deposito.Endereco)
                Exit Sub
            End If

            Dim tipoDeposito As String = objPedido.Depositos(linha).Tipo
            Dim ApagouDepositoPrincipal As Boolean = objPedido.Depositos(linha).Principal

            If objPedido.IUD = "I" And objPedido.Depositos(linha).IUD = "I" Then
                objPedido.Depositos.RemoveAt(linha)
            ElseIf objPedido.IUD = "U" Then
                objPedido.Depositos(linha).IUD = "D"
                objPedido.Depositos(linha).Principal = False
            End If

            If ApagouDepositoPrincipal Then
                Dim NovoDepositoPrincipal As PedidoXDeposito = objPedido.Depositos.Where(Function(s) s.Tipo = tipoDeposito And s.IUD <> "D").FirstOrDefault
                If Not NovoDepositoPrincipal Is Nothing Then
                    NovoDepositoPrincipal.Principal = True
                    If NovoDepositoPrincipal.IUD <> "I" Then NovoDepositoPrincipal.IUD = "U"
                End If
            End If

            GridDepositos.DataSource = objPedido.Depositos.Where(Function(s) s.IUD <> "D").ToArray
            GridDepositos.DataBind()

            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function VerificarDepositos(ByVal codDeposito As String, ByVal endDeposito As String) As Boolean
        If ddlTipoDeposito.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione o Tipo do Depósito")
        ElseIf cmbDepositos.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione o Depósito")
        ElseIf ddlTipoDeposito.SelectedValue = "LE" AndAlso objPedido.CodigoEmpresa = codDeposito AndAlso objPedido.EnderecoEmpresa = endDeposito Then
            MsgBox(Me.Page, "Local de Embarque não pode ser e mesmo da Empresa")
            Return False
        ElseIf ddlTipoDeposito.SelectedValue = "LE" AndAlso objPedido.CodigoCliente = codDeposito AndAlso objPedido.EnderecoCliente = endDeposito Then
            MsgBox(Me.Page, "Local de Embarque não pode ser e mesmo do Cliente")
            Return False
        ElseIf ddlTipoDeposito.SelectedValue = "TR" AndAlso objPedido.CodigoEmpresa = codDeposito AndAlso objPedido.EnderecoEmpresa = endDeposito Then
            MsgBox(Me.Page, "Transbordo não pode ser e mesmo da Empresa")
            Return False
        ElseIf ddlTipoDeposito.SelectedValue = "TR" AndAlso objPedido.CodigoCliente = codDeposito AndAlso objPedido.EnderecoCliente = endDeposito Then
            MsgBox(Me.Page, "Transbordo não pode ser e mesmo do Cliente")
            Return False
        End If
        For intRegistro As Integer = objPedido.Depositos.Count - 1 To 0 Step -1
            If objPedido.Depositos(intRegistro).Codigo = codDeposito And objPedido.Depositos(intRegistro).CodigoEndereco = endDeposito And objPedido.Depositos(intRegistro).Tipo = ddlTipoDeposito.SelectedValue Then
                objPedido.Depositos.RemoveAt(intRegistro)
                Exit For
            ElseIf objPedido.Depositos(intRegistro).Tipo = ddlTipoDeposito.SelectedValue AndAlso chkDepPrincipal.Checked = True AndAlso objPedido.Depositos(intRegistro).Principal = chkDepPrincipal.Checked Then
                objPedido.Depositos(intRegistro).Principal = False
            End If
        Next intRegistro
        Return True
    End Function

    Public Function ValidarDepositos() As Boolean
        Dim temDeposito As Boolean = False
        Dim temOriDes As Boolean = False
        Dim temDepPrincipal As Boolean = False
        Dim temODPrincipal As Boolean = False

        For intRegistro As Integer = objPedido.Depositos.Count - 1 To 0 Step -1
            If objPedido.Depositos(intRegistro).Tipo = "DE" Then
                temDeposito = True
                If objPedido.Depositos(intRegistro).Principal Then temDepPrincipal = True
            End If

            If objPedido.Depositos(intRegistro).Tipo = "OD" Then
                temOriDes = True
                If objPedido.Depositos(intRegistro).Principal Then temODPrincipal = True
            End If
        Next intRegistro

        If GridDepositos.Rows.Count = 0 Then
            MsgBox(Me.Page, "Depósitos não foram informados")
            Return False
        ElseIf temDeposito = False Then
            MsgBox(Me.Page, "É obrigatório informar o Depósito")
            Return False
        ElseIf temOriDes = False Then
            MsgBox(Me.Page, "É obrigatório informar a Origem/Destino")
            Return False
        ElseIf temDepPrincipal = False Then
            MsgBox(Me.Page, "Depósito Principal não foi selecionado")
            Return False
        ElseIf temODPrincipal = False Then
            MsgBox(Me.Page, "Origem/Destino Principal não foi selecionado")
            Return False
        Else
            Return True
        End If
    End Function


    '****************** ROTEIROS *****************************

    Protected Sub lnkGerarRoteiro_Click(sender As Object, e As EventArgs) Handles lnkGerarRoteiro.Click
        Try
            If ValidaRoteiro() Then
                CompoeRoteiro()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparRoteiro_Click(sender As Object, e As EventArgs) Handles lnkLimparRoteiro.Click
        Try
            If objPedido.IUD = "I" Then
                objPedido.Roteiros.Clear()
                grdRoteiros.DataSource = objPedido.Roteiros
                grdRoteiros.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        SessaoRecuperaPedido()
    End Sub


    Private Sub CompoeRoteiro()
        SessaoRecuperaPedido()

        'Se tiver TR(Transbordo) adiciona DE ou LE até o TR exceto OD'
        Dim tr = From a In objPedido.Depositos Where a.Tipo = "TR" Select a

        If tr.Count > 0 Then

            Dim DepTipo As String = tr(0).Tipo
            Dim DepCodigo As String = tr(0).Codigo
            Dim DepEndereco As String = tr(0).CodigoEndereco

            Dim objRoteiro As PedidoXRoteiro
            For Each Deposito In objPedido.Depositos

                If Deposito.Tipo <> DepTipo And Deposito.Tipo <> "OD" Then

                    'Se o pedido é de Venda o OD é o Destino'
                    If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso objPedido.Roteiros.Where(Function(t) t.TipoOrigem = Deposito.Tipo And t.CodigoOrigem = Deposito.Codigo And t.EnderecoOrigem = Deposito.CodigoEndereco _
                                                             And t.TipoDestino = DepTipo And t.CodigoDestino = DepCodigo And t.EnderecoDestino = DepEndereco).Count = 0 Then

                        objRoteiro = New PedidoXRoteiro(objPedido)
                        objRoteiro.IUD = "I"

                        objRoteiro.TipoOrigem = Deposito.Tipo
                        objRoteiro.CodigoOrigem = Deposito.Codigo
                        objRoteiro.EnderecoOrigem = Deposito.CodigoEndereco

                        objRoteiro.TipoDestino = DepTipo
                        objRoteiro.CodigoDestino = DepCodigo
                        objRoteiro.EnderecoDestino = DepEndereco

                        objRoteiro.CodigoViaDeTransporte = 1
                        objPedido.Roteiros.Add(objRoteiro)

                        'Se o pedido é de Compra o OD é a Origem'
                    ElseIf objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso objPedido.Roteiros.Where(Function(t) t.TipoOrigem = DepTipo And t.CodigoOrigem = DepCodigo And t.EnderecoOrigem = DepEndereco _
                                                                 And t.TipoDestino = Deposito.Tipo And t.CodigoDestino = Deposito.Codigo And t.EnderecoDestino = Deposito.CodigoEndereco).Count = 0 Then

                        objRoteiro = New PedidoXRoteiro(objPedido)
                        objRoteiro.IUD = "I"

                        objRoteiro.TipoOrigem = DepTipo
                        objRoteiro.CodigoOrigem = DepCodigo
                        objRoteiro.EnderecoOrigem = DepEndereco

                        objRoteiro.TipoDestino = Deposito.Tipo
                        objRoteiro.CodigoDestino = Deposito.Codigo
                        objRoteiro.EnderecoDestino = Deposito.CodigoEndereco

                        objRoteiro.CodigoViaDeTransporte = 1
                        objPedido.Roteiros.Add(objRoteiro)
                    End If

                End If

            Next
        End If

        'adiciona DE ou LE e TR até o OD'
        Dim od = From a In objPedido.Depositos Where a.Tipo = "OD" Select a

        If od.Count > 0 Then

            Dim DepTipo As String = od(0).Tipo
            Dim DepCodigo As String = od(0).Codigo
            Dim DepEndereco As String = od(0).CodigoEndereco

            Dim objRoteiro As PedidoXRoteiro
            For Each Deposito In objPedido.Depositos
                If Deposito.Tipo <> DepTipo Then

                    'Se o pedido é de Venda o OD é o Destino'
                    If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso objPedido.Roteiros.Where(Function(t) t.TipoOrigem = Deposito.Tipo AndAlso t.CodigoOrigem = Deposito.Codigo AndAlso t.EnderecoOrigem = Deposito.CodigoEndereco _
                                                         AndAlso t.TipoDestino = DepTipo AndAlso t.CodigoDestino = DepCodigo AndAlso t.EnderecoDestino = DepEndereco).Count = 0 Then

                        objRoteiro = New PedidoXRoteiro(objPedido)
                        objRoteiro.IUD = "I"

                        objRoteiro.TipoOrigem = Deposito.Tipo
                        objRoteiro.CodigoOrigem = Deposito.Codigo
                        objRoteiro.EnderecoOrigem = Deposito.CodigoEndereco

                        objRoteiro.TipoDestino = DepTipo
                        objRoteiro.CodigoDestino = DepCodigo
                        objRoteiro.EnderecoDestino = DepEndereco
                        objRoteiro.CodigoViaDeTransporte = 1
                        objPedido.Roteiros.Add(objRoteiro)

                        'Se o pedido é de Compra o OD é a Origem'
                    ElseIf objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso objPedido.Roteiros.Where(Function(t) t.TipoOrigem = DepTipo AndAlso t.CodigoOrigem = DepCodigo AndAlso t.EnderecoOrigem = DepEndereco _
                                                             AndAlso t.TipoDestino = Deposito.Tipo AndAlso t.CodigoDestino = Deposito.Codigo AndAlso t.EnderecoDestino = Deposito.CodigoEndereco).Count = 0 Then

                        objRoteiro = New PedidoXRoteiro(objPedido)
                        objRoteiro.IUD = "I"
                        objRoteiro.TipoOrigem = DepTipo
                        objRoteiro.CodigoOrigem = DepCodigo
                        objRoteiro.EnderecoOrigem = DepEndereco

                        objRoteiro.TipoDestino = Deposito.Tipo
                        objRoteiro.CodigoDestino = Deposito.Codigo
                        objRoteiro.EnderecoDestino = Deposito.CodigoEndereco
                        objRoteiro.CodigoViaDeTransporte = 1
                        objPedido.Roteiros.Add(objRoteiro)
                    End If
                End If
            Next
        End If

        If objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
            Dim listaOrdenada As List(Of PedidoXRoteiro) = objPedido.Roteiros
            listaOrdenada.Sort(Function(p1 As PedidoXRoteiro, p2 As PedidoXRoteiro) p1.TipoOrigem.CompareTo(p2.TipoOrigem))

            grdRoteiros.DataSource = listaOrdenada.Where(Function(s) s.IUD <> "D")
            grdRoteiros.DataBind()
        Else
            grdRoteiros.DataSource = objPedido.Roteiros.Where(Function(s) s.IUD <> "D")
            grdRoteiros.DataBind()
        End If


        Dim QuantidadeFiscal As Decimal

        If objPedido.Itens.QuantidadeTotal > 0 Then
            QuantidadeFiscal = objPedido.Itens.QuantidadeTotal
        Else
            QuantidadeFiscal = objPedido.Itens.Sum(Function(s) s.SaldoItem.QtdeContratadoFiscal)
        End If

        txtTotalQtdPedido.Text = QuantidadeTotalPedido.ToString("N2")

        SessaoSalvaPedido()
    End Sub

    Protected Sub ddlTabelaDeViaDeTransporte_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            Dim Cmb As DropDownList = CType(sender, DropDownList)
            Dim linha As Integer = CType(Cmb.NamingContainer, GridViewRow).RowIndex

            If Cmb.SelectedIndex > 0 Then
                SessaoRecuperaPedido()
                objPedido.Roteiros(linha).CodigoViaDeTransporte = Cmb.SelectedValue
                SessaoSalvaPedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdRoteiros_RowCreated(sender As Object, e As GridViewRowEventArgs) Handles grdRoteiros.RowCreated
        Try
            If e.Row.RowType = DataControlRowType.Header Then
                'Customizando Titulo do Grid
                Dim grvObj As GridView = CType(sender, GridView)

                'Criando a nova linha 
                Dim grvObjLinha As GridViewRow = New GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert)

                'Criando nova Celula para linha
                Dim tclCelula As TableCell = New TableCell()

                'Adiciona a Coluna Origem
                tclCelula = New TableCell()
                tclCelula.Text = "Origem"
                tclCelula.ColumnSpan = 3
                grvObjLinha.Cells.Add(tclCelula)
                grvObj.Controls(0).Controls.AddAt(0, grvObjLinha)

                'Adiciona a Coluna Destino
                tclCelula = New TableCell()
                tclCelula.Text = "Destino"
                tclCelula.ColumnSpan = 3
                grvObjLinha.Cells.Add(tclCelula)
                grvObj.Controls(0).Controls.AddAt(0, grvObjLinha)

                'Adiciona a Coluna Outras Informações
                tclCelula = New TableCell()
                'tclCelula.Text = "Outras Informações"
                'tclCelula.ColumnSpan = 6
                grvObjLinha.Cells.Add(tclCelula)
                grvObj.Controls(0).Controls.AddAt(0, grvObjLinha)

                'Adiciona a Coluna Quatidade
                tclCelula = New TableCell()
                tclCelula.Text = "Quantidade"
                tclCelula.ColumnSpan = 2
                grvObjLinha.Cells.Add(tclCelula)
                grvObjLinha.Cells(3).Style.Add("text-align", "Center")
                grvObj.Controls(0).Controls.AddAt(0, grvObjLinha)

                'Adiciona a Coluna Valores
                tclCelula = New TableCell()
                tclCelula.Text = "Valor do Frete"
                tclCelula.ColumnSpan = 2
                grvObjLinha.Cells.Add(tclCelula)
                grvObjLinha.Cells(4).Style.Add("text-align", "Center")
                grvObj.Controls(0).Controls.AddAt(0, grvObjLinha)

                'Adiciona a Coluna Outras Informações
                tclCelula = New TableCell()
                grvObjLinha.Cells.Add(tclCelula)
                grvObj.Controls(0).Controls.AddAt(0, grvObjLinha)

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdRoteiros_DataBound(sender As Object, e As EventArgs) Handles grdRoteiros.DataBound
        divTotalRoteiro.Visible = grdRoteiros.Rows.Count > 0
    End Sub

    Protected Sub txtQuantidadeAtual_TextChanged(sender As Object, e As EventArgs)
        Try
            Dim txt As TextBox = CType(sender, TextBox)
            Dim linha As Integer = CType(txt.NamingContainer, GridViewRow).RowIndex
            Dim txtquantidadeAtual As TextBox = grdRoteiros.Rows(linha).FindControl("txtquantidadeAtual")
            txtRoteiroLog.Text = String.Empty
            SessaoRecuperaPedido()
            If String.IsNullOrWhiteSpace(txtquantidadeAtual.Text) Then
                txtquantidadeAtual.Text = 0
            End If

            If objPedido.IUD = "I" Then
                'Atribui a quantidade atual a quantidade original'
                objPedido.Roteiros(linha).QuantidadeAtual = CDec(txtquantidadeAtual.Text)
                objPedido.Roteiros(linha).QuantidadeOriginal = objPedido.Roteiros(linha).QuantidadeAtual
            Else
                If objPedido.Roteiros(linha).QuantidadeOriginal = 0 Then
                    objPedido.Roteiros(linha).QuantidadeOriginal = objPedido.Roteiros(linha).QuantidadeAtual
                End If
                objPedido.Roteiros(linha).QuantidadeAtual = CDec(txtquantidadeAtual.Text)
                If objPedido.Roteiros(linha).IUD <> "I" Then
                    objPedido.Roteiros(linha).IUD = "U"
                End If
            End If

            'TOTALIZA'
            txtTotalQtdEmbarque.Text = objPedido.Roteiros.TotalQtdEmbarque.ToString("N2")
            txtTotalQtdDestino.Text = objPedido.Roteiros.TotalQtdDestino.ToString("N2")

            If objPedido.Roteiros.TotalQtdEmbarque <> objPedido.Roteiros.TotalQtdDestino Then
                txtTotalQtdEmbarque.ForeColor = Drawing.Color.Red
            Else
                txtTotalQtdEmbarque.ForeColor = Drawing.Color.Black
            End If

            If objPedido.Roteiros.TotalQtdDestino <> objPedido.Itens.QuantidadeTotal Then
                txtTotalQtdDestino.ForeColor = Drawing.Color.Red
            Else
                txtTotalQtdDestino.ForeColor = Drawing.Color.Black
            End If

            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtValor_TextChanged(sender As Object, e As EventArgs)
        Try
            SessaoRecuperaPedido()
            txtRoteiroLog.Text = String.Empty
            divRoteiroLog.Style.Add("display", "none")

            Dim txt As TextBox = CType(sender, TextBox)
            Dim linha As Integer = CType(txt.NamingContainer, GridViewRow).RowIndex

            If objPedido.Roteiros(linha).IUD Is Nothing Then objPedido.Roteiros(linha).IUD = "U"

            If objPedido.IUD = "U" AndAlso objPedido.Roteiros(linha).IUD = "U" AndAlso Not Funcoes.VerificaPermissao("PedidoXRoteiro", "ALTERAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para ALTERAR o valor do Frete Orçado por KG. Processo:PedidoXRoteiro")
            Else

                Dim txtValor As TextBox = grdRoteiros.Rows(linha).FindControl("txtValor")
                Dim txtValorPorTonelada As TextBox = grdRoteiros.Rows(linha).FindControl("txtValorPorTonelada")

                With objPedido.Roteiros(linha)

                    Dim ValorAnterior As Decimal = .Valor
                    .Valor = CDec(txtValor.Text)
                    txtValorPorTonelada.Text = .ValorPorTonelada.ToString("N2")

                    If objPedido.IUD = "U" AndAlso .IUD = "U" Then
                        If .Valor <> ValorAnterior Then
                            .Log &= Now.ToString("dd/MM/yyyy HH:mm:ss") & " - Usuario: " & HttpContext.Current.Session("ssNomeUsuario") & " alterou o valor do frete de " & ValorAnterior & " para " & .Valor & " o KG." & vbCrLf
                        End If
                    End If
                End With
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Tab Transportes"
    Private Sub LimparTransportes()
        hdfTransportadoras.Value = ""
        txtTransportadoras.Text = ""
        txtDataFrete.Text = ""
        txtDataFrete.Enabled = True
        btnTransportadoras.Visible = True
        txtUnitarioFrete.Text = ""
        txtQuantidadeTransp.Text = ""
        txtQuota.Text = ""
        gridTransportes.SelectedIndex = -1
        txtTransportadoras.Focus()
    End Sub

    Public Function ValidarTransportes() As Boolean
        SessaoRecuperaPedido()
        If hdfTransportadoras.Value = "" Then
            MsgBox(Me.Page, "Informe o Transportador!")
            Return False
        ElseIf txtQuota.Text = "" Then
            MsgBox(Me.Page, "Informe a quota para o transporte!")
            Return False
        ElseIf txtQuantidadeTransp.Text = "" Then
            MsgBox(Me.Page, "Informe a quantidade para o transporte!")
            Return False
        ElseIf rbCIF.Checked AndAlso txtDataFrete.Text.Length = 0 Then
            MsgBox(Me.Page, "Informe a Data do Frete!")
            Return False
        ElseIf rbCIF.Checked AndAlso txtUnitarioFrete.Text.Length = 0 Then
            MsgBox(Me.Page, "Informe o Unitário do Frete!")
            Return False
        ElseIf Not Funcoes.VerificaPermissao("PedidosXItens", "ALTERAR") Then
            MsgBox(Me.Page, "Sem Permissao para adicionar Deposito")
            Return False
        ElseIf objPedido.FiscalAberto = False AndAlso Not Funcoes.VerificaAcesso(objPedido.Empresa.Codigo, objPedido.Empresa.CodigoEndereco, IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text)), "PEDIDOS") Then
            MsgBox(Me.Page, "Movimento " & (IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text))) & " já fechado para esta data!")
            Return False
        End If

        Return True
    End Function

    Protected Sub lnkAdicionarTransportador_Click(sender As Object, e As EventArgs) Handles lnkAdicionarTransportador.Click
        Try
            If Not ValidarTransportes() Then Exit Sub
            SessaoRecuperaPedido()
            Dim tra As String() = hdfTransportadoras.Value.Split("-")

            Dim objNovoTransportador As [Lib].Negocio.PedidoXTransportador
            objNovoTransportador = objPedido.Transportadores.Where(Function(s) s.Codigo = tra(0) And s.CodigoEndereco = tra(1) And s.DataFrete = CDate(txtDataFrete.Text)).FirstOrDefault

            If objNovoTransportador Is Nothing Then
                objNovoTransportador = New [Lib].Negocio.PedidoXTransportador(objPedido)
                objNovoTransportador.IUD = "I"
                objNovoTransportador.Codigo = tra(0)
                objNovoTransportador.CodigoEndereco = tra(1)
                If txtDataFrete.Text.Length = 0 Then
                    objNovoTransportador.DataFrete = Now().ToString("yyyy-MM-dd HH:mm:ss")
                Else
                    objNovoTransportador.DataFrete = CDate(txtDataFrete.Text & " " & Now().ToString("HH:mm:ss"))
                End If

                objPedido.Transportadores.Add(objNovoTransportador)
            Else
                If objNovoTransportador.IUD <> "I" Then objNovoTransportador.IUD = "U"
            End If

            objNovoTransportador.UnitarioFrete = IIf(txtUnitarioFrete.Text.Length = 0, 0, txtUnitarioFrete.Text)
            objNovoTransportador.Quantidade = CDec(txtQuantidadeTransp.Text.Replace(".", ""))
            objNovoTransportador.QuotaDiaria = CDec(txtQuota.Text.Replace(".", ""))
            objNovoTransportador.UsuarioInclusao = Session("ssNomeUsuario")

            LimparTransportes()

            gridTransportes.DataSource = objPedido.Transportadores.Where(Function(s) s.IUD <> "D").ToArray
            gridTransportes.DataBind()

            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparTransportador_Click(sender As Object, e As EventArgs) Handles lnkLimparTransportador.Click
        Try
            LimparTransportes()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnTransportadoras_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTransportadoras.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente(eTipoCliente.Transportadores)
            Popup.ConsultaDeClientes(Me.Page, "objTransportadorPXI" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluirTransportador_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgTransportador As ImageButton = CType(sender, ImageButton)
            Dim Linha As Integer = CType(imgTransportador.NamingContainer, GridViewRow).RowIndex

            If Funcoes.VerificaPermissao("PedidosXTransportadores", "EXCLUIR") Then
                SessaoRecuperaPedido()
                If objPedido.FiscalAberto = False AndAlso Not Funcoes.VerificaAcesso(objPedido.Empresa.Codigo, objPedido.Empresa.CodigoEndereco, IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text)), "PEDIDOS") Then
                    MsgBox(Me.Page, "Movimento " & (IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text))) & " já fechado para esta data!")
                    Exit Sub
                End If

                If objPedido.Empresa.Empresa.FretePedido AndAlso objPedido.SubOperacao.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida AndAlso objPedido.FreteCIFFOB = eTiposFrete.CIF AndAlso gridTransportes.Rows.Count = 1 Then
                    MsgBox(Me.Page, "Operação de Saída com Frete por conta da Empresa, não pode ficar sem Transportador")
                    Exit Sub
                ElseIf objPedido.Empresa.Empresa.FretePedido AndAlso objPedido.SubOperacao.EntradaSaida = [Lib].Negocio.eEntradaSaida.Entrada AndAlso objPedido.FreteCIFFOB = eTiposFrete.FOB AndAlso gridTransportes.Rows.Count = 1 Then
                    MsgBox(Me.Page, "Operação de Entrada com Frete por conta da Empresa, não pode ficar sem Transportador")
                    Exit Sub
                End If

                If objPedido.IUD = "I" Or objPedido.Transportadores(Linha).IUD = "I" Then
                    objPedido.Transportadores.RemoveAt(Linha)
                ElseIf objPedido.IUD = "U" Then
                    objPedido.Transportadores(Linha).IUD = "D"
                End If

                SessaoSalvaPedido()

                gridTransportes.DataSource = objPedido.Transportadores.Where(Function(s) s.IUD <> "D").ToArray
                gridTransportes.DataBind()

            Else : MsgBox(Me.Page, "Usuario sem permissão para excluir Valor do Frete do Pedido!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridTransportes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridTransportes.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()

            Dim objTransportador As [Lib].Negocio.PedidoXTransportador = objPedido.Transportadores(gridTransportes.SelectedIndex)
            hdfTransportadoras.Value = objTransportador.Codigo & "-" & objTransportador.CodigoEndereco
            txtTransportadoras.Text = Funcoes.FormatarListItemCliente(objTransportador.Transportador).Text

            txtDataFrete.Text = objTransportador.DataFrete
            txtUnitarioFrete.Text = objTransportador.UnitarioFrete.ToString("N10")
            txtQuantidadeTransp.Text = objTransportador.Quantidade.ToString("N4")
            txtQuota.Text = objTransportador.QuotaDiaria.ToString("N4")

            txtDataFrete.Enabled = False
            btnTransportadoras.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Tab Representantes / Comissões"
    Public Function ValidaComissoes() As Boolean
        If txtPerRepre.Text = "" Then txtPerRepre.Text = 0

        If hdfRepresentante.Value = "" Then
            MsgBox(Me.Page, "Informe o Representante!")
            Return False
        ElseIf chkPercFixoComissao.Checked AndAlso (Not IsNumeric(txtPerRepre.Text) OrElse CDec(txtPerRepre.Text) = 0) Then
            If Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "05272759" Then
                Return True
            ElseIf Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "24450490" Then
                If (Not IsNumeric(txtValorRepre.Text) OrElse CDec(txtValorRepre.Text) = 0) Then
                    MsgBox(Me.Page, "Informe o %  ou o valor!")
                    Return False
                Else
                    Return True
                End If
            Else
                MsgBox(Me.Page, "Informe o %  para obter o valor!")
                Return False
            End If
        Else
            Return True
        End If
    End Function

    Private Sub LimparComissoes()
        hdfRepresentante.Value = ""
        txtRepresentante.Text = ""
        txtPerRepre.Text = ""
        txtValorRepre.Text = ""

        chkPercFixoComissao.Checked = True
        txtPerRepre.Enabled = True
        txtValorRepre.Enabled = True

        divComissoes.Visible = False
        lblTabeladeComissao.Text = "Tabela de Comissão"
        gridTabelaDeComissao.DataSource = Nothing
        gridTabelaDeComissao.DataBind()
        txtRepresentante.Focus()
    End Sub

    Protected Sub btnRepresentante_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente(eTipoCliente.Representantes)
            Popup.ConsultaDeClientes(Me.Page, "objRepresentantePXI" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkPercFixoComissao_CheckedChanged(sender As Object, e As EventArgs) Handles chkPercFixoComissao.CheckedChanged
        Try
            If chkPercFixoComissao.Checked Then
                txtPerRepre.Enabled = True
                txtValorRepre.Text = 0.0
                txtValorRepre.Enabled = False
                txtPerRepre.Focus()
            Else
                txtPerRepre.Text = 0
                txtPerRepre.Enabled = False
                txtValorRepre.Enabled = True
                txtValorRepre.Focus()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtPerRepre_TextChanged(sender As Object, e As EventArgs) Handles txtPerRepre.TextChanged
        Try
            If IsNumeric(txtPerRepre.Text) Then
                SessaoRecuperaPedido()
                txtValorRepre.Text = Math.Round((objPedido.Itens.TotalOficial * CDec(txtPerRepre.Text)) / 100, 2).ToString("N2")
            Else
                txtPerRepre.Text = "0"
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAdicionarRepresentante_Click(sender As Object, e As EventArgs) Handles lnkAdicionarRepresentante.Click
        Try
            SessaoRecuperaPedido()

            If objPedido.FiscalAberto = False AndAlso Not Funcoes.VerificaAcesso(objPedido.Empresa.Codigo, objPedido.Empresa.CodigoEndereco, IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text)), "PEDIDOS") Then
                MsgBox(Me.Page, "Movimento " & (IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text))) & " já fechado para esta data!")
                Exit Sub
            End If

            If ValidaComissoes() Then
                Dim strRepresentante As String() = hdfRepresentante.Value.Split("-")

                Dim objRepresentante As PedidoXRepresentante

                'Caso seja o primeiro registro a ser inserido marca a opção principal
                If objPedido.Representantes.Count = 0 Then
                    chkRepPrincipal.Checked = True
                End If

                objRepresentante = objPedido.Representantes.Where(Function(s) s.CodigoRepresentante = strRepresentante(0) And s.CodigoEnderecoRepresentante = strRepresentante(1)).FirstOrDefault

                If objRepresentante Is Nothing Then
                    If Not Funcoes.VerificaPermissao("PedidosXItens", "GRAVAR") Then
                        MsgBox(Me.Page, "Usuario sem Permissao para Incluir!")
                        Exit Sub
                    End If

                    objRepresentante = New [Lib].Negocio.PedidoXRepresentante(objPedido)
                    objRepresentante.IUD = "I"
                    objRepresentante.CodigoRepresentante = strRepresentante(0)
                    objRepresentante.CodigoEnderecoRepresentante = strRepresentante(1)

                    objPedido.Representantes.Add(objRepresentante)
                Else
                    If Not Funcoes.VerificaPermissao("PedidosXItens", "ALTERAR") Then
                        MsgBox(Me.Page, "Usuario sem Permissao para Alterar!")
                        Exit Sub
                    End If

                    If objRepresentante.IUD <> "I" Then
                        objRepresentante.IUD = "U"
                    End If
                End If

                If chkPercFixoComissao.Checked Then
                    If objRepresentante.Comissoes.Count > 0 Then
                        For Each comissoes In objRepresentante.Comissoes
                            comissoes.IUD = "D"
                        Next
                    End If

                    objRepresentante.Percentual = CDec(txtPerRepre.Text)
                    objRepresentante.ValorComissao = Math.Round((objPedido.Itens.TotalOficial * objRepresentante.Percentual) / 100, 2)
                Else
                    'objRepresentante.Percentual = 0
                    'objRepresentante.ValorComissao = CDec(txtValorRepre.Text)
                    'objRepresentante.Comissoes.AtualizarLista()
                    'gridTabelaDeComissao.DataSource = objRepresentante.Comissoes.ToArray
                    'gridTabelaDeComissao.DataBind()
                    If objRepresentante.Comissoes.Count > 0 Then
                        For Each comissoes In objRepresentante.Comissoes
                            comissoes.IUD = "D"
                        Next
                    End If

                    objRepresentante.Percentual = 0
                    objRepresentante.ValorComissao = CDec(txtValorRepre.Text)
                End If

                objRepresentante.Principal = chkRepPrincipal.Checked

                'Verifica se têm algum representante marcado como principal e caso não tenha marcará o primeiro da lista'
                Dim RepresentantePrincipal As PedidoXRepresentante = objPedido.Representantes.Where(Function(s) s.IUD <> "D" And s.Principal = True).FirstOrDefault
                If RepresentantePrincipal Is Nothing Then
                    objPedido.Representantes.Item(0).Principal = True
                End If

                objRepresentante.PercentualFixo = chkPercFixoComissao.Checked

                If chkRepPrincipal.Checked Then
                    Dim RepAntigoPrincipal As PedidoXRepresentante = objPedido.Representantes.Where(Function(s) s.IUD <> "D" And s.Principal = True And (s.CodigoRepresentante <> objRepresentante.CodigoRepresentante Or s.CodigoEnderecoRepresentante <> objRepresentante.CodigoEnderecoRepresentante)).FirstOrDefault

                    If Not RepAntigoPrincipal Is Nothing Then
                        If RepAntigoPrincipal.IUD <> "I" Then RepAntigoPrincipal.IUD = "U"
                        RepAntigoPrincipal.Principal = False
                    End If
                End If

                LimparComissoes()

                chkRepPrincipal.Checked = False

                GridRepresentantes.DataSource = objPedido.Representantes.Where(Function(s) s.IUD <> "D").ToArray
                GridRepresentantes.DataBind()
            End If

            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparRepresentante_Click(sender As Object, e As EventArgs) Handles lnkLimparRepresentante.Click
        Try
            txtRepresentante.Text = String.Empty
            txtPerRepre.Text = String.Empty
            txtValorRepre.Text = String.Empty
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluirRepresentante_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim linha As Integer = CType(CType(sender, ImageButton).NamingContainer, GridViewRow).RowIndex

            SessaoRecuperaPedido()

            If objPedido.FiscalAberto = False AndAlso Not Funcoes.VerificaAcesso(objPedido.Empresa.Codigo, objPedido.Empresa.CodigoEndereco, IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text)), "PEDIDOS") Then
                MsgBox(Me.Page, "Movimento " & (IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text))) & " já fechado para esta data!")
                Exit Sub
            End If

            If objPedido.IUD = "I" Or objPedido.Representantes(linha).IUD = "I" Then
                objPedido.Representantes.RemoveAt(linha)
            Else
                objPedido.Representantes(linha).IUD = "D"
                For Each comissoes In objPedido.Representantes(linha).Comissoes
                    comissoes.IUD = "D"
                Next
            End If
            LimparComissoes()

            GridRepresentantes.DataSource = objPedido.Representantes.Where(Function(s) s.IUD <> "D").ToArray
            GridRepresentantes.DataBind()

            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridComissoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridRepresentantes.SelectedIndexChanged, GridRepresentantes.SelectedIndexChanged, GridRepresentantes.SelectedIndexChanged, GridRepresentantes.SelectedIndexChanged, GridRepresentantes.SelectedIndexChanged, GridRepresentantes.SelectedIndexChanged, GridRepresentantes.SelectedIndexChanged, GridRepresentantes.SelectedIndexChanged, GridRepresentantes.SelectedIndexChanged, GridRepresentantes.SelectedIndexChanged, GridRepresentantes.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()

            Dim objRepresentante As [Lib].Negocio.PedidoXRepresentante = objPedido.Representantes(GridRepresentantes.SelectedIndex)
            HidLinhaRepresentante.Value = GridRepresentantes.SelectedIndex
            hdfRepresentante.Value = objRepresentante.CodigoRepresentante & "-" & objRepresentante.CodigoEnderecoRepresentante

            Funcoes.FormatarClienteTXT(txtRepresentante, objRepresentante.Representante)

            chkRepPrincipal.Checked = objRepresentante.Principal
            txtPerRepre.Text = objRepresentante.Percentual.ToString("N10")
            txtValorRepre.Text = objRepresentante.ValorComissao.ToString("N2")

            chkPercFixoComissao.Checked = objRepresentante.PercentualFixo
            If objRepresentante.PercentualFixo Then
                txtPerRepre.Enabled = True
                txtValorRepre.Enabled = True

                divComissoes.Visible = False
                lblTabeladeComissao.Text = "Tabela de Comissão"
                gridTabelaDeComissao.DataSource = Nothing
                gridTabelaDeComissao.DataBind()
            Else
                txtPerRepre.Text = 0
                txtPerRepre.Enabled = False
                txtValorRepre.Text = 0.0
                txtValorRepre.Enabled = False

                divComissoes.Visible = True
                lblTabeladeComissao.Text = "Tabela de Comissão - " & objRepresentante.Representante.Nome
                gridTabelaDeComissao.DataSource = objRepresentante.Comissoes.Where(Function(s) s.IUD <> "D")
                gridTabelaDeComissao.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlTabelaDeComissao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim Cmb As DropDownList = CType(sender, DropDownList)
            Dim linha As Integer = CType(Cmb.NamingContainer, GridViewRow).RowIndex

            If Cmb.SelectedIndex > 0 Then
                SessaoRecuperaPedido()
                objPedido.Representantes(HidLinhaRepresentante.Value).Comissoes.Where(Function(s) s.IUD <> "D")(linha).CodigoTabela = Cmb.SelectedValue
                SessaoSalvaPedido()
                cmbSafra.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Tab Vencimentos"
    Private Sub LimparCondicoes(ByVal LimparGrade As Boolean)
        cmbCondicoes.SelectedIndex = 0
        txtValorVencimento.Text = ""
        txtDataVencimento.Text = ""

        If LimparGrade Then
            GridCondicoes.SelectedIndex = -1
            GridCondicoes.DataSource = Nothing
            GridCondicoes.DataBind()
        End If

        txtTotalPacelado.Text = "0,00"
        txtTotalPago.Text = "0,00"
        txtSaldoVencimentos.Text = "0,00"
    End Sub

    Private Sub CalcularParcelamento()
        Try

            If (objPedido.Moeda.Classificacao = eTiposMoeda.Oficial AndAlso objPedido.Itens.TotalOficial = 0) Or (objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso objPedido.Itens.TotalMoeda = 0) Then
                For i As Integer = objPedido.Vencimentos.Count - 1 To 0 Step -1
                    If Not objPedido.Vencimentos(i).Provisao = [Lib].Negocio.eProvisao.Baixa Then
                        objPedido.Vencimentos.Remove(objPedido.Vencimentos(i))
                    End If
                Next i
            Else
                objPedido.Vencimentos.CriarParcelamento(False, CType(Session("objPedVencimentos" & HID.Value), Hashtable))
            End If

            GridCondicoes.DataSource = objPedido.Vencimentos.ToDataTablePedidos()
            GridCondicoes.DataBind()

            objPedido.Vencimentos.TotalDasParcelas()

            txtTotalPacelado.Text = (objPedido.Vencimentos.OficialAMovimentar - objPedido.Vencimentos.OficialADevolver).ToString("N2")
            txtTotalPago.Text = (objPedido.Vencimentos.OficialMovimentado - objPedido.Vencimentos.OficialDevolvido).ToString("N2")
            txtSaldoVencimentos.Text = (objPedido.Vencimentos.OficialAMovimentar - objPedido.Vencimentos.OficialADevolver - (objPedido.Vencimentos.OficialMovimentado - objPedido.Vencimentos.OficialDevolvido)).ToString("N2")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Private Sub CalcularParcelamento(ByVal Manter As Boolean)
        Try
            If (objPedido.Moeda.Classificacao = eTiposMoeda.Oficial AndAlso objPedido.Itens.TotalOficial = 0) Or (objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso objPedido.Itens.TotalMoeda = 0) Then
                For i As Integer = objPedido.Vencimentos.Count - 1 To 0 Step -1
                    If Not objPedido.Vencimentos(i).Provisao = [Lib].Negocio.eProvisao.Baixa Then
                        objPedido.Vencimentos.Remove(objPedido.Vencimentos(i))
                    End If
                Next i
            Else
                objPedido.Vencimentos.CriarVariasProvisao(False, CType(Session("objPedVencimentos" & HID.Value), Hashtable))
            End If

            Dim parcelas As New PedidoXParcelas(objPedido)
            parcelas.Clear()
            For Each p In objPedido.Vencimentos.Where(Function(x) x.Situacao = eSituacao.Normal Or x.Situacao = eSituacao.RemessaBancaria Or x.Situacao = eSituacao.EndossoTitulo)
                parcelas.Add(p)
            Next

            GridCondicoes.DataSource = parcelas.ToDataTablePedidos()
            GridCondicoes.DataBind()

            objPedido.Vencimentos.TotalDasParcelas()

            txtTotalPacelado.Text = (objPedido.Vencimentos.OficialAMovimentar - objPedido.Vencimentos.OficialADevolver).ToString("N2")
            txtTotalPago.Text = (objPedido.Vencimentos.OficialMovimentado - objPedido.Vencimentos.OficialDevolvido).ToString("N2")
            txtSaldoVencimentos.Text = (objPedido.Vencimentos.OficialAMovimentar - objPedido.Vencimentos.OficialADevolver - (objPedido.Vencimentos.OficialMovimentado - objPedido.Vencimentos.OficialDevolvido)).ToString("N2")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbCondicoes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbCondicoes.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()
            If cmbCondicoes.SelectedIndex > 0 Then

                btnOkVencimento.Enabled = True
                txtValorVencimento.Enabled = True
                txtDataVencimento.Enabled = True

                Dim Sql As String = "Select isnull(AVista,0) as Avista " & vbCrLf &
                                     "from Pagamentos " & vbCrLf &
                                     "where Pagamento_id = " & cmbCondicoes.SelectedValue

                Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Pagamentos")

                If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                    Dim temAVista As Boolean

                    For Each row As DataRow In ds.Tables(0).Rows
                        If CBool(row("Avista")) AndAlso objPedido.Cliente.ApenasAVista Then
                            btnOkVencimento.Enabled = False
                            txtValorVencimento.Enabled = False
                            txtDataVencimento.Enabled = False

                            temAVista = False
                        ElseIf objPedido.Cliente.ApenasAVista Then
                            temAVista = True
                        Else
                            temAVista = False
                        End If
                    Next

                    If temAVista Then
                        cmbCondicoes.SelectedIndex = 0
                        MsgBox(Me.Page, "Condição para esse cliente é apenas à Vista.", eTitulo.Info)
                        Exit Sub
                    End If
                End If

                objPedido.Vencimentos = New PedidoXParcelas(objPedido)
                objPedido.CodigoCondicaoPagamento = cmbCondicoes.SelectedValue
                CalcularParcelamento(False)
            Else
                objPedido.CodigoCondicaoPagamento = 0
                LimparCondicoes(True)
            End If
            txtValorVencimento.Enabled = True
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnOkVencimento_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOkVencimento.Click
        Try
            SessaoRecuperaPedido()
            If GridCondicoes.SelectedIndex > -1 AndAlso objPedido.Vencimentos(GridCondicoes.SelectedIndex).Provisao = [Lib].Negocio.eProvisao.Baixa Then
                MsgBox(Me.Page, "Esta parcela já esta Baixada, não pode ser alterada pelo Pedido.")
                'Else If Original 26-11-2014 Edson problema FEX
                'ElseIf GridCondicoes.SelectedIndex > -1 AndAlso objPedido.Vencimentos(GridCondicoes.SelectedIndex).Provisao = [Lib].Negocio.eProvisao.Previsao Then
            ElseIf GridCondicoes.SelectedIndex > -1 AndAlso objPedido.Vencimentos(GridCondicoes.SelectedIndex).Provisao <> eProvisao.Baixa Then

                If Not IsDate(txtDataVencimento.Text) OrElse Convert.ToDateTime(txtDataVencimento.Text) < Convert.ToDateTime(txtDataPedido.Text) Then
                    MsgBox(Me.Page, "Data de vencimento não pode ser menor que " & txtDataPedido.Text)
                    Exit Sub
                End If

                If objPedido.IUD = "I" Then
                    objPedido.Vencimentos(GridCondicoes.SelectedIndex).Vencimento = Funcoes.ValidaDataUtil(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, CDate(txtDataVencimento.Text))
                    objPedido.Vencimentos(GridCondicoes.SelectedIndex).DataProrrogacao = objPedido.Vencimentos(GridCondicoes.SelectedIndex).Vencimento
                Else
                    objPedido.Vencimentos(GridCondicoes.SelectedIndex).DataProrrogacao = Funcoes.ValidaDataUtil(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, CDate(txtDataVencimento.Text))
                End If

                'Dim strData As String() = txtDataVencimento.Text.Split("/")

                Dim msg As String = String.Empty

                If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.MoedaEstrangeira Then
                    'objPedido.Vencimentos.ModificarParcelaDolar(GridCondicoes.SelectedIndex, _
                    '                                       New DateTime(Convert.ToInt32(strData(2)), Convert.ToInt32(strData(1)), Convert.ToInt32(strData(0))), _
                    '                                       Convert.ToDecimal(txtValorVencimento.Text))

                    If Not objPedido.Vencimentos(GridCondicoes.SelectedIndex).ValorDocumentoMoeda = CDec(txtValorVencimento.Text) Then msg = objPedido.Vencimentos.AjustaParcelas(GridCondicoes.SelectedIndex, objPedido.Vencimentos(GridCondicoes.SelectedIndex).ValorDocumentoMoeda, CDec(txtValorVencimento.Text))
                Else
                    'objPedido.Vencimentos.ModificarParcela(GridCondicoes.SelectedIndex, _
                    '                                       New DateTime(Convert.ToInt32(strData(2)), Convert.ToInt32(strData(1)), Convert.ToInt32(strData(0))), _
                    '                                       Convert.ToDecimal(txtValorVencimento.Text))
                    If Not objPedido.Vencimentos(GridCondicoes.SelectedIndex).ValorDocumentoOficial = CDec(txtValorVencimento.Text) Then msg = objPedido.Vencimentos.AjustaParcelas(GridCondicoes.SelectedIndex, objPedido.Vencimentos(GridCondicoes.SelectedIndex).ValorDocumentoOficial, CDec(txtValorVencimento.Text))
                End If

                If String.IsNullOrWhiteSpace(msg) Then
                    GridCondicoes.DataSource = objPedido.Vencimentos.ToDataTablePedidos()
                    GridCondicoes.DataBind()

                    'txtTotalPacelado.Text = objPedido.Vencimentos.OficialAMovimentar.ToString("N2")
                    'txtTotalPago.Text = objPedido.Vencimentos.OficialMovimentado.ToString("N2")
                    'txtSaldoVencimentos.Text = (objPedido.Vencimentos.OficialAMovimentar - objPedido.Vencimentos.OficialMovimentado).ToString("N2")

                    'txtTotalPacelado.Text = objPedido.Vencimentos.OficialAMovimentar.ToString("N2")
                    txtTotalPacelado.Text = (objPedido.Vencimentos.OficialAMovimentar - objPedido.Vencimentos.OficialADevolver).ToString("N2")
                    txtTotalPago.Text = (objPedido.Vencimentos.OficialMovimentado - objPedido.Vencimentos.OficialDevolvido).ToString("N2")
                    'txtSaldoVencimentos.Text = (objPedido.Vencimentos.OficialAMovimentar - (objPedido.Vencimentos.OficialMovimentado - objPedido.Vencimentos.OficialDevolvido)).ToString("N2")
                    txtSaldoVencimentos.Text = (objPedido.Vencimentos.OficialAMovimentar - objPedido.Vencimentos.OficialADevolver - (objPedido.Vencimentos.OficialMovimentado - objPedido.Vencimentos.OficialDevolvido)).ToString("N2")

                    txtDataVencimento.Text = ""
                    txtValorVencimento.Text = ""

                    SessaoSalvaPedido()
                Else
                    MsgBox(Me.Page, msg)
                End If
            Else : MsgBox(Me.Page, "Nenhuma parcela foi selecionada para alteração!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridCondicoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridCondicoes.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()

            Dim Sql As String = "Select isnull(AVista,0) as Avista " & vbCrLf &
                                     "from Pagamentos " & vbCrLf &
                                     "where Pagamento_id = " & objPedido.CodigoCondicaoPagamento

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Pagamentos")

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                Dim temAVista As Boolean = False

                For Each row As DataRow In ds.Tables(0).Rows
                    If CBool(row("Avista")) AndAlso objPedido.Cliente.ApenasAVista Then
                        temAVista = True
                    End If
                Next

                If temAVista Then
                    btnOkVencimento.Enabled = False
                    txtValorVencimento.Enabled = False
                    txtDataVencimento.Enabled = False
                    MsgBox(Me.Page, "Condição para esse cliente é apenas à Vista e vencimento não pode ser modificado.", eTitulo.Info)
                    Exit Sub
                Else
                    btnOkVencimento.Enabled = True
                    txtValorVencimento.Enabled = True
                    txtDataVencimento.Enabled = True
                End If
            End If

            txtDataVencimento.Text = objPedido.Vencimentos(GridCondicoes.SelectedIndex).DataProrrogacao.ToString("dd/MM/yyyy")
            If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.MoedaEstrangeira Then
                txtValorVencimento.Text = objPedido.Vencimentos(GridCondicoes.SelectedIndex).ValorLiquidoMoeda.ToString("N2")
            Else
                txtValorVencimento.Text = objPedido.Vencimentos(GridCondicoes.SelectedIndex).ValorLiquidoOficial.ToString("N2")
            End If

            If GridCondicoes.Rows.Count = 1 Or GridCondicoes.SelectedIndex = GridCondicoes.Rows.Count Then
                txtValorVencimento.Enabled = False
            ElseIf Not String.IsNullOrWhiteSpace(objPedido.Vencimentos(GridCondicoes.SelectedIndex).ContratoBancario) Then
                txtValorVencimento.Enabled = False
                MsgBox(Me.Page, "Valor da Parcela com Contrato Bancário não pode ser alterado, qualquer dúvida entre em contato com o Financeiro.")
            Else
                If objPedido.Vencimentos.All(Function(s) s.Provisao = eProvisao.Provisao) Or objPedido.Vencimentos.All(Function(s) s.Provisao = eProvisao.Previsao) Then
                    txtValorVencimento.Enabled = True
                Else
                    txtValorVencimento.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridCondicoes_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridCondicoes.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                e.Row.Cells(7).BackColor = Drawing.Color.LightSkyBlue
                e.Row.Cells(8).BackColor = Drawing.Color.LightSkyBlue
                e.Row.Cells(9).BackColor = Drawing.Color.LightSkyBlue
                e.Row.Cells(10).BackColor = Drawing.Color.LightSkyBlue
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlMomentoFinanceiro_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlMomentoFinanceiro.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()
            If objPedido Is Nothing Then Exit Sub

            If objPedido.Troca Then
                objPedido.MomentoFinanceiro = 4
                ddlMomentoFinanceiro.SelectedIndex = 3
            ElseIf (objPedido.SubOperacao.Financeiro Or Not objPedido.SubOperacao.PrecoFixo) And ddlMomentoFinanceiro.SelectedIndex <> 0 Then
                objPedido.MomentoFinanceiro = ddlMomentoFinanceiro.SelectedValue
            Else
                ddlMomentoFinanceiro.SelectedValue = objPedido.MomentoFinanceiro
            End If

            If Left(objPedido.CodigoEmpresa, 8) = "04440724" AndAlso objPedido.SubOperacao.Financeiro AndAlso objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                divEntrega.Visible = True
            Else
                divEntrega.Visible = False
            End If

            objPedido.Vencimentos.Clear()

            '26-11-2014 Edson resolver problema implantacao FEX
            objPedido.CodigoCondicaoPagamento = 0
            LimparCondicoes(True)

            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lstCondicoesPgtoEntrega_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstCondicoesPgtoEntrega.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()
            If lstCondicoesPgtoEntrega.SelectedIndex > 0 Then
                objPedido.CodigoCondicaoPagamentoDaEntrega = lstCondicoesPgtoEntrega.SelectedValue
                SessaoSalvaPedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.ToString))
        End Try
    End Sub

    Protected Sub txtQuotaDeEntrega_TextChanged(sender As Object, e As EventArgs) Handles txtQuotaDeEntrega.TextChanged
        Try
            If Not String.IsNullOrWhiteSpace(txtQuotaDeEntrega.Text) Then
                SessaoRecuperaPedido()
                objPedido.QuotaEntrega = txtQuotaDeEntrega.Text
                SessaoSalvaPedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlPeriodicidadeEntrega_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlPeriodicidadeEntrega.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()
            objPedido.PeriodicidadeEntrega = ddlPeriodicidadeEntrega.SelectedValue
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Tab Observacao"
    Protected Sub lnkAdicionarObservacao_Click(sender As Object, e As EventArgs) Handles lnkAdicionarObservacao.Click
        Try
            SessaoRecuperaPedido()
            objPedido.Observacoes = Date.Now.ToString("dd/MM/yyyy HH:mm:ss") & " - Usuario: " & HttpContext.Current.Session("ssNomeUsuario") & vbCrLf & "       " & Funcoes.EliminarCaracteresEspeciais(txtAddObservacao.Text) & IIf(objPedido.Observacoes.Length > 0, vbCrLf & objPedido.Observacoes, "")
            txtObservacao.Text = objPedido.Observacoes
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub AtualizarObservacao()
        SessaoRecuperaPedido()
        txtObservacao.Text = objPedido.Observacoes
        SessaoSalvaPedido()
    End Sub
#End Region

#Region "Tab Pedido Troca"
    Protected Sub ddlContaAdto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlContaAdto.SelectedIndex > 0 Then
                SessaoRecuperaPedido()
                objPedido.ContaAdiantamentoTroca = ddlContaAdto.SelectedValue
                SessaoSalvaPedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub LnkVincular_Click(sender As Object, e As EventArgs) Handles LnkVincular.Click
        Try

            SessaoRecuperaPedido()
            If objPedido.Troca AndAlso objPedido.Operacao IsNot Nothing AndAlso objPedido.Operacao.CodigoClasse = eClassesOperacoes.COMPRAS.ToString Then
                If String.IsNullOrWhiteSpace(objPedido.CodigoCliente) OrElse String.IsNullOrWhiteSpace(objPedido.CodigoSafra) Then
                    MsgBox(Me.Page, "Informe cliente e safra antes de selecionar o pedido de troca!")
                    Exit Sub
                End If
                Dim parameters As New Dictionary(Of String, Object)
                parameters.Add("Empresa", "")
                parameters.Add("EndEmpresa", "")
                parameters.Add("Cliente", objPedido.CodigoCliente)
                parameters.Add("EndCliente", objPedido.EnderecoCliente)
                parameters.Add("Safra", objPedido.CodigoSafra)
                parameters.Add("Moeda", objPedido.CodigoMoeda)
                parameters.Add("Where", "")
                Popup.ConsultaPedidoDeTroca(Me.Page, "objPedidoTroca" & HID.Value)
                ucConsultaPedidoDeTroca.BindGridView(parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub LnkDesvincular_Click(sender As Object, e As EventArgs) Handles LnkDesvincular.Click
        Try
            SessaoRecuperaPedido()
            If objPedido.FiscalAberto = False AndAlso Not Funcoes.VerificaAcesso(objPedido.Empresa.Codigo, objPedido.Empresa.CodigoEndereco, IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text)), "PEDIDOS") Then
                MsgBox(Me.Page, "Movimento " & (IIf(CDate(txtDataVencimentoPedido.Text) > Now.Date, Now.Date, CDate(txtDataVencimentoPedido.Text))) & " já fechado para esta data!")
                Exit Sub
            End If

            If objPedido.Troca AndAlso objPedido.Operacao IsNot Nothing AndAlso objPedido.Operacao.CodigoClasse = eClassesOperacoes.COMPRAS.ToString Then
                If objPedido.TemFinanceiro(1) Then
                    MsgBox(Me.Page, "Pedido com Financeiro Baixado não pode ser desvinculado.")
                    Exit Sub
                End If

                Dim CodigoPedidoTroca As Integer = objPedido.CodigoPedidoTroca

                objPedido.CodigoEmpresaTroca = ""
                objPedido.EnderecoEmpresaTroca = 0
                objPedido.CodigoPedidoTroca = 0
                objPedido.PedidoTroca = Nothing

                chkTroca.Checked = False
                chkTroca_CheckedChanged(Nothing, Nothing)

                lblEmpresaTroca.Text = ""
                lblClienteTroca.Text = ""
                lblPedidoTroca.Text = ""
                lblCNTroca.Text = ""
                LnkVincular.Parent.Visible = True
                chkTroca.Enabled = True
                LnkDesvincular.Parent.Visible = False
                objPedido.PedidoTroca = New [Lib].Negocio.Pedido()
                objPedido.CodigoEmpresaTroca = ""
                objPedido.EnderecoEmpresaTroca = 0
                objPedido.CodigoPedidoTroca = 0
                SessaoSalvaPedido()
                txtCompraTroca.Text = objPedido.Itens.LiquidoOficial

                MsgBox(Me.Page, "Troca desvinculada, Atualize o Financeiro e regrave o pedido para confirmar a Alteracao. Lembre-se de retirar a opção troca do pedido " & CodigoPedidoTroca, eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Tab Parcelas/Financeiro"
    Protected Sub LnkParcelar_Click(sender As Object, e As EventArgs) Handles LnkParcelar.Click
        SessaoRecuperaPedido()
        If Not IsDate(txtDataCondPagParcela.Text) Then
            MsgBox(Me.Page, "Informe a Data inicial do Parcelamento")
            Exit Sub
        ElseIf CDate(txtDataCondPagParcela.Text).ToString("yyyy/MM/dd") < objPedido.DataPedido.ToString("yyyy/MM/dd") Then
            MsgBox(Me.Page, "Data do inicio do parcelamento nao pode ser menor que a data do Pedido")
            Exit Sub
        End If

        objPedido.Parcelas.Parcelar(CDate(txtDataCondPagParcela.Text))
        gridParcelas.DataSource = objPedido.Parcelas
        gridParcelas.DataBind()
    End Sub

    Protected Sub ddlCondPagPed_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlCondPagPed.SelectedIndexChanged
        SessaoRecuperaPedido()
        objPedido.CodigoCondicaoPagamento = ddlCondPagPed.SelectedValue
        SessaoSalvaPedido()
    End Sub

    Protected Sub gridParcelas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridParcelas.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()

            txtCodigoParcela.Text = objPedido.Parcelas(gridParcelas.SelectedIndex).CodigoParcela
            txtDataVencParcela.Text = objPedido.Parcelas(gridParcelas.SelectedIndex).DataVencimento.ToString("dd/MM/yyyy")
            txtValorParcela.Text = objPedido.Parcelas(gridParcelas.SelectedIndex).Valor.ToString("N2")

            If gridParcelas.Rows.Count = 1 Or gridParcelas.SelectedIndex + 1 = gridParcelas.Rows.Count Then
                txtValorParcela.Enabled = False
            Else
                txtValorParcela.Enabled = True
            End If

            If Not String.IsNullOrWhiteSpace(objPedido.Contrato) Then
                txtValorParcela.Enabled = False
                MsgBox(Me.Page, "Valor da Parcela com Contrato Bancário não pode ser alterado, qualquer dúvida entre em contato com o Financeiro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub LnkAtualizarParcela_Click(sender As Object, e As EventArgs) Handles LnkAtualizarParcela.Click
        If gridParcelas.Rows.Count = 0 Then Exit Sub
        Dim nPar As Integer = CInt(txtCodigoParcela.Text)
        If Not IsNumeric(txtValorParcela.Text) Then
            MsgBox(Me.Page, "Informe um valor para a Parcela.")
            Exit Sub
        End If

        SessaoRecuperaPedido()
        If objPedido.Parcelas.Count > nPar Then
            If objPedido.Parcelas(nPar - 1).DataVencimento > objPedido.Parcelas(nPar).DataVencimento Then
                MsgBox(Me.Page, "A data da Parcela atual nao pode ser maior que a data da parcela subsequente.")
                Exit Sub
            End If
        End If

        objPedido.Parcelas.ModificarParcela(nPar, CDate(txtDataVencParcela.Text), CDec(txtValorParcela.Text))

        If objPedido.Parcelas.MSG.Length > 0 Then
            MsgBox(Me.Page, objPedido.Parcelas.MSG)
            objPedido.Parcelas.MSG = ""
            Exit Sub
        End If

        SessaoSalvaPedido()
        gridParcelas.DataSource = objPedido.Parcelas
        gridParcelas.DataBind()
    End Sub

    Protected Sub GridOrigemFinanceiro_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GridOrigemFinanceiro.SelectedIndexChanged
        SessaoRecuperaPedido()
        gridFinanceiro.DataSource = objPedido.Financeiro.Where(Function(s) s.Origem = GridOrigemFinanceiro.SelectedRow.Cells(1).Text).ToList
        gridFinanceiro.DataBind()
    End Sub
#End Region

    Protected Sub chkTemVariacao_CheckedChanged(sender As Object, e As EventArgs) Handles chkTemVariacao.CheckedChanged
        Try
            SessaoRecuperaPedido()
            objPedido.TemVariacao = chkTemVariacao.Checked
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkFixacoes_Click(sender As Object, e As EventArgs)
        Try
            Dim linha As Integer = CType(CType(sender, LinkButton).NamingContainer, GridViewRow).RowIndex
            ucPedidoxFixacao.Limpar()
            ucPedidoxFixacao.SetarParametros(linha, txtRegistro.Text)
            Popup.ConsultaFixacaoPedido(Me.Page, "objFixacoPedido" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkBloquear_Click(sender As Object, e As EventArgs) Handles lnkBloquear.Click
        Try
            SessaoRecuperaPedido()

            If objPedido.BloquearPedido Then
                MsgBox(Me.Page, "Pedido bloqueado com sucesso!", eTitulo.Info)

                LimparPedido("U")
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRoteiroLog_Click(sender As Object, e As ImageClickEventArgs)
        If objPedido Is Nothing Then SessaoRecuperaPedido()
        Dim linha As Integer = CType(CType(sender, ImageButton).NamingContainer, GridViewRow).RowIndex

        Dim teste As String = objPedido.Roteiros(linha).Log
        txtRoteiroLog.Text = objPedido.Roteiros(linha).Log
        divRoteiroLog.Style.Clear()
    End Sub

    Protected Sub btnAdicionar_Click(sender As Object, e As EventArgs) Handles btnAdicionar.Click

        If String.IsNullOrWhiteSpace(txtDescricaoContrato.Text) Then
            MsgBox(Me.Page, "Descrição não foi informada", eTitulo.Info)
            txtDescricaoContrato.Focus()
            Exit Sub
        End If

        If fupArquivo.HasFile Then
            Dim NomeDoArquivo As String = Path.GetFileName(fupArquivo.PostedFile.FileName)
            Dim TamanhoDoArquivo As Long = fupArquivo.PostedFile.ContentLength
            Dim extensao As String = Path.GetExtension(NomeDoArquivo)
            Dim contentType As String = String.Empty

            If Not extensao.ToLower.Equals(".pdf") AndAlso
                Not extensao.ToLower.Equals(".xls") AndAlso
                Not extensao.ToLower.Equals(".xlsx") AndAlso
                Not extensao.ToLower.Equals(".doc") AndAlso
                Not extensao.ToLower.Equals(".docx") Then
                MsgBox(Me.Page, "São permitidos apenas documentos com extensões pdf, xls, xlsx, doc e docx.")
                Exit Sub
            End If

            SessaoRecuperaPedido()

            Dim pXc = New PedidoXContrato(objPedido)

            pXc.IUD = "I"
            pXc.Codigo = [Lib].Negocio.Numerador.PegarNumero(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, [Lib].Negocio.eTiposNumerador.PedidoXContrato)
            pXc.Descricao = txtDescricaoContrato.Text
            pXc.NomeDoArquivo = NomeDoArquivo

            pXc.Arquivo = fupArquivo.FileBytes
            txtNomeDoArquivo.Text = NomeDoArquivo

            If objPedido.IUD = "U" Then
                If pXc.Salvar Then
                    objPedido.Contratos = New ListPedidoXContrato(objPedido)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                End If
            Else
                objPedido.Contratos.Add(pXc)
            End If

            SessaoSalvaPedido()

            gridContratos.DataSource = objPedido.Contratos
            gridContratos.DataBind()

            Dim x As Integer = 0
            While x < gridContratos.Rows.Count
                Dim pExtensao() As String = gridContratos.Rows(x).Cells(2).Text.Split(".")

                If pExtensao(1).ToUpper.Equals("XML") Then
                    CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/xml16x16.png"
                ElseIf pExtensao(1).ToUpper.Equals("PDF") Then
                    CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icopdf16X16.jpg"
                ElseIf pExtensao(1).ToUpper.Equals("XLS") Then
                    CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                ElseIf pExtensao(1).ToUpper.Equals("XLSX") Then
                    CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                ElseIf pExtensao(1).ToUpper.Equals("DOC") Then
                    CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                ElseIf pExtensao(1).ToUpper.Equals("DOCX") Then
                    CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                Else
                    CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icoJpg.gif"
                End If

                CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ToolTip = String.Format("Download - {0}.{1}", pExtensao(0), pExtensao(1))
                CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"

                x += 1
            End While

            txtDescricaoContrato.Text = String.Empty
            txtNomeDoArquivo.Text = String.Empty
        Else
            MsgBox(Me.Page, "Selecione um arquivo.")
        End If
    End Sub

    Protected Sub imgDownload_Click(sender As Object, e As ImageClickEventArgs)
        Try
            SessaoRecuperaPedido()

            Dim imgArquivo As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgArquivo.NamingContainer, GridViewRow)

            Dim CaminhoNomeArquivo As String = Server.MapPath("~/Files/" & objPedido.Contratos(row.RowIndex).NomeDoArquivo)
            If Dir(CaminhoNomeArquivo).Length > 0 Then Kill(CaminhoNomeArquivo)

            Dim Fs As FileStream = New FileStream(CaminhoNomeArquivo, FileMode.Create)
            Fs.Write(objPedido.Contratos(row.RowIndex).Arquivo, 0, objPedido.Contratos(row.RowIndex).Arquivo.Length)
            Fs.Flush()
            Fs.Close()

            Funcoes.AbrirArquivo(Me.Page, "Files/" & objPedido.Contratos(row.RowIndex).NomeDoArquivo)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluir_Click(sender As Object, e As ImageClickEventArgs)

        Try
            If Funcoes.VerificaPermissao("PedidosXItens", "EXCLUIR") Then
                SessaoRecuperaPedido()

                Dim img As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

                Dim Pos = objPedido.Contratos.FindIndex(Function(s) s.Codigo = row.Cells(0).Text())

                objPedido.Contratos(Pos).IUD = "D"

                If objPedido.IUD = "U" Then
                    If objPedido.Contratos(Pos).Salvar Then

                        objPedido.Contratos = New ListPedidoXContrato(objPedido)

                        SessaoSalvaPedido()

                        gridContratos.DataSource = objPedido.Contratos
                        gridContratos.DataBind()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                        Exit Sub
                    End If
                Else
                    SessaoSalvaPedido()

                    gridContratos.DataSource = objPedido.Contratos.Where(Function(s) s.IUD <> "D")
                    gridContratos.DataBind()
                End If

                Dim x As Integer = 0
                While x < gridContratos.Rows.Count
                    Dim pExtensao() As String = gridContratos.Rows(x).Cells(2).Text.Split(".")

                    If pExtensao(1).ToUpper.Equals("XML") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/xml16x16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("PDF") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icopdf16X16.jpg"
                    ElseIf pExtensao(1).ToUpper.Equals("XLS") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("XLSX") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("DOC") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("DOCX") Then
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                    Else
                        CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icoJpg.gif"
                    End If

                    CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).ToolTip = String.Format("Download - {0}.{1}", pExtensao(0), pExtensao(1))
                    CType(gridContratos.Rows(x).FindControl("imgDownload"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"

                    x += 1
                End While

            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class
