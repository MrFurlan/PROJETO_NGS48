Imports System.IO
Imports System.Data
Imports System.Net.Configuration
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading
Imports System.Xml
Imports System.Globalization

Public Class OrdemDeCarregamento
    Inherits BasePage

#Region "Variáveis"

    Private Sql As String
    Private erroMsg As String = String.Empty
    Private MsgAlerta As String = String.Empty
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal
    Private objPlaca As [Lib].Negocio.Placa

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If IsConnect AndAlso Not IsPostBack Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Expedicao.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("OrdemDeCarregamento", "ACESSAR") Then

                    txtDataDeEmissao.Text = Today.ToString("dd/MM/yyyy")
                    txtDataCarregamento.Text = Today.ToString("dd/MM/yyyy")
                    primeiraVez.Value = True

                    LimparCampos("", "", False, Session("ssEmpresa"), Session("ssEndEmpresa"))

                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
            'SessaoVerificaCarregaObjetos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtObservacoesDeEmbarque_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        SessaoRecuperaNotaFiscal()
        If Not String.IsNullOrEmpty(txtObservacoesDeEmbarque.Text) Then
            objNotaFiscal.ObservacoesDeEmbarque = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesDeEmbarque.Text)
        Else
            objNotaFiscal.ObservacoesDeEmbarque = txtObservacoesDeEmbarque.Text

        End If

        txtObservacoesDeEmbarque.Text = objNotaFiscal.ObservacoesDeEmbarque

        SessaoSalvaNotaFiscal()
    End Sub

    Public Sub LimparCampos(ByVal evento As String, ByVal Numero As String, ByVal NossaEmissao As Boolean, ByVal pEmpresa As String, ByVal pEndEmpresa As Integer)

        If Not objNotaFiscal Is Nothing Then
            If String.IsNullOrWhiteSpace(objNotaFiscal.IUD) Then
                SessaoRecuperaNotaFiscal()
            End If
        End If

        Session.Remove("objHorarioServidor" & HID.Value)
        Session.Remove("objPedidoSelecionado" & HID.Value)
        Session.Remove("objItensPedidoSelecionadosNXI" & HID.Value)
        Session.Remove("objClienteNXI" & HID.Value)
        Session.Remove("objDepositoNXI" & HID.Value)
        Session.Remove("objEmbarqueNXI" & HID.Value)
        Session.Remove("objTransportadorNXI" & HID.Value)
        Session.Remove("objOrigemDestinoNXI" & HID.Value)
        Session.Remove("objTransbordoNXI" & HID.Value)
        Session.Remove("ProcurandoRomaneio" & HID.Value)
        Session.Remove("ProcurandoTrocaDeNota" & HID.Value)
        Session.Remove("ProcurandoClassificacao" & HID.Value)
        Session.Remove("objAutorizacaoNXI" & HID.Value)
        Session.Remove("SemTrocaDeNota" & HID.Value)
        Session.Remove("SemAutorizacao" & HID.Value)
        Session.Remove("SemProcuracao" & HID.Value)
        Session.Remove("MsgControle" & HID.Value)
        Session.Remove("ssCampo" & HID.Value)
        Session.Remove("ProcurandoRomaneio" & HID.Value)
        Session.Remove("ProcurandoClassificacao" & HID.Value)
        Session.Remove("ProcurandoProcuracao" & HID.Value)
        Session.Remove("TotalProcuracao" & HID.Value)
        Session.Remove("ProcurandoTrocaDeNota" & HID.Value)
        Session.Remove("ssMessage" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)
        Session.Remove("objNotaFiscalOriginal" & HID.Value)
        Session.Remove("objEmpresaNXI" & HID.Value)
        Session.Remove("objPlaca" & HID.Value)
        Session.Remove("objNFConsultaNXI" & HID.Value)
        Session.Remove("objTrocaDeNotaNXI" & HID.Value)
        Session.Remove("objRomaneioNXI" & HID.Value)
        Session.Remove("objProcuracaoNxI" & HID.Value)
        Session.Remove("TotalProcuracao" & HID.Value)
        Session.Remove("ProcurandoProcuracao" & HID.Value)
        Session.Remove("SemProcuracao" & HID.Value)
        Session.Remove("objClassificacaoNXI" & HID.Value)
        Session.Remove("ProcurandoClassificacao" & HID.Value)
        Session.Remove("objAutorizacaoNXI" & HID.Value)
        Session.Remove("SemAutorizacao" & HID.Value)
        Session.Remove("Observacoes" & HID.Value)
        Session.Remove("SemVendaAOrdem" & HID.Value)
        Session.Remove("ProcurandoVendaAOrdem" & HID.Value)
        Session.Remove("objVendaAOrdemNXI" & HID.Value)
        Session.Remove("objUFEmbarqueDI" & HID.Value)
        Session.Remove("objUFDesembarqueDI" & HID.Value)
        Session.Remove("objFabricanteDI" & HID.Value)
        Session.Remove("objEmailNFe" & HID.Value)
        Session.Remove("ObjNFObsProduto" & HID.Value)
        Session.Remove("objNFDeProdutor" & HID.Value)
        Session.Remove("objNFReferencialSaida" & HID.Value)
        Session.Remove("objEstoqueMinimo" & HID.Value)
        Session.Remove("objLoteFornecedor" & HID.Value)
        Session.Remove("chaveXMLautomacao" & HID.Value)
        Session.Remove("objConsultarNaviosXInvoice" & HID.Value)
        Session.Remove("objNavioXInvoice" & HID.Value)
        Session.Remove("objTransferencias" & HID.Value)
        Session.Remove("dsXml" & HID.Value)
        Session.Remove("ssNomeArquivoPedido" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
        ucConsultaEmpresas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucPedidoxSaldo.SetarHID(HID.Value)
        ucConsultaNotaTroca.SetarHID(HID.Value)
        ucConsultaRomaneios.SetarHID(HID.Value)
        ucConsultaObservacoes.SetarHID(HID.Value)
        ucConsultaPlacas.SetarHID(HID.Value)
        ucConsultaEstados.SetarHID(HID.Value)
        ucConsultaCodMunicipios.SetarHID(HID.Value)
        ucNotaFiscalXClassificacao.SetarHID(HID.Value)
        ucConsultaNotaVendaAOrdem.SetarHID(HID.Value)
        ucConsultaAutorizacaoDeRetirada.SetarHID(HID.Value)
        ucConsultaEncargosPlanoDeContas.SetarHID(HID.Value)
        ucConsultaDadosBancarios.SetarHID(HID.Value)
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucNotaFiscalReferencial.SetarHID(HID.Value)
        ucNotaDeDevolucaoXNota.SetarHID(HID.Value)
        ucConsultaProcuracao.SetarHID(HID.Value)
        ucInutilizacao.SetarHID(HID.Value)
        ucVencimentos.SetarHID(HID.Value)
        ucMonitorDeNotas.SetarHID(HID.Value)
        ucEmailNFe.SetarHID(HID.Value)
        ucNFObsProduto.SetarHID(HID.Value)
        ucNFEncargo.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)
        ucNFProdutor.SetarHID(HID.Value)
        ucNFReferencialSaida.SetarHID(HID.Value)
        ucConsultarNaviosXInvoice.SetarHID(HID.Value)
        ucTransferencias.SetarHID(HID.Value)

        Session("Carregando" & HID.Value) = False

        imgUsuario.Visible = False
        lblUsuario.Visible = False
        imgExtratoPedido.Visible = False
        imgRomaneio.Visible = False
        btnRomaneio.Visible = True

        HabilitarBotoes()

        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkLimpar.Parent.Visible = True
        lnkEnviarEmail.Parent.Visible = False
        lnkPdf.Parent.Visible = False
        lnkImpressora.Parent.Visible = False

        txtDataDeEmissao.Enabled = True
        txtDataCarregamento.Enabled = True

        txtDataDeEmissao.Text = Now.ToString("dd/MM/yyyy")
        txtDataCarregamento.Text = Now.ToString("dd/MM/yyyy")
        txtPedido.Text = ""

        txtRomaneio.Text = ""
        txtPesoRomaneio.Text = ""
        txtAutorizacao.Text = ""

        ucVencimentos.Limpar()
        ddlEmbarque.Enabled = True
        ddlEmbarque.Items.Clear()

        txtValorTotalDosProdutos.Text = ""
        txtValorTotalDaNota.Text = ""
        txtBaseIcmsNota.Text = ""
        txtValorIcmsNota.Text = ""
        txtValorIPINota.Text = ""
        txtValorBaseIcmsSTNota.Text = ""
        txtValorIcmsSTNota.Text = ""
        txtValorFrete.Text = ""
        txtDesconto.Text = ""
        txtOutras.Text = ""
        txtCodigoPlaca.Value = ""

        txtVolumes.Text = ""
        txtEspecie.Text = ""
        txtMarca.Text = ""
        txtNumeracao.Text = ""

        txtPesoBruto.Text = ""
        txtPesoLiquido.Text = ""
        txtObservacoesDeEmbarque.Text = ""
        txtObservacoesFiscais.Text = ""
        txtObservacoesInternas.Text = ""

        objNotaFiscal = New [Lib].Negocio.NotaFiscal
        objNotaFiscal.IUD = "I"

        objNotaFiscal.Usuario = Session("ssNomeUsuario")
        objNotaFiscal.UsuarioInclusao = Session("ssNomeUsuario")
        objNotaFiscal.DataInclusao = Format(Date.Now, "yyyy-MM-dd HH:mm:ss")
        ddlUsuarios.Items.Clear()
        ddlUsuarios.Items.Add("Inc.- " & objNotaFiscal.UsuarioInclusao)

        CarregarDDLFrete(False)

        SessaoSalvaNotaFiscal()

        If primeiraVez.Value = True Then
            primeiraVez.Value = False
            SetarEmpresa(pEmpresa, pEndEmpresa, True)
        Else
            SetarEmpresa(pEmpresa, pEndEmpresa, False)
        End If

        TabContainer1.ActiveTabIndex = 0

        Dim Ip As String = Context.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        If Ip = String.Empty Then
            Ip = Context.Request.ServerVariables("REMOTE_ADDR")
            If Ip = "::1" Then Ip = "127.0.0.1"
        End If

        lblAcessoUsuario.Text = Ip

        If evento = "E" Then
            MsgBox(Me.Page, "Nota fiscal " & Numero & " removida com Sucesso.", eTitulo.Sucess)
        ElseIf evento = "I" Then
            MsgBox(Me.Page, "Nota fiscal " & Numero & " incluída com Sucesso.", eTitulo.Sucess)
        ElseIf evento = "U" Then
            MsgBox(Me.Page, "Nota fiscal " & Numero & " alterada com Sucesso.", eTitulo.Sucess)
        ElseIf evento = "C" Then
            MsgBox(Me.Page, "Nota fiscal " & Numero & " cancelada com Sucesso.", eTitulo.Sucess)
        End If
    End Sub

    Private Sub CarregarDDLFrete(ByVal bCarregarXML As Boolean)
        Try
            'Carregas opções dos tipos de frete em relação ao tipo definido na NF Caso seja CIF então somente estarão disponível CIF e
            ' NEN, caso seja FOB então FOB e NEN. 
            ddlFrete.Items.Clear()
            Dim TipoDeFrete As New Dictionary(Of String, String)

            If objNotaFiscal IsNot Nothing AndAlso (objNotaFiscal.Pedido IsNot Nothing AndAlso objNotaFiscal.CodigoPedido > 0) OrElse bCarregarXML Then
                If objNotaFiscal.CIFFOB = eTiposFrete.CIF Then
                    TipoDeFrete.Add("CIF", "CIF - Emitente")    'Saída CIF - 0 - Por conta do emitente = EMPRESA
                    TipoDeFrete.Add("TER", "TER - Terceiro")
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")
                ElseIf objNotaFiscal.CIFFOB = eTiposFrete.FOB Then
                    TipoDeFrete.Add("FOB", "FOB - Destinatario") 'Entrada FOB - 1 – Por conta do destinatário = EMPRESA
                    TipoDeFrete.Add("TER", "TER - Terceiro")
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")
                ElseIf objNotaFiscal.CIFFOB = eTiposFrete.TER Then
                    TipoDeFrete.Add("TER", "TER - Terceiro") '2 – Por conta de terceiros
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")
                ElseIf objNotaFiscal.CIFFOB = eTiposFrete.NEN AndAlso Not objNotaFiscal.Pedido Is Nothing AndAlso objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.CIF Then
                    If objNotaFiscal.SubOperacao.Devolucao Then
                        TipoDeFrete.Add("FOB", "FOB - Destinatario")
                    Else
                        TipoDeFrete.Add("CIF", "CIF - Emitente")
                    End If
                    TipoDeFrete.Add("TER", "TER - Terceiro")
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")
                ElseIf objNotaFiscal.CIFFOB = eTiposFrete.NEN AndAlso Not objNotaFiscal.Pedido Is Nothing AndAlso objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.FOB Then
                    If objNotaFiscal.SubOperacao.Devolucao Then
                        TipoDeFrete.Add("CIF", "CIF - Emitente")
                    Else
                        TipoDeFrete.Add("FOB", "FOB - Destinatario")
                    End If
                    TipoDeFrete.Add("TER", "TER - Terceiro")
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")

                Else
                    TipoDeFrete.Add("NEN", "NEN - Nenhum")
                End If

                ddlFrete.DataTextField = "Value"
                ddlFrete.DataValueField = "Key"
                ddlFrete.DataSource = TipoDeFrete
                ddlFrete.DataBind()

                ddlFrete.Enabled = True

                ddlFrete.SelectedValue = objNotaFiscal.CIFFOB.ToString

                If bCarregarXML = False Then
                    'Caso a opção nenhum NEN esteja setado tanto na NF quando no pedido então não será possível escolher outra
                    'Só está sendo feita a validação do tipo na NF para casos que ainda não tenham sido corrigidos já que é necessário fazer carta de correção eletrônica.
                    'A regra é que se no pedido estiver NEN então em suas notas também será NEN sem opção de mudança.
                    If objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.NEN AndAlso objNotaFiscal.CIFFOB = eTiposFrete.NEN Then
                        ddlFrete.SelectedValue = eTiposFrete.NEN.ToString
                        ddlFrete.Enabled = False
                    End If
                End If

            Else
                TipoDeFrete.Add("NEN", "NEN - Nenhum")

                ddlFrete.DataTextField = "Value"
                ddlFrete.DataValueField = "Key"
                ddlFrete.DataSource = TipoDeFrete
                ddlFrete.DataBind()

                ddlFrete.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub HabilitarBotoes()

        btnPedido.Enabled = True
        btnRomaneio.Enabled = True
        BtnRetirada.Enabled = True
        btnTransportador.Enabled = True
        btnObservacoesFiscais.Enabled = True

    End Sub

    Public Sub BuscarPlaca()
        If txtCodigoTransportador.Text.Length > 0 Then
            ucConsultaPlacas.Limpar()
            Popup.ConsultaDePlacas(Me.Page, "objPlaca" & HID.Value)
            Popup.CenterDialog(Me.Page, "divConsultaPlacas")
        Else
            MsgBox(Me.Page, "Transportador não foi selecionado")
        End If
    End Sub

    Private Sub SetarEmpresa(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer, ByVal verificamodo As Boolean)
        Dim objEmpresa As New [Lib].Negocio.Cliente(pEmpresa, pEndEmpresa)

        SessaoRecuperaNotaFiscal()

        objNotaFiscal.Empresa = CType(objEmpresa, [Lib].Negocio.Cliente)

        With objNotaFiscal.Empresa
            objNotaFiscal.CodigoEmpresa = .Codigo
            objNotaFiscal.EnderecoEmpresa = .CodigoEndereco
            txtMarca.Text = .Empresa.Marca
            objNotaFiscal.Marca = .Empresa.Marca
        End With

        SessaoSalvaNotaFiscal()

        'Verifica se a empresa está habilitada para gravar arquivo
        Dim Empresa As New [Lib].Negocio.Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)

        LiberaEmpresa()

    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            btnEmpresa.Enabled = False
        End If

        SessaoRecuperaNotaFiscal()

        If objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica AndAlso objNotaFiscal.Empresa.Empresa.NossaEmissao Then
            Dim fileEmpresa As String = "C:\NGS\xmlCopy\" & objNotaFiscal.Empresa.Codigo & ".xml"

            If Not File.Exists(fileEmpresa) Then
                MsgBox(Me.Page, "Arquivo de Configuracão da Empresa não foi encontrado.", "~/Expedicao.aspx")
                Exit Sub
            End If

            Dim xmlConf As New XmlDocument
            xmlConf.Load(fileEmpresa)

            Session("objHorarioServidor" & HID.Value) = xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("horarioServidor").InnerText
        End If
    End Sub

    Public Sub ConsultaPedidos(Optional ByVal carregarDados As Boolean = False)
        SessaoRecuperaNotaFiscal()
        If String.IsNullOrWhiteSpace(objNotaFiscal.CodigoCliente) Then
            MsgBox(Me.Page, "Cliente não foi selecionado")
            Exit Sub
        End If
        ucPedidoxSaldo.LimparCampos()
        If (carregarDados) Then
            ucPedidoxSaldo.CarregarDadosNotaFiscalXItem()
        End If
        Popup.ConsultaDePedidoXSaldoXFinanceiro(Me, "objItensPedidoSelecionadosNXI" & HID.Value, "txtPedido", True, 500)
    End Sub

    Public Sub ListarClientes()
        ucConsultaClientes.Limpar()
        Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
        Popup.ConsultaDeClientes(Me, "objClienteNXI" & HID.Value, txtNome.ClientID, True, 500)
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Session("objEmpresaNXI" & HID.Value) IsNot Nothing Then
                SetarEmpresa(CType(Session("objEmpresaNXI" & HID.Value), [Lib].Negocio.Cliente).Codigo, CType(Session("objEmpresaNXI" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco, True)
                Session.Remove("objEmpresaNXI" & HID.Value)
                ListarClientes()
            ElseIf Session("objTransportadorNXI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.Transportador = CType(Session("objTransportadorNXI" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemTransportador As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Transportador)
                With objNotaFiscal.Transportador
                    objNotaFiscal.CodigoTransportador = .Codigo
                    objNotaFiscal.EnderecoTransportador = .CodigoEndereco
                    lblNomeTransportador.Text = .Nome
                    txtCodigoTransportador.Text = .Codigo
                    lblEnderecoTransportador.Text = .Endereco & ", " & .Numero.ToString()
                    lblCidadeTransportador.Text = .Cidade
                    lblUfTransportador.Text = .CodigoEstado
                    lblIeTransportador.Text = .InscricaoEstadual
                End With
                Session.Remove("objTransportadorNXI" & HID.Value)
                SessaoSalvaNotaFiscal()
                BuscarPlaca()

            ElseIf Session("objEmbarqueNXI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.CodigoLocalEmbarque = CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).Codigo
                objNotaFiscal.EndLocalEmbarque = CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
                Dim intIndice As Integer = ddlEmbarque.Items.IndexOf(ddlEmbarque.Items.FindByValue(objNotaFiscal.CodigoLocalEmbarque & "-" & CStr(objNotaFiscal.EndLocalEmbarque)))
                If intIndice = -1 Then
                    Funcoes.AdicionarClienteAoDDL(ddlEmbarque, objNotaFiscal.LocalEmbarque)
                End If

                objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|Local de Embarque: " & CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).CodigoFormatado & " " & CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).Nome & " - " & CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).Cidade & "/" & CType(Session("objEmbarqueNXI" & HID.Value), [Lib].Negocio.Cliente).CodigoEstado

                ddlEmbarque.SelectedValue = objNotaFiscal.CodigoLocalEmbarque & "-" & objNotaFiscal.EndLocalEmbarque

                Session.Remove("objEmbarqueNXI" & HID.Value)
                SessaoSalvaNotaFiscal()

            ElseIf Session("objPlaca" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                objPlaca = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa)
                objNotaFiscal.PlacaTransportador = LTrim(RTrim(objPlaca.Placa01.ToUpper))
                txtCodigoPlaca.Value = objNotaFiscal.PlacaTransportador
                lblPlaca1.Text = objPlaca.Placa01.ToUpper
                lblUfPlaca1.Text = objPlaca.EstadoPlaca01

                If Not objPlaca.Motorista Is Nothing AndAlso objPlaca.Motorista.Codigo.Length > 0 AndAlso Not objNotaFiscal.ObservacoesDeEmbarque.Contains("Motorista:") Then
                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|Motorista: " & objPlaca.Motorista.Nome & " - CPF " & objPlaca.Motorista.CodigoFormatado & " - CNH " & objPlaca.Habilitacao & " - Placa " & objPlaca.Placa01 & " " & objPlaca.Placa02 & " " & objPlaca.Placa03 & " " & objPlaca.Placa04
                End If

                Session.Remove("objPlaca" & HID.Value)

                SessaoSalvaNotaFiscal()
            ElseIf Session("objItensPedidoSelecionadosNXI" & HID.Value) IsNot Nothing Then
                SessaoRecuperaNotaFiscal()

                objNotaFiscal.DataDevolucao = objNotaFiscal.Movimento

                If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                    objNotaFiscal.Especie = objNotaFiscal.Pedido.DescricaoEmbalagem
                End If

                'OBRIGATÓRIO O ENVIO DA DANFE E XML QUANDO FOR NOSSA EMISSAO - 14/09/2016 - FURLAN
                'VERIFICA SE EXISTE O MUNICIPIO IBGE - FURLAN - 03/10/2016 - FURLAN
                If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then
                    If String.IsNullOrWhiteSpace(objNotaFiscal.Cliente.EmailNFE) Then
                        MsgBox(Me.Page, "Email para Recebimento da NFE é obrigatório quando for NOSSA EMISSÃO, favor preencher no Cadastro do Cliente " & objNotaFiscal.Cliente.CodigoFormatado & "-" & objNotaFiscal.Cliente.Nome & " o campo EMAIL NFE.")
                        Exit Sub
                    ElseIf objNotaFiscal.Cliente.Municipio.CodigoIbge = 0 Then
                        MsgBox(Me.Page, "Código do Município da tabela do IBGE não foi encontrado, favor verificar no Cadastro do Cliente " & objNotaFiscal.Cliente.CodigoFormatado & "-" & objNotaFiscal.Cliente.Nome & " o campo CIDADE e o código do Município informado.")
                        Exit Sub
                    End If
                End If

                'PAGAMENTO/RECEBIMENTO ANTECIPADO O DEVE TER TITULO EM ADIANTAMENTO
                If FinanceiroNovo And objNotaFiscal.SubOperacao.Financeiro Then
                    If objNotaFiscal.Pedido.CondicaoPagamento IsNot Nothing AndAlso objNotaFiscal.Pedido.CondicaoPagamento.Antecipado AndAlso objNotaFiscal.Titulos.AdiantamentosAbertos.Count = 0 Then
                        Session.Remove("objItensPedidoSelecionadosNXI")
                        MsgBox(Me.Page, "Pagamento/Recebimento antecipado deve ter titulo em adiantamento")
                        Exit Sub
                    End If
                End If

                If objNotaFiscal.SubOperacao.ProprietarioDaMercadoria Then
                    Dim numdep As Integer
                    numdep = objNotaFiscal.Pedido.Depositos.Where(Function(s) s.Tipo = "PM").Count
                    If numdep = 0 Then
                        MsgBox(Me.Page, "Não foi informado o Proprietário da Mercadoria no Pedido.")
                        Exit Sub
                    End If

                    objNotaFiscal.CodigoProprietarioDaMercadoria = objNotaFiscal.Pedido.Depositos.Where(Function(s) s.Tipo = "PM")(0).Codigo
                    objNotaFiscal.EnderecoProprietarioDaMercadoria = objNotaFiscal.Pedido.Depositos.Where(Function(s) s.Tipo = "PM")(0).CodigoEndereco
                End If

                txtDataDeEmissao.Enabled = False

                '2013.08.27 - Acrescentadas verificações para permitir diferença de 1 centavo entre os valores dos pedidos para troca'
                If objNotaFiscal.Pedido.Troca Then
                    If [Enum].Parse(GetType(eClassesOperacoes), objNotaFiscal.Pedido.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.COMPRAS Then

                        If Not objNotaFiscal.SubOperacao.Devolucao AndAlso ((objNotaFiscal.Pedido.CodigoMoeda = eTiposMoeda.Oficial AndAlso Math.Abs(objNotaFiscal.Pedido.PedidoTroca.Itens.LiquidoOficial - objNotaFiscal.Pedido.Itens.LiquidoOficial) > 0.01) _
                            OrElse (objNotaFiscal.Pedido.CodigoMoeda = eTiposMoeda.MoedaEstrangeira AndAlso Math.Abs(objNotaFiscal.Pedido.PedidoTroca.Itens.LiquidoOficial - objNotaFiscal.Pedido.Itens.LiquidoOficial) > 0.1)) Then
                            Session.Remove("objItensPedidoSelecionadosNXI" & HID.Value)
                            MsgBox(Me.Page, "Os pedidos de troca estão com valores diferentes!")
                            Exit Sub
                        End If

                    ElseIf objNotaFiscal.Pedido.PedidoTroca Is Nothing Then
                        Session.Remove("objItensPedidoSelecionadosNXI" & HID.Value)
                        MsgBox(Me.Page, "Pedido de Troca não está vinculado, favor verificar!")
                        Exit Sub
                    Else
                        If Not objNotaFiscal.SubOperacao.Devolucao _
                            AndAlso ((objNotaFiscal.Pedido.CodigoMoeda = eTiposMoeda.Oficial AndAlso Math.Abs(objNotaFiscal.Pedido.PedidoTroca.Itens.LiquidoOficial - objNotaFiscal.Pedido.Itens.LiquidoOficial) > 0.01) _
                            OrElse (objNotaFiscal.Pedido.CodigoMoeda = eTiposMoeda.MoedaEstrangeira AndAlso Math.Abs(objNotaFiscal.Pedido.PedidoTroca.Itens.LiquidoOficial - objNotaFiscal.Pedido.Itens.LiquidoOficial) > 0.1)) Then

                            Session.Remove("objItensPedidoSelecionadosNXI" & HID.Value)
                            MsgBox(Me.Page, "Os pedidos de troca estão com valores diferentes!")
                            Exit Sub
                        End If
                    End If
                End If

                If objNotaFiscal.Pedido.Itens(0).Fixacoes.Count > 0 AndAlso objNotaFiscal.Pedido.Itens(0).Fixacoes(0).IndiceFixado > 0 Then
                    txtPedido.Text = objNotaFiscal.Pedido.Codigo & " - Indice Fixado " & objNotaFiscal.Pedido.Itens(0).Fixacoes(0).IndiceFixado
                ElseIf objNotaFiscal.Pedido.IndiceFixado > 0 Then
                    txtPedido.Text = objNotaFiscal.Pedido.Codigo & " - Indice Fixado " & objNotaFiscal.Pedido.IndiceFixado
                Else
                    txtPedido.Text = objNotaFiscal.Pedido.Codigo
                End If

                objNotaFiscal.CodigoPedido = objNotaFiscal.Pedido.Codigo

                objNotaFiscal.CodigoSituacao = 1

                erroMsg = String.Empty
                MsgAlerta = String.Empty

                If CarregarItensComSaldo(CType(obj, [Lib].Negocio.SaldoPedido2015), erroMsg, MsgAlerta) Then
                    SessaoRecuperaNotaFiscal()
                Else
                    If erroMsg.Length > 0 Then
                        If String.IsNullOrWhiteSpace(MsgAlerta) Then
                            MsgBox(Me.Page, erroMsg, eTitulo.Erro, False)
                        Else
                            MsgBox(Me.Page, MsgAlerta & erroMsg, eTitulo.Erro, False)
                        End If
                    End If
                    LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                    Exit Sub
                End If

                If Not String.IsNullOrWhiteSpace(MsgAlerta) Then
                    MsgBox(Me.Page, MsgAlerta, eTitulo.Info, False)
                End If

                ''Enviar E-mail caso tenha chego no estoque mínimo - Furlan - 28/10/2020
                'If CType(Session("objEstoqueMinimo" & HID.Value), DataTable).Rows.Count > 0 Then enviarEstoqueMinimo()

                'Nenhum
                If objNotaFiscal.SubOperacao.Devolucao AndAlso Not objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.NEN Then 'Quando é devolução Inverte'
                    objNotaFiscal.CIFFOB = IIf(objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.CIF, eTiposFrete.FOB, eTiposFrete.CIF)
                Else 'senão coloca o que estiver no pedido'
                    objNotaFiscal.CIFFOB = objNotaFiscal.Pedido.FreteCIFFOB
                End If

                CarregarDDLFrete(False)

                If objNotaFiscal.Pedido.FreteCIFFOB = eTiposFrete.NEN OrElse Not objNotaFiscal.SubOperacao.QuantidadeFisico Then

                    If Left(objNotaFiscal.CodigoEmpresa, 8) = "05272759" OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.OUTRAS OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.TRANSFERENCIAS Then
                        'SE FOR A 'QUÍMICA LIBERA POIS NÃO USA ESTOQUE FÍSICO FURLAN - 29/09/2022
                    Else
                        ddlFrete.SelectedValue = eTiposFrete.NEN.ToString
                        ddlFrete.Enabled = False
                    End If
                End If

                'INICIA OBSERVACAO NOTA FISCAL AQUI - FURLAN - 25/07-2014
                If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica Then

                    If objNotaFiscal.Empresa.Empresa.ObservacaoSefazNFE.Length > 0 Then
                        Dim obsSefaz() As String = objNotaFiscal.Empresa.Empresa.ObservacaoSefazNFE.Split(";")
                        If objNotaFiscal.Empresa.CodigoEstado = "MT" Then
                            If obsSefaz.Length > 1 AndAlso Left(objNotaFiscal.CodigoEmpresa, 8) = "04854422" AndAlso objNotaFiscal.Itens.Count > 0 AndAlso objNotaFiscal.Itens(0).Produto.ControlarEmbalagem Then
                                objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(1)
                            Else
                                objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(0)
                            End If
                        ElseIf objNotaFiscal.Empresa.CodigoEstado = "MS" Then
                            objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(0)
                        ElseIf objNotaFiscal.Empresa.CodigoEstado = "PR" Then
                            If objNotaFiscal.Empresa.Cidade = "UMUARAMA" Or objNotaFiscal.Empresa.Cidade = "CASCAVEL" Then
                                objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(0)
                            Else
                                If obsSefaz.Length > 1 AndAlso (objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10101" OrElse
                                    objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10201" OrElse
                                    objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10202" OrElse
                                    objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10203") Then

                                    objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(1)
                                Else
                                    objNotaFiscal.ObservacoesDeEmbarque = obsSefaz(0)
                                End If
                            End If
                        End If
                    End If

                End If

                If objNotaFiscal.Itens(0).OperacaoEstado.CodigoBeneficio.Length > 0 Then
                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|" & objNotaFiscal.Itens(0).OperacaoEstado.BeneficioICMS.Descricao
                End If

                If objNotaFiscal.ObservacoesDeEmbarque.Length > 0 Then
                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|PEDIDO: " & objNotaFiscal.CodigoPedido & ". "
                Else
                    objNotaFiscal.ObservacoesDeEmbarque = "|PEDIDO: " & objNotaFiscal.CodigoPedido & ". "
                End If

                'POR HORA MANUAL POIS AINDA NÃO TEM O NR DO PROCESSO, DEPOIS VAMOS VER PARA COLOCARMOS EM UMA TABELA PARA PEGAR AUTOMÁTICO - FURLAN - 05-07-2024
                If Left(objNotaFiscal.CodigoEmpresa, 8) = "44979506" AndAlso objNotaFiscal.Itens(0).Produto.CodigoGrupo = "10101" Then
                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "PRODUTO PROVENIENTE DE ARMAZEM PARTICIPANTE CNPJ: " & objNotaFiscal.Empresa.CodigoFormatado & "."
                End If

                If objNotaFiscal.Pedido.PedidoEfetivo.Length > 0 Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & " - CONFIRMACAO DE NEGOCIO: " & objNotaFiscal.Pedido.PedidoEfetivo
                If objNotaFiscal.Pedido.Contrato.Length > 0 Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & " - CONTRATO: " & objNotaFiscal.Pedido.Contrato

                Dim Certidao As New [Lib].Negocio.CertidaoNegativa(objNotaFiscal.CodigoEmpresa, False)
                If Certidao.CodigoCliente.Length > 0 Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "| Certidao Negativa de Debitos nr " & Certidao.Numero & " / Cod. de Autenticidade " & Certidao.CodigoAutenticidade

                If objNotaFiscal.Operacao.CodigoClasse = "VENDAS" AndAlso objNotaFiscal.Pedido.Troca Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "| Pedido com Vencimento em " & objNotaFiscal.Pedido.DataVencimentoPedido

                Dim teveEncrgo As Boolean = False

                For Each item In objNotaFiscal.Itens

                    If objNotaFiscal.CodigoOperacao = objNotaFiscal.Pedido.CodigoOperacao AndAlso objNotaFiscal.CodigoSubOperacao = objNotaFiscal.Pedido.CodigoSubOperacao AndAlso item.QuantidadeFiscal = objNotaFiscal.Pedido.Itens.FirstOrDefault(Function(s) s.CodigoProduto = item.CodigoProduto).QuantidadePedidoFaturamento Then
                        For Each enc In item.Encargos
                            If enc.Codigo = "IPI" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "IPI").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "FRETES" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "FRETES").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "DESP.ADUANEIRAS" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "DESP.ADUANEIRAS").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "DESCONTOS" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "DESCONTOS").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "SEGURO" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "SEGURO").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "ICMS" Then
                                enc.Base = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "ICMS" OrElse s.CodigoEncargo = "ICMS A REC.").Sum(Function(t) t.BaseOficial)
                                enc.Percentual = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "ICMS" OrElse s.CodigoEncargo = "ICMS A REC.").Sum(Function(t) t.Percentual)
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "ICMS" OrElse s.CodigoEncargo = "ICMS A REC.").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            ElseIf enc.Codigo = "CUSTO ICMS" Then
                                enc.Valor = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = item.CodigoProduto).FirstOrDefault.Encargos.Where(Function(s) s.CodigoEncargo = "CUSTO ICMS").Sum(Function(t) t.ValorOficial)
                                teveEncrgo = True
                            End If
                        Next
                    End If
                Next

                If teveEncrgo Then objNotaFiscal.AtualizaTotais()

                If Not String.IsNullOrWhiteSpace(objNotaFiscal.PlacaTransportador) AndAlso Not objNotaFiscal.ObservacoesDeEmbarque.Contains("Motorista:") Then
                    Dim objPlaca As New [Lib].Negocio.Placa(objNotaFiscal.PlacaTransportador)
                    If Not objPlaca.Motorista Is Nothing Then
                        objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|Motorista: " & objPlaca.Motorista.Nome & " - CPF " & objPlaca.Motorista.CodigoFormatado & " - CNH " & objPlaca.Habilitacao & " - Placa " & objPlaca.Placa01 & " " & objPlaca.Placa02 & " " & objPlaca.Placa03 & " " & objPlaca.Placa04
                    Else
                        objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & " - Placa " & objPlaca.Placa01 & " " & objPlaca.Placa02 & " " & objPlaca.Placa03 & " " & objPlaca.Placa04
                    End If
                End If


                'Código do Resgistro do Ministéri da Agriculcura Tabela Produtos
                If objNotaFiscal.Empresa.Empresa.UsarRegistroMinAgr Then

                    Dim nrRegistro As String = ""
                    Dim separa As String = ""
                    Dim temRegistro As Boolean = False

                    For Each item In objNotaFiscal.Itens

                        If item.Produto.RegistroMinisterioAgricultura.Length > 0 Then
                            nrRegistro = nrRegistro & item.Produto.RegistroMinisterioAgricultura & separa
                            separa = ","
                            temRegistro = True
                        End If
                    Next

                    If temRegistro Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "REGISTRO DO PRODUTO: " & nrRegistro

                End If

                If Left(objNotaFiscal.Empresa.Codigo, 8) = "24450490" AndAlso objNotaFiscal.Eletronica AndAlso objNotaFiscal.NossaEmissao AndAlso Left(objNotaFiscal.Itens(0).Produto.CodigoGrupo, 1) = "2" Then
                    Dim origemDaMercadoria As String = String.Empty

                    If objNotaFiscal.Pedido.InvoiceNavio > 0 Then
                        origemDaMercadoria = "IMPORTADO"
                    ElseIf objNotaFiscal.Itens(0).Encargos.EncProduto.SituacaoTributaria > 99 Then
                        origemDaMercadoria = "IMPORTADO"
                    Else
                        origemDaMercadoria = "NACIONAL"
                    End If

                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|PRODUTO " & origemDaMercadoria & ". NATUREZA FISICA DO PRODUTO: " & objNotaFiscal.Itens(0).Produto.EstadoFisico.Descricao & ". "
                    Dim dValidade = objNotaFiscal.Movimento.AddYears(1)
                    objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque & "|MODO DE APLICACAO PRINCIPAL: VIA SOLO. ARMAZENAR EM LUGAR SECO, COBERTO E AREJADO. DATA DE FABRICACAO: " & objNotaFiscal.Movimento.ToString("dd/MM/yyyy") & " DATA DE VALIDADE: " & dValidade.ToString("dd/MM/yyyy") & ". "

                End If

                SessaoSalvaNotaFiscal()

                SessaoRecuperaNotaFiscal()

                lnkNovo.Parent.Visible = True
                lnkConsultar.Parent.Visible = False

                If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.CONTAEORDEM Then
                    If objNotaFiscal.SubOperacao.Devolucao Then
                        If objNotaFiscal.Itens(0).Produto.Agrupar = "N" AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico Then
                            BuscaClassificacao()
                        Else
                            objNotaFiscal.Itens(0).PesoFiscal = 0
                            objNotaFiscal.Itens(0).QuantidadeFisica = 0
                            objNotaFiscal.Itens(0).QuantidadeFiscal = 0
                            objNotaFiscal.Itens(0).Unitario = 0
                            objNotaFiscal.Itens(0).ValorTotal = 0

                            SessaoSalvaNotaFiscal()

                        End If

                    End If

                End If

            ElseIf Session("objEmailNFe" & HID.Value) IsNot Nothing Then

                Try
                    SessaoRecuperaNotaFiscal()
                    objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNotaFiscal)
                    Dim lst As [Lib].Negocio.ListCliente = CType(Session("objEmailNFe" & HID.Value), [Lib].Negocio.ListCliente)
                    Dim fm As New FilesManager()
                    If fm.IsConnect() Then
                        Dim Sqls As New ArrayList
                        Dim objFil As New [Lib].Negocio.Fil()
                        objFil.IUD = "I"
                        objFil.NomeArquivo = String.Format("email{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa)
                        objFil.Texto = getTextoEmail(objNotaFiscal, lst, Session("strAssunto" & HID.Value), Session("strMensagem" & HID.Value), CBool(Session("strCompactado" & HID.Value)))
                        objFil.SalvarSql(Sqls)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        End If

                        Dim resp As [Lib].Negocio.Resp = Nothing
                        Dim fileName As String = String.Format("resp-email{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa)

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

                            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "4014" Then
                                MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                                Exit Sub
                            Else
                                MsgBox(Me.Page, strMsg)
                            End If

                            Sqls.Clear()
                            Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-email{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                            Sqls.Add(Sql)

                            If Not Banco.GravaBanco(Sqls) Then
                                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            End If
                        Else
                            MsgBox(Me.Page, "Falha no retorno da Sefaz, tente o reenvio do Email novamente.")
                        End If
                    Else
                        MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
                    End If

                    Session.Remove("objEmailNFe" & HID.Value)

                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getTextoEmail(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal lstClientes As [Lib].Negocio.ListCliente, ByVal strAssunto As String, ByVal strMensagem As String, Optional ByVal compactado As Boolean = False) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        For Each objCliente As [Lib].Negocio.Cliente In lstClientes.Where(Function(s) Not String.IsNullOrWhiteSpace(s.EmailNFE)).ToList()
            sb.Append("DESTINATARIO=" & objCliente.EmailNFE & ControlChars.CrLf)
        Next
        sb.Append("ASSUNTO=" & strAssunto & " - NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & strMensagem & " - Envio NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        If Not String.IsNullOrWhiteSpace(nf.Empresa.EmailNFE) Then
            Dim strEmpresa As String = nf.Empresa.EmailNFE.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)(0).Trim()
            sb.Append("EMAILEMITENTE=" & strEmpresa & ControlChars.CrLf)
            sb.Append("NOMEEMITENTE=" & nf.Empresa.Nome & ControlChars.CrLf)
        End If
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)
        sb.Append("COMPACTADO=" & IIf(compactado, "SIM", "NAO") & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Function ValidarCertidaoClienteNota() As Boolean
        If objNotaFiscal.Empresa.Empresa.CertidaoNegativa And objNotaFiscal.Empresa.CodigoEstado = objNotaFiscal.Cliente.CodigoEstado Then
            Dim encontrou As Boolean = False
            For Each row As [Lib].Negocio.ClientexTipo In objNotaFiscal.Cliente.Tipos
                If row.CodigoTipo = 4 Or row.CodigoTipo = 5 Then
                    encontrou = True
                    Exit For
                End If
            Next

            Dim strMsg As String = ""
            Dim cn As [Lib].Negocio.CertidaoNegativa
            Dim DiasParaOVencimento As Integer = 0

            'Cliente
            If encontrou Then
                cn = New [Lib].Negocio.CertidaoNegativa(objNotaFiscal.Cliente.Codigo, True)
                DiasParaOVencimento = DateDiff("d", FormatDateTime(Date.Now, DateFormat.ShortDate), FormatDateTime(cn.DataValidade, DateFormat.ShortDate))
                If cn.CodigoCliente.Length = 0 OrElse DiasParaOVencimento < 0 Then
                    MsgBox(Me.Page, "Cliente sem certidão negativa! ")
                    Return False
                ElseIf DiasParaOVencimento > 0 AndAlso DiasParaOVencimento <= 5 Then
                    strMsg = "Falta(m) " & DiasParaOVencimento & " dia(s) para o Vencimento da Certidão Negativa do Cliente" & vbCrLf
                End If
            End If

            'Empresa
            cn = New [Lib].Negocio.CertidaoNegativa(objNotaFiscal.CodigoEmpresa, False)
            DiasParaOVencimento = DateDiff("d", FormatDateTime(Date.Now, DateFormat.ShortDate), FormatDateTime(cn.DataValidade, DateFormat.ShortDate))
            If cn.CodigoCliente.Length = 0 OrElse DiasParaOVencimento < 0 Then
                strMsg &= IIf(strMsg.Length > 0, ", ", "") & "EMPRESA Sem Certidao Negativa!"
                MsgBox(Me.Page, strMsg)
                Return False
            ElseIf DiasParaOVencimento > 0 AndAlso DiasParaOVencimento <= 2 Then
                strMsg &= IIf(strMsg.Length > 0, ", ", "") & "Falta(m) " & DiasParaOVencimento & " dia(s) para o Vencimento da Certidão Negativa da EMPRESA!" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(strMsg) Then
                MsgBox(Me.Page, strMsg)
            End If

        End If
        Return True
    End Function

    Public Function BuscaClassificacao(Optional ByVal delay As Boolean = False) As Boolean
        SessaoRecuperaNotaFiscal()
        If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL Then
            Session("MsgControle" & HID.Value) = "Não é possivel haver classificação de produto para notas globais."
            Return False
        End If
        If objNotaFiscal.CodigoRomaneio > 0 AndAlso Not objNotaFiscal.Romaneio.Processo = "NOTA FISCAL" Then
            Session("MsgControle" & HID.Value) = "A nota já possui um romaneio / classificação."
            Return False
        End If

        Dim validaRomaneio As Boolean = False

        If objNotaFiscal.Itens.Count > 1 Then
            For Each item In objNotaFiscal.Itens
                If item.Produto.Agrupar = "S" Then
                    validaRomaneio = True
                    Exit For
                End If
            Next
        End If

        If validaRomaneio Then Return False

        Session("ProcurandoClassificacao" & HID.Value) = True

        Popup.ConsultaDeNotaFiscalXClassificacao(Me, "objClassificacaoNXI" & HID.Value, "txtPrimeiraPesagem", delay, 1000)

        ucNotaFiscalXClassificacao.LimparCampos()
        ucNotaFiscalXClassificacao.BindGridView()
        ucNotaFiscalXClassificacao.CarregarPercentualPadraoClassificacao(objNotaFiscal.Quantidade, objNotaFiscal.Pedido.Itens(0).Produto.Codigo, False)

        Return True
    End Function

    Private Function CarregarItensComSaldo(ByVal ListaProdutos As [Lib].Negocio.SaldoPedido2015, ByRef erroMsg As String, ByRef MsgAlerta As String) As Boolean
        Dim Seq As Integer = 0
        Dim ind As Integer = 1
        Dim ValorUnitarioOficial As Decimal
        Dim ValorUnitarioMoeda As Decimal

        objNotaFiscal.CarregandoItens = True
        objNotaFiscal.Itens.Clear()

        Dim ListaDeProdutosSelecionados = ListaProdutos.Itens.Where(Function(s) s.Selecionado).ToList

        '**************************************************************************************************
        '************************* 1 - Carrega Item da Nota Fiscal ****************************************
        '**************************************************************************************************
        For Each row As [Lib].Negocio.SaldoPedido2015 In ListaDeProdutosSelecionados
            If ind > 0 Then
                If row.CodigoProduto = ListaDeProdutosSelecionados(ind - 1).CodigoProduto Then
                    Seq += 1
                Else
                    Seq = 0
                End If
            End If
            ind += 1

            '**************************************************************************************************
            '************************* 2 - Carrega Valor unitario Oficial / Moeda    **************************
            '**************************************************************************************************
            'Qual o Indice
            objNotaFiscal.VerificarIndice() 'obsoleto?????

            Try

                If row.XmlvUnCom > 0 Then
                    ValorUnitarioOficial = row.XmlvUnCom
                    ValorUnitarioMoeda = row.UnitarioMoeda
                Else

                    'Nao Tem Qtde na nota
                    If Not objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                        ValorUnitarioOficial = 0
                        ValorUnitarioMoeda = 0
                        'DEPOSITO NAO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS And Not objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = row.UnitarioOficial
                        ValorUnitarioMoeda = row.UnitarioMoeda
                        'DEPOSITO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS And objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.VlrNotaOficialDepositoBruto / row.QtdeEntregueFiscalDeposito, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.VlrNotaMoedaDepositoBruto / row.QtdeEntregueFiscalDeposito, 10, MidpointRounding.AwayFromZero)
                        'GLOBAL NAO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And Not objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.SaldoValorOficialGlobalDireto / row.SaldoQtdeGlobalFiscal, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.SaldoValorMoedaGlobalDireto / row.SaldoQtdeGlobalFiscal, 10, MidpointRounding.AwayFromZero)
                        'GLOBAL É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.VlrNotaOficialGlobalBruto / row.QtdeEntregueFiscalGlobal, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.VlrNotaMoedaGlobalBruto / row.QtdeEntregueFiscalGlobal, 10, MidpointRounding.AwayFromZero)
                        'REMESSA NAO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And Not objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.SaldoValorOficialRemessa / row.SaldoQtdeRemessaFiscal, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.SaldoValorMoedaRemessa / row.SaldoQtdeRemessaFiscal, 10, MidpointRounding.AwayFromZero)
                        'REMESSA É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.VlrNotaOficialRemessaBruto / row.QtdeEntregueFiscalRemessa, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.VlrNotaMoedaRemessaBruto / row.QtdeEntregueFiscalRemessa, 10, MidpointRounding.AwayFromZero)
                        'AFIXAR NAO É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR And Not objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = row.UnitarioOficial
                        ValorUnitarioMoeda = row.UnitarioMoeda
                        'AFIXAR É DEVOLUCAO
                    ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR And objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round((row.VlrNotaOficialAFixarBruto - row.VlrFixacaoOficial) / (row.QtdeEntregueFiscalAFixar - row.QtdeFixacao), 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round((row.VlrNotaMoedaAFixarBruto - row.VlrFixacaoMoeda) / (row.QtdeEntregueFiscalAFixar - row.QtdeFixacao), 10, MidpointRounding.AwayFromZero)
                    ElseIf objNotaFiscal.SubOperacao.Devolucao Then
                        ValorUnitarioOficial = Math.Round(row.VlrNotaOficialDiretaBruto / row.QtdeEntregueFiscalDireta, 10, MidpointRounding.AwayFromZero)
                        ValorUnitarioMoeda = Math.Round(row.VlrNotaMoedaDiretaBruto / row.QtdeEntregueFiscalDireta, 10, MidpointRounding.AwayFromZero)
                    Else

                        If row.SaldoValorOficialGlobalDireto > 0 AndAlso row.SaldoQtdeDiretoFiscal > 0 Then
                            ValorUnitarioOficial = Math.Round(row.SaldoValorOficialGlobalDireto / row.SaldoQtdeDiretoFiscal, 10, MidpointRounding.AwayFromZero)
                        End If

                        If row.SaldoValorMoedaGlobalDireto > 0 AndAlso row.SaldoQtdeDiretoFiscal > 0 Then

                            'DE INDEXADOR DO PEDIDO FOR FIXO DEVE PEGAR O UNITÁRIO EM DÓLAR DO PEDIDO - FURLAN - 28/01/2024
                            If objNotaFiscal.Pedido.CodigoIndexador = 99 OrElse objNotaFiscal.Pedido.IndexadorFixo Then
                                ValorUnitarioMoeda = row.UnitarioMoeda
                            Else
                                ValorUnitarioMoeda = Math.Round(row.SaldoValorMoedaGlobalDireto / row.SaldoQtdeDiretoFiscal, 10, MidpointRounding.AwayFromZero)
                            End If
                        End If

                    End If
                End If
            Catch ex As Exception
                erroMsg &= "Erro ao calcular o unitario do produto: " & row.CodigoProduto & " - " & row.NomeProduto & " \n "
            End Try

            '********************
            '**** 2 - FIM *******
            '********************

            Dim objItemNF As New [Lib].Negocio.NotaFiscalXItem(objNotaFiscal)

            objItemNF.CodigoProduto = row.CodigoProduto
            objItemNF.Sequencia = Seq

            'Levar informarção do InfaDProd do Produto para NFE caso tenha - Furlan - 21-12-2022
            If objItemNF.Produto.InfaDProd.Length > 0 Then objItemNF.ObservacoesDoProduto = objItemNF.Produto.InfaDProd

            objItemNF.CodigoPedido = ListaProdutos.CodigoPedido

            objItemNF.CodigoOperacao = objNotaFiscal.CodigoOperacao
            objItemNF.CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao

            objItemNF.PesoQuantidade = objItemNF.Produto.PesoQuantidade
            objItemNF.Lote = row.Lote
            objItemNF.Classificacao = row.Classificacao

            'objItemNF.CodigoEmbalagem = row.CodigoEmbalagem
            'objItemNF.CodigoEmbalagemIndea = row.EmbalagemIndea
            objItemNF.CodigoEmbalagem = objItemNF.Produto.CodigoEmbalagem
            objItemNF.CodigoEmbalagemIndea = objItemNF.Produto.Embalagem.EmbalagemIndea

            objItemNF.CodigoTipoDeEmbalagem = row.TipoDeEmbalagem
            objItemNF.CapacidadeEmbalagem = row.CapacidadeEmbalagem

            '**********************************************************************************************************************************
            '***************************************** 3 - SALDOS  ****************************************************************************
            '**********************************************************************************************************************************
            'Verifica o saldo do pedido de acordo com a Operacao da Nota

            'GLOBAL É DEVOLUCAO
            If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And objNotaFiscal.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalGlobal - row.QtdeEntregueFiscalRemessa
                objItemNF.SaldoPedidoFisico = row.QtdeEntregueFiscalGlobal - row.QtdeEntregueFiscalRemessa
                objItemNF.SaldoValorOficial = row.VlrNotaOficialGlobalBruto - row.VlrNotaOficialRemessaBruto
                objItemNF.SaldoValorMoeda = row.VlrNotaMoedaGlobalBruto - row.VlrNotaMoedaRemessaBruto
                'GLOBAL NAO É DEVOLUCAO
            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And Not objNotaFiscal.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = row.SaldoQtdeGlobalFiscal
                objItemNF.SaldoPedidoFisico = row.SaldoQtdeGlobalFiscal
                objItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto
                objItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto
                'REMESSA É DEVOLUCAO
            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And objNotaFiscal.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalRemessa
                objItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoRemessa
                objItemNF.SaldoValorOficial = row.VlrNotaOficialRemessaBruto
                objItemNF.SaldoValorMoeda = row.VlrNotaMoedaRemessaBruto
                'REMESSA NAO É DEVOLUCAO
            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And Not objNotaFiscal.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = row.SaldoQtdeRemessaFiscal
                objItemNF.SaldoPedidoFisico = row.SaldoQtdeRemessaFisica
                objItemNF.SaldoValorOficial = row.SaldoValorOficialRemessa
                objItemNF.SaldoValorMoeda = row.SaldoValorMoedaRemessa
                'AFIXAR É DEVOLUCAO
            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR And objNotaFiscal.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalAFixar - row.QtdeFixacao
                objItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoAFixar - row.QtdeFixacao
                'Criado um novo campo na StoredProcedure spSaldoPedido para trazer o valor
                'total da fixação, mas correspondente ao unitário da NF. 
                'O campo VlrfixacaoNF será utilizado quando for menor  do que o campo VlrfixacaoOficial senão este último será, para que o valor da devolução fique correto.
                objItemNF.SaldoValorOficial = row.VlrNotaOficialAFixarBruto - IIf(row.VlrFixacaoNF < row.VlrFixacaoOficial, row.VlrFixacaoNF, row.VlrFixacaoOficial)
                objItemNF.SaldoValorMoeda = row.VlrNotaMoedaAFixarBruto - row.VlrFixacaoMoeda

                'Devolução utilizada ou para Reajuste de unitário ou para devolução de valor.
            ElseIf objNotaFiscal.SubOperacao.Devolucao And row.Tipo = 1 AndAlso Not objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                objItemNF.SaldoPedidoFiscal = Math.Abs(row.SaldoQtdeDiretoFiscal)
                objItemNF.SaldoPedidoFisico = Math.Abs(row.SaldoQtdeDiretoFisica)
                objItemNF.SaldoValorOficial = Math.Abs(row.VlrNotaOficialDiretaBruto)
                objItemNF.SaldoValorMoeda = Math.Abs(row.VlrNotaMoedaDiretaBruto)
            ElseIf objNotaFiscal.SubOperacao.Devolucao And row.Tipo = 1 Then
                objItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalDireta
                objItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoDireta
                objItemNF.SaldoValorOficial = row.VlrNotaOficialDiretaBruto
                objItemNF.SaldoValorMoeda = row.VlrNotaMoedaDiretaBruto
            ElseIf objNotaFiscal.SubOperacao.Devolucao And row.Tipo = 3 Then
                objItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalDeposito
                objItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoDeposito
                objItemNF.SaldoValorOficial = row.VlrNotaOficialDepositoBruto
                objItemNF.SaldoValorMoeda = row.VlrNotaMoedaDepositoBruto
            Else
                'Verificar ainda pode nao estar contemplando tudo
                If row.QtdeProgramada = 0 And row.QtdeProgramadaComercializacao > 0 Then
                    objItemNF.SaldoPedidoFiscal = row.SaldoQtdeComercializacao
                    objItemNF.SaldoPedidoFisico = row.SaldoQtdeComercializacao
                    objItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto
                    objItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto
                Else
                    objItemNF.SaldoPedidoFiscal = row.SaldoQtdeDiretoFiscal
                    objItemNF.SaldoPedidoFisico = row.SaldoQtdeDiretoFisica
                    objItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto
                    objItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto
                End If
            End If

            '********************************************************************************************************************************************
            '***************************************** 3.1 - SALDOS vs ESTOQUE **************************************************************************
            '********************************************************************************************************************************************
            'Verifica o saldo do pedido de acordo com a Operacao da Nota não é maior que o saldo do produto em Estoque
            'Se o saldo Quantidade do produto em Estoque "Sem Embalagem" for menor que o saldo do pedido o saldo do pedido assume a Quantidade em Estoque
            'Acrescentei a classe CONTAEORDEM pois o estoque só deve ser checado para a VENDA OU COMPRA A ORDEM - FURLAN - 22/03/16 

            If objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL _
            And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.CONTAEORDEM _
            And objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida _
            And objItemNF.Produto.ControlarEstoque Then

                Dim SaldoProdutoEstoque As New [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem(objItemNF.CodigoProduto)
                SaldoProdutoEstoque.CarregarResumoSaldoEmEstoque(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.SubOperacao)

                If objItemNF.SubOperacao.EstoqueFisico Or objItemNF.SubOperacao.EstoqueFiscal Then

                    If Not objItemNF.SubOperacao.Devolucao AndAlso objItemNF.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso objItemNF.SubOperacao.EstoqueFiscal AndAlso row.Produto.EstoqueMinimo > 0 Then

                        If (SaldoProdutoEstoque.SaldoFiscal - objItemNF.SaldoPedidoFiscal) < row.Produto.EstoqueMinimo Then

                            Dim ItemEstoqueMinimo As DataRow = CType(Session("objEstoqueMinimo" & HID.Value), DataTable).NewRow()

                            ItemEstoqueMinimo("Produto") = row.Produto.Codigo

                            If objNotaFiscal.Empresa.Empresa.UsarDescricaoProduto Then
                                ItemEstoqueMinimo("Nome") = row.Produto.Nome & "-" & row.Produto.Descricao
                            Else
                                ItemEstoqueMinimo("Nome") = row.Produto.Nome
                            End If

                            ItemEstoqueMinimo("EstoqueMinimo") = row.Produto.EstoqueMinimo.ToString("N4")
                            ItemEstoqueMinimo("Faturando") = objItemNF.SaldoPedidoFiscal
                            ItemEstoqueMinimo("Saldo") = SaldoProdutoEstoque.SaldoFiscal - objItemNF.SaldoPedidoFiscal
                            CType(Session("objEstoqueMinimo" & HID.Value), DataTable).Rows.Add(ItemEstoqueMinimo)
                        End If
                    End If

                    If SaldoProdutoEstoque.Count = 0 Then
                        objItemNF.SaldoValorOficial = 0
                        objItemNF.SaldoValorMoeda = 0
                        objItemNF.SaldoPedidoFiscal = 0
                        objItemNF.SaldoPedidoFisico = 0
                        erroMsg &= "Produto " & row.CodigoProduto & " - " & row.NomeProduto & " sem estoque. \n"
                    Else
                        Dim faltaSaldo As Boolean = False

                        If objItemNF.SubOperacao.EstoqueFisico AndAlso objItemNF.SaldoPedidoFisico > SaldoProdutoEstoque.SaldoFisico Then
                            MsgAlerta &= "O Produto " & row.CodigoProduto & " - " & row.NomeProduto & " tem em Estoque FISICO -> " & SaldoProdutoEstoque(0).SaldoFisico.ToString() & " que é menor que o saldo do Pedido -> " & objItemNF.SaldoPedidoFisico & ". Será liberado apenas o que tem em Estoque. \n"

                            faltaSaldo = True
                        End If

                        If objItemNF.SubOperacao.EstoqueFiscal AndAlso objItemNF.SaldoPedidoFiscal > SaldoProdutoEstoque.SaldoFiscal Then
                            MsgAlerta &= "O Produto " & row.CodigoProduto & " - " & row.NomeProduto & " tem em Estoque FISCAL -> " & SaldoProdutoEstoque(0).SaldoFiscal.ToString() & " que é menor que o saldo do Pedido -> " & objItemNF.SaldoPedidoFiscal & ". Será liberado apenas o que tem em Estoque. \n"

                            faltaSaldo = True
                        End If

                        If faltaSaldo Then
                            objItemNF.SaldoPedidoFiscal = SaldoProdutoEstoque.SaldoFiscal
                            objItemNF.SaldoPedidoFisico = SaldoProdutoEstoque.SaldoFisico
                            If ValorUnitarioOficial > 0 Then
                                objItemNF.SaldoValorOficial = Math.Round(SaldoProdutoEstoque.SaldoFiscal * ValorUnitarioOficial, 2, MidpointRounding.AwayFromZero)
                                objItemNF.SaldoValorMoeda = Math.Round(SaldoProdutoEstoque.SaldoFiscal * ValorUnitarioMoeda, 2, MidpointRounding.AwayFromZero)
                            End If
                        End If
                    End If
                End If
            End If

            'Na Devolucao Apura o estoque do produto de acordo com o lote/classificacao que foi recebido
            'Acrescentei a classe CONTAEORDEM pois o estoque só deve ser checado para a VENDA OU COMPRA A ORDEM - FURLAN - 22/03/16 
            If Not objNotaFiscal.SubOperacao.Devolucao _
         And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL _
         And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.CONTAEORDEM _
         And objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida _
         And objItemNF.Produto.ControlarEstoque _
         And (objItemNF.Produto.ControlarEmbalagem Or objItemNF.Produto.ControlarLote) _
         And objNotaFiscal.SubOperacao.EstoqueFiscal _
        Then
                Dim SaldoProdutoEstoque As New [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem(objItemNF.CodigoProduto)
                SaldoProdutoEstoque.CarregaSaldoDisponivelSaidaLoteClassificao(objNotaFiscal.SubOperacao, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objItemNF.Lote, objItemNF.Classificacao, objItemNF.CodigoEmbalagem, objItemNF.CodigoTipoDeEmbalagem, objItemNF.CapacidadeEmbalagem)

                If SaldoProdutoEstoque.Count = 0 Then
                    objItemNF.SaldoPedidoFiscal = 0
                    objItemNF.SaldoPedidoFisico = 0
                    objItemNF.SaldoValorOficial = 0
                    objItemNF.SaldoValorMoeda = 0
                    erroMsg &= "Produto não pode ser liberado, saldo em Estoque = 0. \n" & vbCrLf
                Else
                    Dim faltaSaldo As Boolean = False

                    If objItemNF.SubOperacao.EstoqueFiscal AndAlso objItemNF.SaldoPedidoFiscal > SaldoProdutoEstoque(0).SaldoFiscal Then
                        MsgAlerta &= "O Produto " & row.CodigoProduto & " - " & row.NomeProduto & " tem em Estoque FISCAL -> " & SaldoProdutoEstoque(0).SaldoFiscal.ToString() & " que é menor que o saldo do Pedido -> " & SaldoProdutoEstoque(0).SaldoFiscal & ". Será liberado apenas o que tem em Estoque. \n"

                        faltaSaldo = True
                    End If

                    If objItemNF.SubOperacao.EstoqueFisico AndAlso objItemNF.SaldoPedidoFisico > SaldoProdutoEstoque(0).SaldoFisico Then
                        MsgAlerta &= "O Produto " & row.CodigoProduto & " - " & row.NomeProduto & " tem em Estoque FISICO -> " & SaldoProdutoEstoque(0).SaldoFiscal.ToString() & " que é menor que o saldo do Pedido -> " & SaldoProdutoEstoque(0).SaldoFiscal & ". Será liberado apenas o que tem em Estoque. \n"

                        faltaSaldo = True
                    End If

                    If faltaSaldo Then
                        objItemNF.SaldoPedidoFiscal = SaldoProdutoEstoque(0).SaldoFiscal
                        objItemNF.SaldoPedidoFisico = SaldoProdutoEstoque(0).SaldoFisico
                        objItemNF.SaldoValorOficial = Math.Round(SaldoProdutoEstoque(0).SaldoFiscal * ValorUnitarioOficial, 2, MidpointRounding.AwayFromZero)
                        objItemNF.SaldoValorMoeda = Math.Round(SaldoProdutoEstoque(0).SaldoFiscal * ValorUnitarioMoeda, 2, MidpointRounding.AwayFromZero)
                    End If
                End If
            End If
            '********************
            '**** 3 - FIM *******
            '********************

            If row.XmlqCom > 0 Then
                objItemNF.PesoFiscal = row.XmlqCom

                If objItemNF.SubOperacao.QuantidadeFisico Then
                    objItemNF.QuantidadeFisica = row.XmlqCom
                Else
                    objItemNF.QuantidadeFisica = 0
                End If

                objItemNF.QuantidadeFiscal = row.XmlqCom

                If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS _
                        OrElse objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR Then
                    row.QtdeProgramada = objItemNF.QuantidadeFiscal
                End If
            Else
                If Not objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                    objItemNF.PesoFiscal = 0
                    objItemNF.QuantidadeFisica = 0
                    objItemNF.QuantidadeFiscal = 0
                    objNotaFiscal.Quantidade = 0
                Else
                    objItemNF.PesoFiscal = objItemNF.SaldoPedidoFiscal
                    objItemNF.QuantidadeFiscal = objItemNF.SaldoPedidoFiscal
                    objItemNF.QuantidadeFisica = IIf(objItemNF.SubOperacao.QuantidadeFisico, objItemNF.SaldoPedidoFiscal, 0)
                    objNotaFiscal.Quantidade += objItemNF.QuantidadeFiscal
                End If

            End If

            If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                objItemNF.UnitarioMoeda = IIf(ValorUnitarioMoeda < 0, 0, ValorUnitarioMoeda)
                objItemNF.Unitario = IIf(ValorUnitarioOficial < 0, 0, ValorUnitarioOficial)
            Else
                If objNotaFiscal.Pedido.CodigoIndexador = 99 OrElse objNotaFiscal.Pedido.IndexadorFixo Then
                    objItemNF.Unitario = Math.Round(objNotaFiscal.Pedido.IndiceFixado * ValorUnitarioMoeda, 10, MidpointRounding.AwayFromZero)
                Else
                    Dim indiceDolar As Decimal = New Cotacao(objNotaFiscal.Pedido.CodigoIndexador, objNotaFiscal.DataNota).Indice 'PTAX

                    objItemNF.Unitario = Math.Round(indiceDolar * ValorUnitarioMoeda, 10, MidpointRounding.AwayFromZero)
                End If

                objItemNF.UnitarioMoeda = IIf(ValorUnitarioMoeda < 0, 0, ValorUnitarioMoeda)

            End If

            If row.XmlvProd > 0 Then
                objItemNF.ValorTotal = row.XmlvProd
                objItemNF.ValorTotalMoeda = 0
            Else
                'Quando a nota é somente valor, ela nao tem cotacao e o valor é somente em reais
                If Not objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                    'Pendente alimentar o valor
                    objItemNF.ValorTotal = objItemNF.SaldoValorOficial
                    objItemNF.ValorTotalMoeda = objItemNF.SaldoValorMoeda
                Else
                    objItemNF.ValorTotal = Math.Round(objItemNF.QuantidadeFiscal * objItemNF.Unitario, 2, MidpointRounding.AwayFromZero)
                    objItemNF.ValorTotalMoeda = Math.Round(objItemNF.QuantidadeFiscal * objItemNF.UnitarioMoeda, 2, MidpointRounding.AwayFromZero)
                End If
            End If

            objItemNF.CodigoOperacao = objNotaFiscal.CodigoOperacao
            objItemNF.CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao
            objItemNF.Encargos = Nothing

            '**********************************************************************************************************************************
            '*********************************************** OBSERVACOES DO PRODUTO ***********************************************************
            '**********************************************************************************************************************************
            If objItemNF.EmbalagemProduto IsNot Nothing AndAlso objItemNF.EmbalagemProduto.PesoVariavel Then
                objItemNF.QuantidadeDeEmbalagem = objItemNF.QuantidadeFiscal / objItemNF.ProdutoLoteClassificacao.PesoSaco
                objItemNF.ObservacoesDoProduto &= objItemNF.QuantidadeDeEmbalagem & " " & objItemNF.Embalagem.Descricao & " " & objItemNF.TipoDeEmbalagem.Descricao & " " & IIf(objItemNF.EmbalagemProduto.PesoVariavel, objItemNF.ProdutoLoteClassificacao.PesoSaco, objItemNF.CapacidadeEmbalagem & " " & objItemNF.Produto.Unidade) & " / " & objItemNF.Produto.DescricaoTecnica
            ElseIf objItemNF.CapacidadeEmbalagem > 0 Then
                objItemNF.QuantidadeDeEmbalagem = objItemNF.QuantidadeFiscal / objItemNF.CapacidadeEmbalagem
                objItemNF.ObservacoesDoProduto &= objItemNF.QuantidadeDeEmbalagem & " " & objItemNF.Embalagem.Descricao & " " & objItemNF.TipoDeEmbalagem.Descricao & " " & objItemNF.CapacidadeEmbalagem & " " & objItemNF.Produto.Unidade & " / " & objItemNF.Produto.DescricaoTecnica
            End If

            If objItemNF.Lote.Length > 0 Then
                Dim L As New [Lib].Negocio.Lote(objItemNF.Lote, objItemNF.CodigoProduto)
                If L.Tipo = 2 Then
                    objItemNF.ObservacoesDoProduto &= IIf(objItemNF.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & objItemNF.Lote & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
                Else
                    objItemNF.ObservacoesDoProduto &= IIf(objItemNF.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & objItemNF.Lote & " GER " & L.Germinacao & IIf(L.Pureza = 0, "", " PUR " & L.Pureza) & " Peneira/Classif. " & objItemNF.Classificacao & " Renasem " & L.Renasem & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
                End If
            End If

            If Not String.IsNullOrWhiteSpace(row.XmlinfAdProd) Then
                objItemNF.ObservacoesDoProduto &= row.XmlinfAdProd
            End If
            '**********************************************************************************************************************************
            '**********************************************************************************************************************************

            If objNotaFiscal.SubOperacao.QuantidadeFiscal And objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira And objItemNF.SaldoValorMoeda <= 0 And objItemNF.SubOperacao.Classe <> eClassesOperacoes.AFIXAR Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & " esta com saldo 0 ou negativo não pode ser usado para emissão de Nota Fiscal. Saldo: " & objItemNF.SaldoValorMoeda & " Pedido em moeda Extrangeira. \n"
            ElseIf IIf(objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, objItemNF.SaldoValorOficial, objItemNF.SaldoValorMoeda) <= 0 _
                And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS _
                And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR _
                And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.FRETES _
                And objNotaFiscal.Operacao.CodigoClasse <> eClassesOperacoes.COMPRAS.ToString _
                And Not objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE ICMS") _
                And Not objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE IPI") _
                And objNotaFiscal.SubOperacao.QuantidadeFiscal _
                And Not objNotaFiscal.SubOperacao.NotaDebito _
                And Not objNotaFiscal.SubOperacao.NotaCredito _
                And Not objNotaFiscal.SubOperacao.Devolucao Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & " esta com Saldo de Valor 0 ou negativo e não pode ser usado para emissão de Nota Fiscal. Saldo: " & objItemNF.SaldoValorOficial & " \n"
            ElseIf row.QtdeProgramadaComercializacao < 0 Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & " esta com saldo negativo na Quantidade Programada para Comercialização e não pode ser usado para emissão de Nota Fiscal. Saldo: " & row.QtdeProgramadaComercializacao & " \n"
            ElseIf objItemNF.Encargos Is Nothing OrElse objItemNF.Encargos.Count = 0 Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", não tem encargos cadastrados na Operação:" & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " \n"
            ElseIf objItemNF.CFOP = 0 Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", Operação:" & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " está sem CFOP informada \n"
            ElseIf objItemNF.CFOP < 5000 And objNotaFiscal.EntradaSaida = eEntradaSaida.Saida Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", Operação " & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " com CFOP " & objItemNF.CFOP & " não pode ser usada na saída. \n"
            ElseIf objItemNF.CFOP > 5000 And objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", Operação " & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " com CFOP " & objItemNF.CFOP & " não pode ser usada na entrada. \n"
            ElseIf objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS _
                And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR _
                And Not objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE ICMS") _
                And Not objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE IPI") _
                And Not objNotaFiscal.SubOperacao.Devolucao And row.QtdeProgramada < objItemNF.QuantidadeFiscal Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & " esta com saldo da quantidade Programada menor que a quantidade da Nota Fiscal. Saldo Programado: " & row.QtdeProgramada & " - Quantidade da nota: " & objItemNF.QuantidadeFiscal & " \n"
            ElseIf objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS _
                And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR _
                And objItemNF.SaldoValorOficial < objItemNF.ValorTotal _
                And Not objNotaFiscal.SubOperacao.Devolucao = True Then
                erroMsg &= "O pedido: " & objNotaFiscal.Pedido.Codigo & " com o produto: " & objItemNF.Produto.Nome & " esta com saldo de valor fiscal do pedido menor que o valor da nota Fiscal. Saldo Fiscal: " & objItemNF.SaldoValorOficial & " - Valor da nota: " & objItemNF.ValorTotal & " \n"
            ElseIf objItemNF.Encargos.EncProduto.SituacaoTributariaPISCOFINS = 0 Then
                erroMsg &= "Situação Tributária PISCOFINS não cadastrada para o Produto " & objItemNF.CodigoProduto & "-" & objItemNF.Produto.Nome & ", Operação " & objNotaFiscal.CodigoOperacao & "-" & objNotaFiscal.CodigoSubOperacao & " " & objNotaFiscal.SubOperacao.Descricao
            ElseIf (Not objNotaFiscal.SubOperacao.QuantidadeFiscal AndAlso (Math.Abs(objItemNF.SaldoValorOficial) > 0 Or objNotaFiscal.SubOperacao.Devolucao)) _
            OrElse objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE ICMS") _
            OrElse objNotaFiscal.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE IPI") _
            OrElse objNotaFiscal.SubOperacao.Descricao.ToString.Contains("TRANSFERENCIA DE CREDITO DE ICMS") _
            OrElse (objNotaFiscal.SubOperacao.Devolucao = True AndAlso objNotaFiscal.SubOperacao.QuantidadeFiscal AndAlso objItemNF.SaldoValorOficial > 0 AndAlso objItemNF.ValorTotal > 0) _
            OrElse (objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso objNotaFiscal.SubOperacao.Devolucao = True AndAlso objItemNF.SaldoValorOficial > 0) _
            OrElse (row.QtdeProgramada = 0 AndAlso row.QtdeProgramadaComercializacao > 0 AndAlso ((objNotaFiscal.SubOperacao.Devolucao AndAlso row.QtdeComercializacaoEntregue > 0) OrElse (objNotaFiscal.SubOperacao.Devolucao = False AndAlso row.SaldoQtdeComercializacao))) _
            OrElse (objItemNF.QuantidadeFiscal > 0 OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL OrElse objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.FRETES) _
            OrElse objNotaFiscal.SubOperacao.NotaDebito _
            OrElse objNotaFiscal.SubOperacao.NotaCredito _
            AndAlso objItemNF.Encargos.Count > 0 Then

                objNotaFiscal.Itens.Add(objItemNF)

            ElseIf (Not objNotaFiscal.SubOperacao.NotaDebito OrElse Not objNotaFiscal.SubOperacao.NotaCredito) AndAlso objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.OUTRAS Then
                erroMsg &= "O item não pôde ser adicionado, verifique as configurações de Operações, Operações X Encargos, Saldo no Extrato de Pedido e outros."
            End If
        Next

        Dim itensAgrupadoPorProduto = From item In objNotaFiscal.Itens
                                      Group item By item.CodigoProduto Into Grupo = Group
                                      Select New With {
                                          .CodigoProduto = CodigoProduto,
                                          .SaldoPedidoFiscal = Grupo.Sum(Function(x) x.SaldoPedidoFiscal),
                                          .SaldoPedidoFisico = Grupo.Sum(Function(x) x.SaldoPedidoFisico),
                                          .SaldoValorMoeda = Grupo.Sum(Function(x) x.SaldoValorMoeda),
                                          .SaldoValorOficial = Grupo.Sum(Function(x) x.SaldoValorOficial),
                                          .Itens = Grupo.ToList()
                                      }

        For Each itemNF In objNotaFiscal.Itens

            itemNF.CodigoOperacao = objNotaFiscal.CodigoOperacao
            itemNF.CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao

            itemNF.PesoFiscal = itemNF.QuantidadeFiscal

            If itemNF.SubOperacao.QuantidadeFisico Then

                itemNF.QuantidadeFisica = itemNF.QuantidadeFiscal

                If objNotaFiscal.CriarRomaneio = False Then

                    Dim pesoDoItem = itemNF.Produto.UnidadesDeComercializacao.Where(Function(p) p.CodigoUnidade = itemNF.Produto.Unidade).FirstOrDefault().FatorConversao
                    If pesoDoItem > 0 Then
                        'objNotaFiscal.Itens(0).PesoFiscal = (objNotaFiscal.Itens(0).SaldoPedidoFiscal * pesoDoItem)
                        itemNF.PesoFiscal = itemNF.QuantidadeFiscal * pesoDoItem
                        itemNF.PesoLiquido = itemNF.QuantidadeFiscal * pesoDoItem
                        itemNF.PesoBruto = itemNF.QuantidadeFiscal * pesoDoItem
                    End If

                End If

            End If

            For Each itemSaldo In itensAgrupadoPorProduto.Where(Function(x) x.CodigoProduto = itemNF.CodigoProduto)

                itemNF.SaldoPedidoFiscal = itemSaldo.SaldoPedidoFiscal
                itemNF.SaldoPedidoFisico = itemSaldo.SaldoPedidoFisico
                itemNF.SaldoValorMoeda = itemSaldo.SaldoValorMoeda
                itemNF.SaldoValorOficial = itemSaldo.SaldoValorOficial


                itemSaldo.SaldoPedidoFiscal -= itemNF.PesoFiscal
                itemSaldo.SaldoPedidoFisico -= itemNF.PesoFiscal
                itemSaldo.SaldoValorMoeda -= itemNF.ValorTotal
                itemSaldo.SaldoValorOficial -= itemNF.ValorTotalMoeda

            Next

            If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then

                If objNotaFiscal.Pedido.CodigoIndexador = 99 OrElse objNotaFiscal.Pedido.IndexadorFixo Then
                    itemNF.Unitario = Math.Round(objNotaFiscal.Pedido.IndiceFixado * ValorUnitarioMoeda, 10, MidpointRounding.AwayFromZero)
                Else
                    Dim indiceDolar As Decimal = New Cotacao(objNotaFiscal.Pedido.CodigoIndexador, objNotaFiscal.DataNota).Indice 'PTAX

                    itemNF.Unitario = Math.Round(indiceDolar * ValorUnitarioMoeda, 10, MidpointRounding.AwayFromZero)
                End If

                itemNF.UnitarioMoeda = IIf(ValorUnitarioMoeda < 0, 0, ValorUnitarioMoeda)

            End If

            Try

                If itemNF.Encargos Is Nothing Then

                    'Poder instanciar os encargos dos itens que faltam

                End If

            Catch ex As Exception

                Throw New Exception(ex.Message)

            End Try

        Next

        objNotaFiscal.CarregandoItens = False
        objNotaFiscal.AtualizaTotais()
        'objNotaFiscal.CarregandoItens = False

        'Importação do xml de notas de entrada
        Dim ObsTemp As String = objNotaFiscal.Observacoes
        objNotaFiscal.AtualizarObservacoes()
        objNotaFiscal.Observacoes &= ObsTemp

        For Each itemNota In objNotaFiscal.Itens
            If itemNota.Produto.ControlarLote = False AndAlso Not itemNota.Lotes Is Nothing AndAlso itemNota.Lotes.Count > 0 Then
                itemNota.Lotes = Nothing
            End If
        Next

        SessaoSalvaNotaFiscal()

        '**********************
        '**** 1 - FIM *********
        '**********************
        Session.Remove("objItensPedidoSelecionadosNXI" & HID.Value)

        Return True
    End Function

#End Region

#Region "Sessão"

    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

#End Region

#Region "Botões"

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmpresa.Click
        Try
            SessaoRecuperaNotaFiscal()
            LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)

            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me, "objEmpresaNXI" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPedido.Click
        Try
            ConsultaPedidos(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPlaca_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPlaca.Click
        Try
            BuscarPlaca()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnObservacoesFiscais_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo" & HID.Value) = "NotaFiscalXItens"
            HttpContext.Current.Session("Observacoes" & HID.Value) = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesFiscais.Text)
            ucConsultaObservacoes.BindGridView()
            Popup.ConsultaDeObservacoes(Me, "objObservacoesNXI" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEmbarque_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            If ddlEmbarque.SelectedIndex > 0 Then
                Dim dep() As String = ddlEmbarque.SelectedValue.ToString.Split("-")
                objNotaFiscal.CodigoLocalEmbarque = dep(0)
                objNotaFiscal.EndLocalEmbarque = dep(1)
            Else
                objNotaFiscal.CodigoLocalEmbarque = String.Empty
            End If
            SessaoSalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnEmbarque_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeClientes(Me, "objEmbarqueNXI" & HID.Value, txtNome.ClientID)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtObservacoesFiscais_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        objNotaFiscal.Observacoes = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesFiscais.Text)
        txtObservacoesFiscais.Text = objNotaFiscal.Observacoes
        SessaoSalvaNotaFiscal()
    End Sub

    Protected Sub txtObservacoesInternas_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            objNotaFiscal.ObservacoesControleInterno = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesInternas.Text)
            SessaoSalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub txtDataDeEmissao_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        If Not IsDate(txtDataDeEmissao.Text) Then
            MsgBox(Me.Page, "Data de Emissão inválida!")
            txtDataDeEmissao.Text = Today.ToString("dd/MM/yyyy")
            Exit Sub
        End If

        If CDate(txtDataDeEmissao.Text) > Now.Date Then
            txtDataDeEmissao.Text = Today.ToString("dd/MM/yyyy")
            MsgBox(Me.Page, "Data de Emissão não pode ser maior que a data atual")
        Else

            SessaoRecuperaNotaFiscal()

            If Not objNotaFiscal Is Nothing Then
                objNotaFiscal.DataNota = CDate(txtDataDeEmissao.Text)
            End If

            SessaoSalvaNotaFiscal()
        End If

    End Sub

    Protected Sub txtDataCarregamento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        If Not IsDate(txtDataCarregamento.Text) Then
            MsgBox(Me.Page, "Data de Entrada inválida!")
            txtDataCarregamento.Text = Today.ToString("dd/MM/yyyy")
            Exit Sub
        End If

        If CDate(txtDataCarregamento.Text) > Now.Date Then
            txtDataCarregamento.Text = Today.ToString("dd/MM/yyyy")
            MsgBox(Me.Page, "Data da Nota não pode ser maior que a data atual")
        Else
            SessaoRecuperaNotaFiscal()

            objNotaFiscal.Movimento = CDate(txtDataCarregamento.Text)

            SessaoSalvaNotaFiscal()

        End If

    End Sub

#End Region

End Class