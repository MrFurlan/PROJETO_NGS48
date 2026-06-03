Imports System.IO
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports Boleto2Net
Imports System.Globalization
Imports OfficeOpenXml
Imports System.Drawing
Imports OfficeOpenXml.Style
Imports System.Web.Services
Imports System.Net.Mime.MediaTypeNames
Imports System.Xml
Imports System.Web.Services.Description

Public NotInheritable Class Clipboard

End Class

Partial Class WFTitulo
    Inherits BasePage

#Region "Atributos / Propriedades"
    Dim ObjTitulo As NGS.Lib.Negocio.TituloV

    Dim Sql As String
    Dim Sqla As String
    Dim SqlArray As New ArrayList
    Dim Unidades As New ArrayList

    Dim DS As DataSet
    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date
    Dim Codigo As String
    Dim Descricao As String
    Dim Cliente As String
    Dim campo() As String
    Dim Endereco As String
    Dim Registro As Integer
    Dim TemRegistro As String
    Dim Valor As Decimal

    Dim Raz_Empresa As String
    Dim Raz_EndEmpresa As Integer
    Dim Raz_Conta As String
    Dim Raz_Cliente As String
    Dim Raz_EndCliente As String
    Dim Raz_UnidadeDeNegocio As String
    Dim Raz_ValorOficial As String
    Dim Raz_ValorMoeda As String
    Dim Raz_Historico As String
    Dim Raz_DebitoCredito As String

    Dim MoedaJuros As String
    Dim MoedaAcrescimos As String
    Dim MoedaDescontos As String
    Dim MoedaDeducoes As String

    Dim Mensagem As String
    Dim strJavaScript As String
    Dim objClient As [Lib].Negocio.Cliente
    Private objPedido As [Lib].Negocio.Pedido
    Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal



#End Region

#Region "Session"
    Private Sub SessaoSalvaTitulo()
        Session("ObjTitulo" + HID.Value) = ObjTitulo
    End Sub

    Private Sub SessaoRecuperaTitulo()
        If Not ObjTitulo Is Nothing Then Exit Sub
        If Session("ObjTitulo" + HID.Value) Is Nothing Then
            ObjTitulo = New NGS.Lib.Negocio.TituloV
        Else
            ObjTitulo = CType(Session("ObjTitulo" + HID.Value), NGS.Lib.Negocio.TituloV)
        End If
    End Sub

#End Region
    Private Sub SubstituirLabels()
        divCheque.Visible = False
        If HTabela.Value = "ContasAReceber" Then
            lblForm.Text = "Contas a Receber"
            lblCliente.Text = "Cliente:"
            lblCabCliFor.Text = "Dados do Cliente"
            lblEmpresa.Text = "Empresa Recebedora:"
            lblTipoPgtoRec.Text = "Tipo Receb.:"
            divBanco.Visible = False
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If IsConnect AndAlso Not IsPostBack Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Financeiro.aspx", eTitulo.Info)
                Exit Sub
            End If

            Dim RegistroParam As Integer

            If Not String.IsNullOrWhiteSpace(Request.QueryString("idTabela")) _
                AndAlso (Request.QueryString("idTabela").Equals("ContasAPagar") OrElse Request.QueryString("idTabela").Equals("ContasAReceber")) Then
                HTabela.Value = Request.QueryString("idTabela")
            ElseIf Not String.IsNullOrWhiteSpace(Request.QueryString("param")) Then
                Dim param As String() = Funcoes.Decifrar(Request.QueryString("param")).Split("-")
                HTabela.Value = param(0)
                RegistroParam = param(1)
            Else
                Response.Redirect("~/Financeiro.aspx")
            End If

            SubstituirLabels()

            If Funcoes.VerificaPermissao(HTabela.Value, "ACESSAR") Then
                If Funcoes.VerificaPermissao(HTabela.Value, "LIBERAR") Then
                    chkPrevisao.Visible = True
                    chkProvisao.Visible = True
                    chkBaixado.Visible = True
                End If

                CargaUnidadeDeNegocioEmpresaCliente()
                TiposDePagamentos()
                'BuscarMoedas()
                ddl.Carregar(ddlMoeda, CarregarDDL.Tabela.Moeda, "", False)
                BuscarIndexadores()
                ddl.Carregar(DdlProvisoes, CarregarDDL.Tabela.Provisoes, "", False)
                CarteiraDoTitulo()
                'Carteiras()
                ddl.Carregar(ddlSafraAdto, CarregarDDL.Tabela.Safra, "", True)
                Limpar(True)
                Limpar_ConsultaTitulos(True)
                txtMovimento.Text = CDate(Today).ToString("dd/MM/yyyy")
                hdnMovimentoOriginal.Value = CDate(Today).ToString("dd/MM/yyyy")
                ddl.Carregar(dltHistorico, CarregarDDL.Tabela.Historico, "", False)

                If RegistroParam > 0 Then
                    txtRegistro.Text = RegistroParam
                    'lnkConsultar_Click(lnkConsultar, New EventArgs())
                    ConsultaTitulo()
                End If

            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx", eTitulo.Info)
                Exit Sub
            End If

            TabContainer1.ActiveTabIndex = 0
        End If
    End Sub

    Public Sub MostrarCotacao()
        If Not String.IsNullOrWhiteSpace(txtProrrogacao.Text) AndAlso IsDate(txtProrrogacao.Text) AndAlso Not String.IsNullOrWhiteSpace(ddlIndexador.SelectedValue) AndAlso Not ddlIndexador.SelectedValue = 99 Then
            lblCotacao.Text = Funcoes.PegarValorConversao(ddlIndexador.SelectedValue, txtProrrogacao.Text).ToString("N4")
            lblDescCotacao.Text = ddlIndexador.SelectedItem.Text.Split("-")(1)
            divlCot.Visible = True
        End If
    End Sub

    Protected Sub ddlIndexador_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlIndexador.SelectedIndexChanged
        Try
            MostrarCotacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeDeNegocioEmpresaCliente, DdlEmpresaCliente)
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeConsultaTitulos, DdlEmpresaConsultaTitulos)
    End Sub

    Private Sub BuscarMoedas()
        Dim objMoedas As New [Lib].Negocio.Moedas()

        ddlMoeda.Items.Clear()
        ddlMoeda.Items.Add(New ListItem("", 0))

        For Each objMoeda As [Lib].Negocio.Moeda In objMoedas
            ddlMoeda.Items.Add(New ListItem(objMoeda.Codigo.ToString() & "-" & objMoeda.Descricao, objMoeda.Codigo.ToString()))
        Next
    End Sub

    Private Sub BuscarIndexadores()
        Dim objIndexadores As New [Lib].Negocio.Indexadores()

        If objIndexadores.Selecionar() Then
            For Each objIndexador As [Lib].Negocio.Indexador In objIndexadores
                ddlIndexador.Items.Add(New ListItem(objIndexador.Codigo.ToString() & "-" & objIndexador.Descricao,
                                                    objIndexador.Codigo.ToString()))
            Next
        End If

        Funcoes.InserirLinhaEmBranco(ddlIndexador)
    End Sub

    Protected Sub btnBuscaPedidoConsultaTitulos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBuscaPedidoConsultaTitulos.Click
        Try
            If String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Or String.IsNullOrWhiteSpace(txtCodigoClienteConsulta.Value) Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido.")
                Exit Sub
            End If

            Dim strEmpresa As String() = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")
            Dim strCliente As String() = txtCodigoClienteConsulta.Value.Split("-")
            HttpContext.Current.Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"
            HttpContext.Current.Session("ssCnpjDaEmpresa") = strEmpresa(0)
            HttpContext.Current.Session("ssEndDaEmpresa") = strEmpresa(1)
            HttpContext.Current.Session("txtCnpjDoCliente") = strCliente(0)
            HttpContext.Current.Session("txtEndDoCliente") = strCliente(1)

            Dim parameters As New Dictionary(Of String, Object)
            parameters("unidade") = ""
            parameters("empresa") = strEmpresa(0)
            parameters("enderecoEmpresa") = strEmpresa(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")

            Popup.ConsultaDePedidos(Me.Page, "objPedCons" & HID.Value, "txtNome")
            ucConsultaPedidos.BindGridView(parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objFornecedorCP" & HID.Value) Is Nothing Then
            objClient = CType(Session("objFornecedorCP" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objClient)
            txtFornecedor.Text = itemCliente.Text
            txtCodigoFornecedor.Value = itemCliente.Value
            txtFavorecido.Text = itemCliente.Text
            txtCodigoFavorecido.Value = itemCliente.Value
            txtPedido.Text = "0"
            txtNumeroAdto.Text = "0"
            txtTaxaAdto.Text = "0"
            txtCessaoDeCredito.Text = "0"
            Session.Remove("objFornecedorCP" & HID.Value)
        ElseIf Not Session("objFavorecidoCP" & HID.Value) Is Nothing Then
            objClient = CType(Session("objFavorecidoCP" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objClient)
            txtFavorecido.Text = itemCliente.Text
            txtCodigoFavorecido.Value = itemCliente.Value
            Session.Remove("objFavorecidoCP" & HID.Value)
        ElseIf Session("objBancoFRMTIT" & HID.Value) IsNot Nothing Then
            SessaoRecuperaTitulo()
            Dim Banco As [Lib].Negocio.ClienteXContaBancaria = CType(Session("objBancoFRMTIT" & HID.Value), [Lib].Negocio.ClienteXContaBancaria)
            ObjTitulo.CodigoBancoCliente = Banco.CodigoBanco
            ObjTitulo.CodigoAgenciaCliente = Banco.CodigoAgencia
            ObjTitulo.DigitoAgenciaCliente = Banco.DigitoAgencia
            ObjTitulo.ContaCliente = Banco.ContaCorrente
            ObjTitulo.DigitoContaCliente = Banco.DigitoConta
            ObjTitulo.TipoDaContaCliente = Banco.TipoConta
            AtualizaDadosBancariosNoForm(ObjTitulo)
            Session.Remove("objBancoFRMTIT" & HID.Value)
            SessaoSalvaTitulo()
        ElseIf Session("objPedCons" & HID.Value) IsNot Nothing Then
            Dim p As [Lib].Negocio.Pedido = CType(Session("objPedCons" & HID.Value), [Lib].Negocio.Pedido)
            txtPedidoConsultaTitulos.Text = p.Codigo
            Session.Remove("objPedCons" & HID.Value)
        ElseIf Not Session("objClienteCTAXPC" & HID.Value) Is Nothing Then
            Dim objClienteCTAXPC As [Lib].Negocio.Cliente = CType(Session("objClienteCTAXPC" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objClienteCTAXPC)
            txtClienteConsulta.Text = itemCliente.Text
            txtCodigoClienteConsulta.Value = itemCliente.Value
            Session.Remove("objClienteCTAXPC" & HID.Value)
        ElseIf Not Session("objClienteCTAXPCXED" & HID.Value) Is Nothing Then
            Dim objClienteCTAXPCXED As [Lib].Negocio.Cliente = CType(Session("objClienteCTAXPCXED" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objClienteCTAXPCXED)
            txtClienteConsultaEndosso.Text = itemCliente.Text
            txtCodigoClienteConsultaEndosso.Value = itemCliente.Value
            Session.Remove("objClienteCTAXPCXED" & HID.Value)
        ElseIf Not Session("objPedidoCTAPAG" & HID.Value) Is Nothing Then
            Cliente = DdlEmpresaCliente.SelectedValue
            campo = Cliente.Split("-")
            Dim strCliente() As String = txtCodigoFornecedor.Value.Split("-")
            objPedido = CType(Session("objPedidoCTAPAG" & HID.Value), [Lib].Negocio.Pedido)
            If Trim(txtValorEmMoeda.Text) = "" Then txtValorEmMoeda.Text = 0

            If objPedido.MomentoFinanceiro = 3 And Funcoes.VerificaPermissao(HTabela.Value, "LIBERARMOMENTOFINANCEIRO") = False And objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                MsgBox(Me.Page, "Processo Não permitido. Pedido Lançado Com Vencimentos Determinados na Emissão da Nota Fiscal.")
            ElseIf objPedido.CodigoUnidadeNegocio <> DdlUnidadeDeNegocioEmpresaCliente.SelectedValue Then
                MsgBox(Me.Page, "Unidade de Negócio da Empresa do Pedido é diferente da Unidade de Negócio da Empresa Fornecedora.")
            ElseIf objPedido.CodigoEmpresa <> campo(0) Or objPedido.EnderecoEmpresa.ToString <> campo(1) Then
                MsgBox(Me.Page, "Empresa do Pedido é diferente da Empresa Fornecedora.")
            ElseIf objPedido.CodigoCliente <> strCliente(0) Or objPedido.EnderecoCliente.ToString <> strCliente(1) Then
                MsgBox(Me.Page, "Fornecedor do Pedido é diferente do Fornecedor informado.")
            ElseIf ddlMoeda.SelectedValue <> objPedido.CodigoMoeda Then
                MsgBox(Me.Page, "A moeda do Adiantamento e do Pedido tem que ser a Mesma.")

            Else
                If chkManterLancamento.Checked AndAlso Not txtPedido.Text = objPedido.Codigo Then txtHistorico.Text = String.Empty

                txtPedido.Text = objPedido.Codigo
                'ddlMoeda.SelectedValue = objPedido.CodigoMoeda
                'ddlIndexador.SelectedValue = objPedido.CodigoIndexador
                ddlCarteiras.Enabled = False
                ddlMoeda.Enabled = False
                ddlIndexador.Enabled = False
                lnkExcluir.Parent.Visible = False
                ddlSafraAdto.Enabled = False
                ddlSafraAdto.SelectedValue = objPedido.CodigoSafra
                txtCessaoDeCredito.Parent.Visible = True
            End If
            Session.Remove("objPedidoCTAPAG" & HID.Value)
        ElseIf Not Session("objBaixaLoteFinanceiro" & HID.Value) Is Nothing Then
            CarregarBaixaLote()
        ElseIf Session("objNavioXInvoice" & HID.Value) IsNot Nothing Then
            Dim objNavioXInvoice = CType(Session("objNavioXInvoice" & HID.Value), [Lib].Negocio.NavioXInvoice)
            txtNaviosXInvoice.Text = objNavioXInvoice.Codigo & " - " & objNavioXInvoice.Descricao
        ElseIf Session("objCarregarTitulo" & HID.Value) IsNot Nothing Then
            txtRegistro.Text = Session("objCarregarTitulo" & HID.Value)
            ConsultaTitulo()
            Session.Remove("objCarregarTitulo" & HID.Value)
        End If

    End Sub

#Region "Consulta Títulos"
    Private Sub TitulosConsulta()
        Dim Cliente As String
        Dim Campo() As String
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue
        Dim Valor As Decimal = 0

        Sql = "  SELECT CP.Registro_Id AS Registro, " & vbCrLf &
              "         CASE " & vbCrLf &
              "             when LEN(ISNULL(ed.ClienteEndosso_Id,'')) > 0" & vbCrLf &
              "                then convert(varchar(10),ed.Vencimento,103)" & vbCrLf &
              "                else convert(varchar(10),CP.Prorrogacao,103)" & vbCrLf &
              "             end AS Vencimento, " & vbCrLf &
              "         CASE " & vbCrLf &
              "             when ISNULL(ed.Codigo_Id,0) > 0" & vbCrLf &
              "                then CliED.Nome" & vbCrLf &
              "                else Cli.Nome" & vbCrLf &
              "             end AS Cliente, " & vbCrLf &
              "         Historico, isnull(CP.MoedaValorDoDocumento, 0) AS Dolar, CP.ValorDoDocumento AS Valor, " & vbCrLf &
              "         ISNULL(CP.MoedaValorLiquido, 0) AS MoedaLiquido, CP.ValorLiquido AS ValorLiquido, " & vbCrLf &
              "         UsuarioLiberacao as Liberado, CP.Pedido as Pedido, " & vbCrLf &
              "         CASE " & vbCrLf &
              "             WHEN CP.Moeda = 0 THEN 'R$-' + convert(varchar,CP.Moeda) " & vbCrLf &
              "             ELSE " & vbCrLf &
              "                 CASE " & vbCrLf &
              "                     WHEN CP.Moeda = 1  THEN 'R$-' + convert(varchar,CP.Moeda) " & vbCrLf &
              "                     ELSE 'U$-' + convert(varchar,CP.Moeda) " & vbCrLf &
              "                 END " & vbCrLf &
              "         END as Moeda, CP.Indexador, isnull(CP.Grupado,'N') as Grupado, CP.Provisao, CP.Situacao " & vbCrLf &
              "    FROM " & HTabela.Value & " CP " & vbCrLf &
              "    LEFT JOIN NotaFiscalXTitulo NFXT" & vbCrLf &
              "      ON CP.Empresa     = NFXT.Empresa_Id" & vbCrLf &
              "     AND CP.EndEmpresa  = NFXT.EndEmpresa_Id" & vbCrLf &
              "     AND CP.Registro_Id = NFXT.Titulo_Id" & vbCrLf &
              "    LEFT JOIN FaturaDeFreteXTitulo FFxT" & vbCrLf &
              "      ON CP.Registro_Id = FFxT.Titulo_Id" & vbCrLf &
              "    LEFT JOIN FaturasDeFretesXItens FFxI" & vbCrLf &
              "      ON FFxT.Empresa_Id       = FFxI.EmpresaPagadora_Id " & vbCrLf &
              "     AND FFxT.EndEmpresa_Id    = FFxI.EndEmpresaPagadora_Id" & vbCrLf &
              "     AND FFxT.Conveniado_Id    = FFxI.Conveniado_Id" & vbCrLf &
              "     AND FFxT.EndConveniado_Id = FFxI.EndConveniado_Id" & vbCrLf &
              "     AND FFxT.Fatura_Id        = FFxI.Fatura_Id" & vbCrLf &
              "   INNER JOIN Clientes Cli" & vbCrLf &
              "      ON CP.Cliente = Cli.Cliente_Id " & vbCrLf &
              "     AND CP.EndCliente = Cli.Endereco_Id" & vbCrLf &
              "  LEFT JOIN EndossoXTitulo eTxT" & vbCrLf &
              "    ON eTxT.Empresa_Id    = CP.Empresa" & vbCrLf &
              "   AND eTxT.EndEmpresa_Id = CP.EndEmpresa" & vbCrLf &
              "   AND eTxT.Titulo_Id     = CP.Registro_Id" & vbCrLf &
              "  LEFT JOIN Endosso ed" & vbCrLf &
              "    ON ed.Empresa_Id    = eTxT.Empresa_Id" & vbCrLf &
              "   AND ed.EndEmpresa_Id = eTxT.EndEmpresa_Id" & vbCrLf &
              "   AND ed.Codigo_Id     = eTxT.Codigo_Id" & vbCrLf &
              "  LEFT JOIN Clientes CliED" & vbCrLf &
              "    ON CliED.Cliente_Id  = ed.ClienteEndosso_Id" & vbCrLf &
              "   AND CliED.Endereco_Id = ed.EndClienteEndosso_Id" & vbCrLf &
              "   WHERE 1=1 " & vbCrLf

        If RbCancelado.Checked Then
            Sql &= " AND CP.Situacao <> 1 " & vbCrLf
        ElseIf RbAtivo.Checked Then
            If Funcoes.VerificaPermissao("TituloDeFuncionario", "LEITURA") Then
                Sql &= " AND CP.Situacao in(1,101,102,105) " & vbCrLf
            Else
                Sql &= " AND CP.Situacao in(1,101,102) " & vbCrLf
            End If
        End If

        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            Sql &= " AND CP.UnidadeDeNegocio = '" & Cliente & "' " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtPedidoConsultaTitulos.Text) Then
            Sql &= " AND CP.Pedido = '" & txtPedidoConsultaTitulos.Text & "'" & vbCrLf
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")

        If Campo(0) <> "" Then
            Sql &= " AND CP.Empresa = '" & Campo(0) & "'" & vbCrLf  'Empresa
            Sql &= " AND CP.EndEmpresa = " & Campo(1) & vbCrLf      'Endereco da Empresa
        End If

        'Campo = txtCodigoClienteConsulta.Value.Split("-")
        'If Campo(0) <> "" Then
        '    Sql &= " AND CP.Cliente = '" & Campo(0) & "'" & vbCrLf  'Cliente
        '    Sql &= " AND CP.EndCliente = " & Campo(1) & vbCrLf    'Cliente da Empresa
        'End If

        If Not String.IsNullOrWhiteSpace(txtCodigoClienteConsultaEndosso.Value) Then
            Sql &= "   AND ed.ClienteEndosso_Id    = '" & txtCodigoClienteConsultaEndosso.Value.Split("-")(0) & "'" & vbCrLf &
                   "   AND ed.EndClienteEndosso_Id = " & txtCodigoClienteConsultaEndosso.Value.Split("-")(1) & vbCrLf
        ElseIf Not String.IsNullOrWhiteSpace(txtCodigoClienteConsulta.Value) Then
            Campo = txtCodigoClienteConsulta.Value.Split("-")
            Sql &= " AND CP.Cliente    = '" & Campo(0) & "'" & vbCrLf  'Cliente
            Sql &= " AND CP.EndCliente = " & Campo(1) & vbCrLf    'Cliente da Empresa
        End If

        If Not String.IsNullOrWhiteSpace(txtNumNota.Text) Then
            Sql &= " AND (ISNULL(NFXT.Nota_Id,0) in(" & txtNumNota.Text & ") OR ISNULL(FFxI.Fatura_Id,0) in(" & txtNumNota.Text & ") OR ISNULL(FFxI.Nota_Id,0) in(" & txtNumNota.Text & "))" & vbCrLf
        ElseIf Not String.IsNullOrWhiteSpace(txtCodigoClienteConsultaEndosso.Value) Then
            Sql &= " AND ed.Vencimento between '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf
        ElseIf chkClienteEndosso.Checked AndAlso String.IsNullOrWhiteSpace(txtCodigoClienteConsultaEndosso.Value) Then
            Sql &= "   AND ((ed.Vencimento BETWEEN '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "') OR " & vbCrLf
            Sql &= "        (Prorrogacao BETWEEN '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'))" & vbCrLf
        ElseIf txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
            Sql &= " AND Prorrogacao between '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf
        End If

        If ChkAutozizado.Checked = True Then
            Sql &= " AND CP.UsuarioLiberacao <> ''" & vbCrLf  'Autorizados
        End If

        If RbAtivo.Checked Then
            If chkPrevisao.Visible AndAlso chkProvisao.Visible AndAlso chkBaixado.Visible Then
                If chkPrevisao.Checked AndAlso chkProvisao.Checked AndAlso chkBaixado.Checked Then
                    Sql &= " AND (Provisao = 2 OR Provisao = 3 OR Provisao = 1) " & vbCrLf
                ElseIf chkPrevisao.Checked AndAlso chkProvisao.Checked Then
                    Sql &= " AND (Provisao = 2 OR Provisao = 3 ) " & vbCrLf
                ElseIf chkPrevisao.Checked AndAlso chkBaixado.Checked Then
                    Sql &= " AND (Provisao = 2 OR Provisao = 1 ) " & vbCrLf
                ElseIf chkProvisao.Checked AndAlso chkBaixado.Checked Then
                    Sql &= " AND (Provisao = 3 OR Provisao = 1 ) " & vbCrLf
                ElseIf chkPrevisao.Checked = True Then
                    Sql &= " AND Provisao = 2 " & vbCrLf
                ElseIf chkProvisao.Checked = True Then
                    Sql &= " AND Provisao = 3 " & vbCrLf
                ElseIf chkBaixado.Checked = True Then
                    Sql &= " AND Provisao = 1 " & vbCrLf
                End If
            End If
            lnkReprogramar.Enabled = True
        End If
        Sql &= " ORDER BY CP.Prorrogacao, Cli.Nome"

        DS = Banco.ConsultaDataSet(Sql, "Contas")

        If DS Is Nothing OrElse DS.Tables(0).Rows.Count = 0 Then
            GridConsultaTitulos.DataSource = Nothing
            GridConsultaTitulos.DataBind()
            MsgBox(Me.Page, "Nenhum Registro encontrado...")
        Else
            GridConsultaTitulos.DataSource = DS
            GridConsultaTitulos.DataBind()

            Dim i As Integer = 0
            While i < GridConsultaTitulos.Rows.Count
                Dim strMoeda() As String = GridConsultaTitulos.Rows(i).Cells(10).Text.ToString.Split("-")
                If strMoeda(0) = "U$" Then
                    GridConsultaTitulos.Rows(i).ForeColor = System.Drawing.Color.Red
                End If


                If Not GridConsultaTitulos.Rows(i).Cells(12).Text.ToString.Contains("S") Then
                    If (GridConsultaTitulos.Rows(i).Cells(10).Text.Equals("R$-1")) Then
                        Valor = Valor + CDec(GridConsultaTitulos.Rows(i).Cells(8).Text)
                    Else
                        Valor = Valor + CDec(GridConsultaTitulos.Rows(i).Cells(7).Text)
                    End If
                End If

                If GridConsultaTitulos.Rows(i).Cells(9).Text = "&nbsp;" Or GridConsultaTitulos.Rows(i).Cells(9).Text = "" Then
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Enabled = True
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkLiberado"), CheckBox).Enabled = False
                Else
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkLiberado"), CheckBox).Checked = True
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkLiberado"), CheckBox).Enabled = False
                End If

                If GridConsultaTitulos.Rows(i).Cells(12).Text = "S" Or GridConsultaTitulos.Rows(i).Cells(12).Text = "M" Then 'GRUPADO BLOQUEIA
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Enabled = False
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkLiberado"), CheckBox).Enabled = False
                End If

                If RbAtivo.Checked Then
                    CType(GridConsultaTitulos.Rows(i).FindControl("chkReprogramar"), CheckBox).Enabled = True
                Else
                    CType(GridConsultaTitulos.Rows(i).FindControl("chkReprogramar"), CheckBox).Enabled = False
                End If

                If DS.Tables(0).Rows(i).Item("Provisao") = "1" Then
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Enabled = False
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkLiberado"), CheckBox).Enabled = False
                    CType(GridConsultaTitulos.Rows(i).FindControl("chkReprogramar"), CheckBox).Enabled = False
                    GridConsultaTitulos.Rows(i).Cells(3).ForeColor = System.Drawing.Color.Red
                    GridConsultaTitulos.Rows(i).Cells(3).ToolTip = "BAIXADO"
                End If

                If RbCancelado.Checked OrElse GridConsultaTitulos.Rows(i).Cells(15).Text = "101" OrElse GridConsultaTitulos.Rows(i).Cells(15).Text = "102" Then

                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Enabled = False
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkLiberado"), CheckBox).Enabled = False
                    CType(GridConsultaTitulos.Rows(i).FindControl("chkReprogramar"), CheckBox).Enabled = False
                    CType(GridConsultaTitulos.Rows(i).FindControl("chkRecibo"), CheckBox).Enabled = False

                    If Funcoes.VerificaPermissao(HTabela.Value, "LIBERARMOMENTOFINANCEIRO") Then
                        CType(GridConsultaTitulos.Rows(i).FindControl("chkReprogramar"), CheckBox).Enabled = True
                    End If

                End If

                GridConsultaTitulos.Rows(i).Cells(15).Visible = False

                i += 1
            End While

            If Valor > 0 Then
                lblTotalRegistroAgrupado.Parent.Visible = True
                lblTotalRegistroAgrupado.Text = " Total de Título(s) no valor de: " & String.Format("{0:N2}", Valor)
            End If

            lnkSlip.Parent.Visible = True

            If Not RbCancelado.Checked Then
                lnkRecibo.Parent.Visible = True
                lnkReprogramar.Parent.Visible = True
            End If

            rowDolar.Visible = True

            If RbAtivo.Checked AndAlso chkPrevisao.Checked Then

                Dim objEmpresa As ClienteXEmpresa = New ClienteXEmpresa(Session("ssEmpresa"), Session("ssEndEmpresa"))

                If objEmpresa.BaixaFinanceiroPorLote Then
                    If chkProvisao.Checked OrElse chkBaixado.Checked Then
                        lnkBaixarSelecionados.Parent.Visible = False
                    Else
                        lnkBaixarSelecionados.Parent.Visible = True
                    End If
                Else
                    lnkBaixarSelecionados.Parent.Visible = False
                End If
            Else
                lnkBaixarSelecionados.Parent.Visible = False
            End If
        End If
    End Sub
#End Region

#Region "Manutenção dos Títulos"

    Private Sub CargaUnidadeDeNegocioEmpresaCliente()
        ddl.Carregar(DdlUnidadeDeNegocioEmpresaCliente, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        ddl.Carregar(DdlUnidadeConsultaTitulos, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        ddl.Carregar(DdlEmpresaPagadora, CarregarDDL.Tabela.ClientesXEmpresas, " left(CE.Empresa_Id,8) = '" & Left(Session("ssEmpresa").ToString, 8) & "'", True)
    End Sub

#End Region

    Protected Sub DdlUnidadeDeNegocioEmpresaCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(DdlEmpresaCliente, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocioEmpresaCliente.SelectedValue, True)
    End Sub

    Private Sub CargaEmpresaCliente()
        ddl.Carregar(DdlEmpresaCliente, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocioEmpresaCliente.SelectedValue, True)
    End Sub

    Private Sub TiposDePagamentos()

        Dim Parametros As New Hashtable
        Parametros.Clear()
        Parametros.Add("listarTudo", "N")

        ddl.Carregar(DdlTiposDePagamentos, CarregarDDL.Tabela.TipoDePagamento, "", True, Parametros)

        'Sql = "SELECT TipoDePagamento_Id as Codigo, convert(varchar,REPLICATE('0', 2 - LEN(CAST(TipoDePagamento_Id AS varchar))) + CAST(TipoDePagamento_Id AS varchar)) + '  -  ' + Descricao as Descricao FROM TiposDePagamentos Order By TipoDePagamento_Id"

        'DdlTiposDePagamentos.DataValueField = "Codigo"
        'DdlTiposDePagamentos.DataTextField = "Descricao"
        'DdlTiposDePagamentos.DataSource = Banco.ConsultaDataSet(Sql, "TiposDePagamentos")
        'DdlTiposDePagamentos.DataBind()

        'DdlTiposDePagamentos.Items.Insert(0, "")
        'DdlTiposDePagamentos.SelectedIndex = 0
    End Sub

    Private Sub CarteiraDoTitulo()
        Dim objCarteiraDoTitulo As New [Lib].Negocio.ListCarteiraDoTitulo()

        ddlCarteiraDoTitulo.DataValueField = "Codigo"
        ddlCarteiraDoTitulo.DataTextField = "Descricao"
        ddlCarteiraDoTitulo.DataSource = objCarteiraDoTitulo.ToArray()
        ddlCarteiraDoTitulo.DataBind()
        ddlCarteiras.SelectedIndex = 0
    End Sub

    Private Sub Carteiras(Optional ppedido As Integer = 0, Optional Carteira As String = "")
        ddlCarteiras.Items.Clear()
        Sql = "SELECT Cart.Produto_Id AS Codigo," & vbCrLf &
              "       case" & vbCrLf &
              "			when isnull(Cart.ContaClientes,'') = ''" & vbCrLf &
              "			   then Cart.Produto_Id + '  -  ' + Cart.Descricao" & vbCrLf &
              "			   else Cart.Produto_Id + '  -  ' + Cart.Descricao + ' (' + Cart.ContaClientes + '-' + pl.Titulo + ')'" & vbCrLf &
              "		  end AS Descricao" & vbCrLf &
              "" & vbCrLf &
              "  FROM ComprasXProdutos Cart" & vbCrLf &
              "  LEFT JOIN PlanoDeContas pl " & vbCrLf &
              "         on pl.Conta_Id = Cart.ContaClientes " & vbCrLf

        If ppedido > 0 Then
            Dim emp() As String = DdlEmpresaCliente.SelectedValue.Split("-")
            Sql &= " inner Join Pedidos P" & vbCrLf &
                   "    on P.Empresa_Id    ='" & emp(0) & "'" & vbCrLf &
                   "   and P.EndEmpresa_Id = " & emp(1) & vbCrLf &
                   "   and P.Pedido_Id     = " & ppedido & vbCrLf &
                   " inner join SubOperacoes so" & vbCrLf &
                   "    on p.Operacao    = so.Operacao_Id" & vbCrLf &
                   "   and p.SubOperacao = so.SubOperacoes_Id" & vbCrLf &
                   "   and (cart.ContaClientes = case" & vbCrLf &
                   "                              when len(isnull(so.GrupoDeContas,0)) = 0" & vbCrLf &
                   "							   then Cart.ContaClientes" & vbCrLf &
                   "							   else so.GrupoDeContas" & vbCrLf &
                   "							end" & vbCrLf
            If Not String.IsNullOrEmpty(Carteira) Then
                Sql &= "    OR Cart.Produto_Id = '" & Carteira & "'"
            End If
            Sql &= ")" & vbCrLf
        End If

        Sql &= " Where Classificacao = '" & IIf(HTabela.Value = "ContasAReceber", "R", "P") & "'" & vbCrLf &
                " Order By Produto_Id" & vbCrLf

        ddlCarteiras.DataValueField = "Codigo"
        ddlCarteiras.DataTextField = "Descricao"
        ddlCarteiras.DataSource = Banco.ConsultaDataSet(Sql, "Carteiras")
        ddlCarteiras.DataBind()

        ddlCarteiras.Items.Insert(0, "")
        ddlCarteiras.SelectedIndex = 0
    End Sub

    Private Sub CargaBancoPagador()
        DdlBancoPagador.Items.Clear()
        DdlContaPagadora.Items.Clear()

        If DdlEmpresaPagadora.Text <> "" Then

            Cliente = DdlEmpresaPagadora.SelectedValue
            campo = Cliente.Split("-")

            'Sql = " SELECT DISTINCT BxC.Banco_Id,  B.Descricao, BxC.ContaContabil, p.Titulo" & vbCrLf &
            Sql = " SELECT DISTINCT BxC.Banco_Id,  B.Descricao" & vbCrLf &
                  "   FROM BancosXContas BxC " & vbCrLf &
                  "  INNER JOIN Bancos B  " & vbCrLf &
                  "          ON BxC.Banco_Id = B.Banco_Id" & vbCrLf &
                  "  LEFT JOIN PlanoDeContas p" & vbCrLf &
                  "         on p.Conta_Id = BxC.ContaContabil" & vbCrLf &
                  "  WHERE BxC.Empresa_Id     ='" & campo(0) & "'" & vbCrLf &
                  "    AND BxC.EndEmpresa_Id  = " & campo(1)

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Bancos").Tables(0).Rows
                'Descricao = Format(Dr("Banco_Id"), "0000") & " - " & Dr("Descricao") & " - (" & Dr("ContaContabil") & "-" & Dr("Titulo") & ")"
                Descricao = Format(Dr("Banco_Id"), "0000") & " - " & Dr("Descricao")
                DdlBancoPagador.Items.Add(New ListItem(Descricao, Dr("Banco_Id")))
            Next

            DdlBancoPagador.Items.Insert(0, "")
            DdlBancoPagador.SelectedIndex = 0
        End If
    End Sub

    Private Function ValidaCampos() As Boolean
        SessaoRecuperaTitulo()
        Cliente = DdlEmpresaCliente.SelectedValue
        campo = Cliente.Split("-")
        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)

        'EMPRESA PAGAGORA DEVE SER DO MESMO GRUPO DA EMPRESA DO CLIENTE/FORNECEDOR - FURLAN - 13/05/2025
        If DdlProvisoes.SelectedValue = 1 AndAlso Not Left(DdlEmpresaPagadora.SelectedValue.Split("-")(0), 8) = Left(DdlEmpresaCliente.SelectedValue.Split("-")(0), 8) Then
            If HTabela.Value = "ContasAPagar" Then
                MsgBox(Me.Page, "Empresa Pagadora deve ser do mesmo grupo da Empresa do Fornecedor.", eTitulo.Info)
            Else
                MsgBox(Me.Page, "Empresa Recebedora deve ser do mesmo grupo da Empresa do Cliente.", eTitulo.Info)
            End If

            Return False
        End If

        If DdlEmpresaPagadora.SelectedIndex <= 0 AndAlso DdlProvisoes.SelectedValue = 1 Then
            If HTabela.Value = "ContasAPagar" Then
                MsgBox(Me.Page, "Empresa Pagadora é Obrigatória.")
            Else
                MsgBox(Me.Page, "Empresa Recebedora é Obrigatória.")
            End If
            Return False
        ElseIf DdlTiposDePagamentos.SelectedIndex <= 0 AndAlso DdlProvisoes.SelectedValue = 1 Then
            MsgBox(Me.Page, "O Tipo de Pagamento é Obrigatório.")
            Return False
        ElseIf DdlBancoPagador.SelectedIndex <= 0 AndAlso DdlProvisoes.SelectedValue = 1 Then
            If HTabela.Value = "ContasAPagar" Then
                MsgBox(Me.Page, "O Banco Pagador é Obrigatório.")
            Else
                MsgBox(Me.Page, "O Banco Recebedor é Obrigatório.")
            End If
            Return False
        ElseIf DdlContaPagadora.SelectedIndex <= 0 AndAlso DdlProvisoes.SelectedValue = 1 Then
            If HTabela.Value = "ContasAPagar" Then
                MsgBox(Me.Page, "A Conta Pagadora é Obrigatória.")
            Else
                MsgBox(Me.Page, "A Conta Recebedora é Obrigatória.")
            End If
            Return False
        End If

        'Empresa Pagadora para utilizar na validação de datas não programáveis
        Dim Emp() As String = DdlEmpresaPagadora.SelectedValue.Split("-")

        If objCarteira.isAdiantamento AndAlso ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento Then
            MsgBox(Me.Page, "A finalidade financeira e a conta do banco pagador nao podem ser de adiantamento simultaneamente, verifique o lançamento.")
            Return False
        End If

        If txtRegistro.Text = "" Then
            If String.IsNullOrWhiteSpace(txtProrrogacao.Text) Then
                txtProrrogacao.Text = IIf(DdlProvisoes.SelectedValue = 1, txtMovimento.Text, txtDataEntradaSistema.Text)
            End If
        End If

        If String.IsNullOrWhiteSpace(txtMovimento.Text) Then
            MsgBox(Me.Page, "Data de Movimento é obrigatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtProrrogacao.Text) Then
            MsgBox(Me.Page, "Data de Vencimento é obrigatório.")
            Return False
        ElseIf DdlProvisoes.SelectedValue = 1 And Not IsDate(txtMovimento.Text) Then
            MsgBox(Me.Page, "Data da Baixa é obrigatório.")
            Return False
        ElseIf DdlProvisoes.SelectedValue = 2 AndAlso CDate(txtMovimento.Text) > CDate(txtProrrogacao.Text) Then
            MsgBox(Me.Page, "Data de Vencimento não pode ser maior que a data de Entrada no Sistema do Titulo.")
            Return False
        ElseIf DdlProvisoes.SelectedValue = 1 AndAlso Not ValidaData(txtProrrogacao.Text, "Vencimento", Emp(0), Emp(1)) Then
            MsgBox(Me.Page, Mensagem)
            Return False
        ElseIf DdlUnidadeDeNegocioEmpresaCliente.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            Return False
        ElseIf DdlEmpresaCliente.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa do Fornecedor é obrigatório.")
            Return False
        ElseIf txtCodigoFornecedor.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Fornecedor é obrigatório.")
            Return False
        ElseIf DdlProvisoes.SelectedValue = 3 AndAlso (String.IsNullOrWhiteSpace(txtPedido.Text) OrElse CInt(txtPedido.Text = 0)) Then
            MsgBox(Me.Page, "PROVISÃO não pode ser usado sem Pedido.")
            Return False
        End If

        If ObjTitulo.TituloOriginal IsNot Nothing AndAlso ObjTitulo.TituloOriginal.CodigoProvisao = eProvisao.Baixa Then
            If Not ValidaData(ObjTitulo.TituloOriginal.Baixa.ToString("dd-MM-yyyy"), "Movimento", Emp(0), Emp(1)) Then
                MsgBox(Me.Page, "No Titulo Original, " & Mensagem)
                Return False
            End If
        End If

        If DdlProvisoes.SelectedValue = 1 Then
            If Not ValidaData(txtMovimento.Text, "Movimento", Emp(0), Emp(1)) Then
                MsgBox(Me.Page, Mensagem)
                Return False
            End If
        End If

        If HTabela.Value = "ContasAPagar" AndAlso DdlProvisoes.SelectedValue = 1 AndAlso ChkLiberado.Checked = False AndAlso Not Funcoes.VerificaPermissao("AutorizacoesDePagamentos", "GRAVAR") Then
            If ddlCarteiras.SelectedValue = "001001057" OrElse
               ddlCarteiras.SelectedValue = "001001069" OrElse
               ddlCarteiras.SelectedValue = "001001067" OrElse
               ddlCarteiras.SelectedValue = "001001005" OrElse
               ddlCarteiras.SelectedValue = "001001052" OrElse
               ddlCarteiras.SelectedValue = "001001074" Then
            ElseIf CDbl(txtValorCobrado.Text) < 151 Then
            Else
                MsgBox(Me.Page, "Titulo Não Autorizado para Pagamento")
                Return False
            End If
        End If

        If DdlProvisoes.SelectedValue = 1 Then
            If DdlEmpresaPagadora.Text <> "" AndAlso DdlTiposDePagamentos.Text = "" Then
                MsgBox(Me.Page, "Tipo de Pagamento é obrigatório")
                Return False
            ElseIf DdlEmpresaPagadora.Text = "" Then
                If HTabela.Value = "ContasAPagar" Then
                    MsgBox(Me.Page, "Empresa Pagadora é obrigatório")
                Else
                    MsgBox(Me.Page, "Empresa Recebedora é obrigatório")
                End If
                Return False
            ElseIf DdlBancoPagador.Text = "" Then
                MsgBox(Me.Page, "Banco é obrigatório")
                Return False
            ElseIf DdlContaPagadora.Text = "" Then
                MsgBox(Me.Page, "Conta Bancária é obrigatório")
                Return False
            ElseIf DdlTiposDePagamentos.Text = "" Then
                MsgBox(Me.Page, "Tipo De Pagamento é obrigatório para Baixa")
                Return False
            ElseIf txtValorDoDocumento.Text = 0 And ddlCarteiras.SelectedValue <> "999999999" And ddlCarteiras.SelectedValue <> "999999998" Then
                MsgBox(Me.Page, "Valor Em Reais é Obrigatório")
                Return False
            ElseIf txtValorCobrado.Text = "" And lblAgrupar.Text <> "BT" Then
                If HTabela.Value = "ContasAPagar" Then
                    MsgBox(Me.Page, "Valor Pago é obrigatório")
                Else
                    MsgBox(Me.Page, "Valor Recebido é obrigatório")
                End If
                Return False
            ElseIf Not lblAgrupar.Text = "AP" AndAlso ddlCarteiras.SelectedIndex > 0 AndAlso DdlTributos.Items.Count > 1 AndAlso DdlTributos.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Encargo é obrigatório.")
                Return False
            End If

            Dim SelecaoBancoxConta As Array = DdlContaPagadora.SelectedValue.Split("-") 'Insol - Somente baixar pela filial conta <> caixa atravez de autorização.
            If Not SelecaoBancoxConta(5) = "101010101" And Not Funcoes.VerificaPermissao("BAIXAMATRIZ", "GRAVAR") Then
                MsgBox(Me.Page, "Usuário Sem Autorização para Baixar Registro, Somente Baixas Via Matriz São Autorizadas.")
                Return False
            End If

            If ObjTitulo.ObjContaContabilPagadora.Adiantamento AndAlso lblAgrupar.Text = "AP" Then
                MsgBox(Me.Page, "Agrupamento não pode ser feito com conta de Adiantamrnto ou Baixa de Adiantamento.")
                Return False
            End If

            'Documento NFG necessita de conferencia fiscal realizada para liberação
            'If Not SelecaoBancoxConta(4) = "101010101" AndAlso ViewState("CONTROLE_FISCAL") IsNot Nothing AndAlso ViewState("CONTROLE_FISCAL").Equals("BLOQUEADO") Then
            '    MsgBox(Me.Page, "Registro aguardando conferência fiscal")
            '    Return False
            'End If
        End If

        If ddlIndexador.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Indexador é obrigatório.")
            Return False
            'ElseIf DdlProvisoes.SelectedValue = 3 And Not Funcoes.VerificaPermissao(HTabela.Value, "LIBERAR") Then
            '    MsgBox(Me.Page, "Usuário sem autorização para lançamento de Provisão")
            '    Return False
        ElseIf txtHistorico.Text = "" And lblAgrupar.Text <> "BT" Then
            MsgBox(Me.Page, "Histórico é obrigatório.")
            Return False
        ElseIf ddlCarteiras.Text = "" AndAlso Not lblAgrupar.Text = "AP" Then
            MsgBox(Me.Page, "Finalidade Financeira é obrigatório.")
            Return False
        End If

        If Not Funcoes.VerificaPermissao("AJUSTEFINANCEIRO", "GRAVAR") Then
            If (String.IsNullOrWhiteSpace(txtValorDoDocumento.Text) AndAlso String.IsNullOrWhiteSpace(txtValorEmMoeda.Text)) Then
                MsgBox(Me.Page, "Valor Do Documento em R$ ou U$ é obrigatório e maior que zero!")
                Return False
            ElseIf (ddlMoeda.SelectedValue = 1 AndAlso (String.IsNullOrWhiteSpace(txtValorDoDocumento.Text) OrElse CDbl(txtValorDoDocumento.Text) <= 0)) And ddlCarteiras.SelectedValue <> "999999999" And ddlCarteiras.SelectedValue <> "999999998" Then
                MsgBox(Me.Page, "Valor Do Documento em R$ é obrigatório.")
                Return False
            ElseIf (ddlMoeda.SelectedValue = 3 AndAlso (String.IsNullOrWhiteSpace(txtValorEmMoeda.Text) OrElse CDbl(txtValorEmMoeda.Text) <= 0)) And ddlCarteiras.SelectedValue <> "999999999" And ddlCarteiras.SelectedValue <> "999999998" Then
                MsgBox(Me.Page, "Valor Do Documento em U$ é obrigatório.")
                Return False
            ElseIf (CDec(txtValorCobrado.Text) = 0 And (CDec(txtDescontos.Text) + CDec(txtDeducoes.Text)) <> CDec(txtValorDoDocumento.Text)) And ddlCarteiras.SelectedValue <> "999999999" And ddlCarteiras.SelectedValue <> "999999998" Then
                MsgBox(Me.Page, "Valor Do Documento em U$ é obrigatório.")
                Return False
            End If
        End If

        If HTabela.Value = "ContasAPagar" Then
            If DdlTiposDePagamentos.Text <> "" Then
                If DdlTiposDePagamentos.SelectedValue = 3 And ObjTitulo.CodigoBancoCliente = 0 Then
                    MsgBox(Me.Page, "Dados Bancários é obrigatório.")
                    Return False
                End If
            End If
        End If

        Dim Empresa As New [Lib].Negocio.ClienteXEmpresa(campo(0), campo(1))

        If Not String.IsNullOrWhiteSpace(Empresa.CodigoContaFornecedorFrete) _
            AndAlso Not lblAgrupar.Text = "AP" AndAlso String.IsNullOrWhiteSpace(txtRegistro.Text) _
            AndAlso Empresa.CodigoContaFornecedorFrete = objCarteira.CodigoContaCliente Then
            MsgBox(Me.Page, "Titulo não pode ser incluído na Conta de Fornecedor de Frete.")
            Return False
        End If

        If Not Funcoes.VerificaPermissao("AJUSTEFINANCEIRO", "GRAVAR") Then
            If txtLiberarTitulo.Value = "N" AndAlso txtPedido.Text <> "0" AndAlso lblAgrupar.Text <> "AP" Then
                If ddlMoeda.SelectedValue = 1 AndAlso Not objCarteira.isAdiantamento Then
                    If HDValorOriginalOficial.Value <> 0 Then
                        If CDbl(txtValorDoDocumento.Text) > HDValorOriginalOficial.Value + 1 Then
                            MsgBox(Me.Page, "Valor do documento não pode ser maior do que o valor programado pelo pedido.")
                            Return False
                        End If
                    End If
                End If

                If ddlMoeda.SelectedValue = 3 AndAlso lblAgrupar.Text <> "AP" AndAlso Not objCarteira.isAdiantamento Then
                    If HDValorOriginalMoeda.Value <> 0 Then
                        If CDbl(txtValorEmMoeda.Text) > HDValorOriginalMoeda.Value + 1 Then
                            MsgBox(Me.Page, "Valor do documento não pode ser maior do que o valor programado pelo pedido.")
                            Return False
                        End If
                    End If
                End If
            End If
        End If

        If txtNumeroCheque.Text.Length > 0 AndAlso txtNumeroCheque.Text > 0 Then
            MsgBox(Me.Page, "Título com cheque emitido não pode ser alterado.")
            Return False
        End If

        If DdlTributos.SelectedIndex > 0 Then 'Consulta Contas Contabeis dos Tributos
            Dim Encargo As New Encargo(DdlTributos.SelectedValue)
            If String.IsNullOrEmpty(Encargo.ContaDebito) And String.IsNullOrEmpty(Encargo.ContaCredito) Then
                MsgBox(Me.Page, "Encargo Sem contas Credito e Debito Cadastradas, Verifique o encargo.")
                Return False
            ElseIf Not String.IsNullOrWhiteSpace(Empresa.CodigoContaFornecedorFrete) _
                AndAlso String.IsNullOrWhiteSpace(txtRegistro.Text) AndAlso (Empresa.CodigoContaFornecedorFrete = Encargo.ContaDebito OrElse Empresa.CodigoContaFornecedorFrete = Encargo.ContaCredito) Then
                MsgBox(Me.Page, "Titulo não pode ser incluído na Conta de Fornecedor de Frete.")
                Return False
            End If
        End If


        '*****************************************************************************************************************
        '************************** Primeira Carteira Adiantamento/Baixa *************************************************
        '*****************************************************************************************************************

        'Sem finalidade, quando é Adiantamento e não é Baixa sempre vai gerar um novo - Furlan - 29/08/2016
        'If objCarteira.isAdiantamento AndAlso Not objCarteira.BaixaAdiantamento AndAlso (Not IsNumeric(txtNumeroAdto.Text) OrElse CInt(txtNumeroAdto.Text) <= 0) Then
        '    MsgBox(Me.Page, "Numero de adiantamento não Carregado.")
        '    Return False
        'End If

        If Not objCarteira.isAdiantamento AndAlso ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso Not ObjTitulo.ObjContaContabilPagadora.Adiantamento AndAlso IsNumeric(txtNumeroAdto.Text) AndAlso CInt(txtNumeroAdto.Text) > 0 Then
            MsgBox(Me.Page, "Numero de adiantamento nao pode ser informado com carteira que nao sao de adiantamento ou baixa. reinicie o lançamento.")
            Return False
        End If

        If objCarteira.isAdiantamento AndAlso Not objCarteira.BaixaAdiantamento AndAlso String.IsNullOrWhiteSpace(txtVencimentoAdto.Text) Then
            MsgBox(Me.Page, "Vencimento para o Adiantamento não foi informado, Verifique.")
            Return False
        End If

        If DdlProvisoes.SelectedValue = 1 AndAlso objCarteira.BaixaAdiantamento Then
            If ddlMoeda.SelectedValue = 1 Then
                If Not ObjTitulo.Adiantamentos.Select(Function(s) s.VlrABaixarOficial).Sum() = CDec(txtValorDoDocumento.Text) Then
                    MsgBox(Me.Page, "Valor da Baixa não pode ser maior do que o saldo do adiantamento, em " & ddlMoeda.SelectedItem.Text & ".")
                    Return False
                End If
            Else
                If Not ObjTitulo.Adiantamentos.Select(Function(s) s.VlrABaixarMoeda).Sum() = CDec(txtValorEmMoeda.Text) Then
                    MsgBox(Me.Page, "Valor da Baixa não pode ser maior do que o saldo do adiantamento, em " & ddlMoeda.SelectedItem.Text & ".")
                    Return False
                End If
            End If
        End If

        '*****************************************************************************************************************
        '************************** Segunda Carteira Adiantamento/Baixa *************************************************
        '*****************************************************************************************************************
        Dim ValorAdtoOficial As Decimal
        Dim ValorAdtoMoeda As Decimal

        If ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento Then
            ValorAdtoOficial = txtValorCobrado.Text
            ValorAdtoMoeda = txtValorCobradoMoeda.Text
        End If

        If objCarteira.BaixaAdiantamento Then
            ValorAdtoOficial = txtValorDoDocumento.Text
            ValorAdtoMoeda = txtValorEmMoeda.Text
        End If


        If DdlProvisoes.SelectedValue = 1 AndAlso (ValorAdtoOficial > 0 OrElse ValorAdtoMoeda > 0) Then

            If Left(Session("ssEmpresa").ToString, 8) = "24450490" AndAlso DdlProvisoes.SelectedValue = 1 AndAlso CDate(txtMovimento.Text) < CDate("01/07/2024") Then
                'LIBERADO PARA BAIXAR ADIANTAMENTO DE SALDO INICIAL RT GRÃOS - FURLAN 18/07/2024
            ElseIf Left(ObjTitulo.ContaContabilPagadora, 1) = IIf(HTabela.Value = "ContasAPagar", "2", "1") Then
                'Continua sem verificar Lista de Adiantamentos, pois está dando um adiantamento, não baixando - Furlan - 09/05/2020
            Else
                Dim moeda As New Moeda(ddlMoeda.SelectedValue)
                If moeda.Classificacao = eTiposMoeda.Oficial Then
                    If ValorAdtoOficial <> ObjTitulo.Adiantamentos.Sum(Function(s) s.VlrABaixarOficial) Then
                        MsgBox(Me.Page, "O Valor da Baixa do adiantamento, " & ValorAdtoOficial.ToString("N2") & " não bate com a Distribuicão da Baixa dos adiantamentos.")
                        Return False
                    End If
                Else
                    If ValorAdtoMoeda <> ObjTitulo.Adiantamentos.Sum(Function(s) s.VlrABaixarMoeda) Then
                        MsgBox(Me.Page, "O Valor da Baixa do adiantamento, " & ValorAdtoMoeda.ToString("N2") & " não bate com a Distribuicão da Baixa dos adiantamentos.")
                        Return False
                    End If
                End If
            End If
        End If

        '******************************************************************************************************
        '******************************************************************************************************
        '******************************************************************************************************

        If Not lblAgrupar.Text = "AP" AndAlso DdlTributos.SelectedIndex = 0 AndAlso Not objCarteira.CodigoContaCliente Is Nothing AndAlso objCarteira.CodigoContaCliente.Trim.Length = 0 And ddlCarteiras.SelectedValue <> "999999999" And ddlCarteiras.SelectedValue <> "999999998" Then
            ddlCarteiras.SelectedIndex = 0
            MsgBox(Me.Page, "Carteira Financeira Sem Conta Contábil, Verifique.")
            Return False
        End If

        If Session("objDestinoContabil" & HID.Value) Is Nothing AndAlso DdlProvisoes.SelectedValue = 1 AndAlso objCarteira.isAdiantamento Then
            Dim strCliente() As String = txtCodigoFornecedor.Value.Split("-")
            Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
            If objCliente.DesdobrarFornecedor = True Then
                MsgBox(Me.Page, "Fornecedor do Desdobramento do Título não foi informado.")
                Return False
            End If
        End If

        If HTabela.Value = "ContasAPagar" AndAlso DdlProvisoes.SelectedValue = 1 AndAlso IIf(Trim(DdlTiposDePagamentos.Text <> ""), DdlTiposDePagamentos.SelectedValue, "0") = 4 Then 'baixa e boleto
            If ckPreImpresso.Checked = False Then
                If Trim(txtCodigoDeBarras.Text) <> "" Then
                    If CkbCodigoDeBarras.Checked Then txtCodigoDeBarras.Text = Funcoes.FormataLinhaDigitavelOriginal(txtCodigoDeBarras.Text)
                    'Deve passar a Empresa do Título para procurar em Datas não programáveis, não o Fornecedor - Furlan - 18/11/2013
                    'Dim strFornecedor As String() = txtCodigoFornecedor.Value.Split("-")
                    Dim strEmpresa As String() = DdlEmpresaCliente.SelectedValue.Split("-")
                    If Funcoes.ValidaCodigoBarras(txtCodigoDeBarras.Text, CkbCodigoDeBarras.Checked, txtProrrogacao.Text, txtValorCobrado.Text, strEmpresa(0), strEmpresa(1), Banco) = False Then
                        MsgBox(Me.Page, "Baixa Não Permitida... Codigo de Barras, Vencimento ou valor Invalido(s).")
                        Return False
                    End If
                Else
                    MsgBox(Me.Page, "Preenchimento do Codigo de Barras Obrigatório para Boletos Bancarios na Baixa.")
                    Return False
                End If
            End If
        End If

        If DdlProvisoes.SelectedValue = 1 Then
            If ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento AndAlso DdlEmpresaCliente.SelectedValue <> DdlEmpresaPagadora.SelectedValue Then
                'É feita a transferência entre contas automático - 03/07/2017 - furlan
                'If HTabela.Value = "ContasAPagar" Then
                '    MsgBox(Me.Page, "Empresa Pagadora não pode ser diferente da Empresa do titulo, em uma baixa de adiantamento.")
                'Else
                '    MsgBox(Me.Page, "Empresa Recebedora não pode ser diferente da Empresa do titulo, em uma baixa de adiantamento.")
                'End If
                'Return False
            ElseIf CDec(txtValorDoDocumento.Text) = 0 AndAlso DdlEmpresaCliente.SelectedValue <> DdlEmpresaPagadora.SelectedValue Then
                If HTabela.Value = "ContasAPagar" Then
                    MsgBox(Me.Page, "Empresa Pagadora não pode ser diferente da Empresa do titulo, com valor do documento igual a zero.")
                Else
                    MsgBox(Me.Page, "Empresa Recebedora não pode ser diferente da Empresa do titulo, com valor do documento igual a zero.")
                End If
                Return False
            ElseIf gridAdiantamentosDisponiveis.Rows.Count > 0 Then
                Dim i As Integer = 0
                While i < gridAdiantamentosDisponiveis.Rows.Count
                    Dim strData As String = gridAdiantamentosDisponiveis.Rows(i).Cells(5).Text()

                    If CDec(CType(gridAdiantamentosDisponiveis.Rows(i).FindControl("txtVlrBaixaOficial"), TextBox).Text) > 0 AndAlso CDate(txtMovimento.Text) < CDate(gridAdiantamentosDisponiveis.Rows(i).Cells(5).Text()) Then
                        MsgBox(Me.Page, "Atenção para a data da baixa do Título, está inferior a Data do Adiantamento. Verifique.")
                        Return False
                    End If

                    i += 1
                End While
            End If
        End If

        If txtRegistro.Text.Length = 0 AndAlso Not String.IsNullOrWhiteSpace(txtPedido.Text) AndAlso CInt(txtPedido.Text) > 0 Then
            Dim objpedido As New [Lib].Negocio.Pedido(campo(0), campo(1), txtPedido.Text)
            Dim ValorPedido As Decimal = 0
            Dim ValorTitulo As Decimal = 0
            For Each item As [Lib].Negocio.PedidoXItem In objpedido.Itens
                Dim Encargo As New ListPedidoXEncargo(item)
                ValorPedido += Encargo.Where(Function(L) L.CodigoEncargo = "LIQUIDO").Sum(Function(E) E.ValorOficial)
            Next
            For i As Integer = 0 To objpedido.Vencimentos.Count - 1
                ValorTitulo += objpedido.Vencimentos(i).ValorDocumentoOficial
            Next
        End If

        If Not String.IsNullOrWhiteSpace(txtRegistro.Text) AndAlso Not String.IsNullOrWhiteSpace(txtPedido.Text) _
           AndAlso CInt(txtPedido.Text) > 0 AndAlso Not Session("ControleCP" & HID.Value) Is Nothing Then

            Dim crtl As String() = Session("ControleCP" & HID.Value).ToString.Split(";")

            If crtl(0) <> txtRegistro.Text Then
                MsgBox(Me.Page, "Numero de registro de verificação invalido.")
                Return False
            End If

            If Not Funcoes.VerificaPermissao("AJUSTEFINANCEIRO", "GRAVAR") Then
                If New [Lib].Negocio.Moeda(crtl(3)).Classificacao = eTiposMoeda.Oficial Then
                    If CDec(txtValorDoDocumento.Text) > CDec(crtl(1)) AndAlso Not objCarteira.isAdiantamento Then
                        MsgBox(Me.Page, "Valor do Documento R$ não pode ser maior que o original: " & CDec(crtl(1)).ToString("N2"))
                        Return False
                    End If
                Else
                    If CDec(txtValorEmMoeda.Text) > CDec(crtl(2)) AndAlso Not objCarteira.isAdiantamento Then
                        MsgBox(Me.Page, "Valor do Documento U$ não pode ser maior que o original: " & CDec(crtl(2)).ToString("N2"))
                        Return False
                    End If
                End If
            End If
        End If

        'Quando for colocada a tela nova onde a conta será manipulada pelo user control essa validação poderá ser removida - Cleberson
        'Verificar existência do banco e conta
        'If DdlBancos.SelectedIndex > 0 AndAlso Not String.IsNullOrWhiteSpace(txtAgencia.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDigitoAgencia.Text) _
        '    AndAlso Not String.IsNullOrWhiteSpace(txtContaCorrente.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDigitoDaConta.Text) Then

        '    Dim Cli = txtCodigoFavorecido.Value.Split("-")
        '    Dim CliFav As Cliente = New Cliente(Cli(0), Cli(1))

        '    Dim ObjListaCxc As ListClienteXContaBancaria = New ListClienteXContaBancaria(CliFav)
        '    If ObjListaCxc.Where(Function(s) s.CodigoBanco = DdlBancos.SelectedValue _
        '                       AndAlso s.CodigoAgencia.ToString.Trim = txtAgencia.Text _
        '                          AndAlso s.DigitoAgencia.ToString.Trim = txtDigitoAgencia.Text _
        '                          AndAlso s.ContaCorrente.ToString.Trim = txtContaCorrente.Text _
        '                          AndAlso s.DigitoConta.ToString.Trim = txtDigitoDaConta.Text).Count <= 0 Then
        '        MsgBox(Me.Page, "A conta bancária do cliente ainda não foi gravada. Grave a conta, selecione-a e então grave o título.")
        '        TabContainer1.ActiveTabIndex = 1
        '        Return False
        '    End If
        'End If
        '------------------------------------------------------


        'Verificar se as informações como Unidade, Empresa e Cliente não tenham sido alteradas no "meio tempo" entre a consulta e baixa do título
        If Not String.IsNullOrWhiteSpace(txtPedido.Text) AndAlso CInt(txtPedido.Text) > 0 Then
            Dim sql As String
            sql = "SELECT Empresa_Id, EndEmpresa_Id, UnidadeDeNegocio, Cliente, EndCliente " & vbCrLf &
                  "  FROM Pedidos " & vbCrLf &
                  " WHERE pedido_id = " & txtPedido.Text & vbCrLf &
                  "   AND UnidadeDeNegocio = '" & DdlUnidadeDeNegocioEmpresaCliente.SelectedValue & "'" & vbCrLf &
                  "   AND Empresa_Id= '" & DdlEmpresaCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                  "   AND EndEmpresa_Id = " & DdlEmpresaCliente.SelectedValue.Split("-")(1) & vbCrLf &
                  "   AND Cliente =  '" & txtCodigoFornecedor.Value.Split("-")(0) & "'" & vbCrLf &
                  "   AND EndCliente = " & txtCodigoFornecedor.Value.Split("-")(1) & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Pedido")

            If ds.Tables("Pedido").Rows.Count <= 0 Then
                MsgBox(Me.Page, "Recarregue o título, pois houveram alterações em sua Nota ou Pedido.")
                Return False
            End If
        End If

        Return True
    End Function

    Private Function LanctosContabeis(ByVal pSequencia As Integer, Optional ByVal dMovimento As String = "")
        If Len(Raz_Conta) = 0 Or (Raz_ValorMoeda = 0 And Raz_ValorOficial = 0) Then Return False

        Sql = "INSERT INTO Razao " & vbCrLf &
              "       (Empresa_Id, " & vbCrLf &
              "       EndEmpresa_Id, " & vbCrLf &
              "       Conta_Id, " & vbCrLf &
              "       Cliente_Id, " & vbCrLf &
              "       EndCliente_Id, " & vbCrLf &
              "       Movimento_Id, " & vbCrLf &
              "       Lote_Id, " & vbCrLf &
              "       Sequencia_Id, " & vbCrLf &
              "       Titulo, " & vbCrLf &
              "       UnidadeDeNegocio, " & vbCrLf &
              "       Indexador, " & vbCrLf &
              "       DataMoeda, " & vbCrLf &
              "       DebitoOficial, " & vbCrLf &
              "       CreditoOficial, " & vbCrLf &
              "       DebitoMoeda, " & vbCrLf &
              "       CreditoMoeda, " & vbCrLf &
              "       Conciliacao, " & vbCrLf &
              "       DataDaBaixa, " & vbCrLf &
              "       Historico, " & vbCrLf &
              "       PrevistoRealizado," & vbCrLf &
              "       Processo," & vbCrLf &
              "       UsuarioInclusao," & vbCrLf &
              "       UsuarioInclusaoData)" & vbCrLf &
              "VALUES ('" & Raz_Empresa & "'," & vbCrLf &
              "         " & Raz_EndEmpresa & "," & vbCrLf &
              "        '" & Raz_Conta & "'" & vbCrLf

        If Len(Raz_Conta) = 7 Then
            Sql &= ", '" & Raz_Cliente & "'"        'Cliente
            Sql &= ", " & Raz_EndCliente            'Endereco do Cliente
        Else
            Sql &= ", ''"                           'Cliente
            Sql &= ", 0"                            'Endereco do Cliente
        End If

        If Not String.IsNullOrWhiteSpace(dMovimento) Then
            Sql &= ", '" & dMovimento & "'"     'Data de Movimento
        Else
            Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'"     'Data de Movimento
        End If

        Sql &= ", 0070"
        Sql &= ", " & pSequencia                    'Sequencia no Razao = Registro do Titulo
        Sql &= ", " & Registro                      'Numero do Titulo
        Sql &= ", '" & Raz_UnidadeDeNegocio & "'"   'Unidade de Negócio
        Sql &= ", " & ddlIndexador.SelectedValue    'Indexador
        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'"     'Data da Moeda

        'Valor Oficial
        If Raz_DebitoCredito = "D" Then
            Sql &= ", " & Replace(Raz_ValorOficial, ",", ".")  'Valor Débito Oficial
            Sql &= ", 0.0"                                     'Valor Crédito Oficial
        Else
            Sql &= ", 0.0"                                     'Valor Debito Oficial
            Sql &= ", " & Replace(Raz_ValorOficial, ",", ".")  'Valor Crédito Oficial
        End If

        If IsVariacao(DdlEmpresaCliente.SelectedValue, Raz_Conta) Then
            Sql &= ", 0.0"      'Valor Débito Moeda
            Sql &= ", 0.0"      'Valor Crédito Moeda
        Else
            If Raz_DebitoCredito = "D" Then
                Sql &= ", " & Replace(Raz_ValorMoeda, ",", ".")  'Valor Débito Moeda
                Sql &= ", 0.0"                                   'Valor Crédito Moeda
            Else
                Sql &= ", 0.0"                                   'Valor Debito Moeda
                Sql &= ", " & Replace(Raz_ValorMoeda, ",", ".")  'Valor Crédito Moeda
            End If
        End If

        If HTabela.Value = "ContasAPagar" AndAlso Raz_DebitoCredito = "C" AndAlso chkConciliado.Checked Then
            Sql &= ", 'B'"                                                              'Conciliação
            Sql &= ", '" & CDate(txtDataConciliacao.Value).ToString("yyyy/MM/dd") & "'" 'Data Conciliação

        ElseIf HTabela.Value = "ContasAReceber" AndAlso Raz_DebitoCredito = "D" AndAlso chkConciliado.Checked Then
            Sql &= ", 'B'"                                                              'Conciliação
            Sql &= ", '" & CDate(txtDataConciliacao.Value).ToString("yyyy/MM/dd") & "'" 'Data Conciliação
        Else
            Sql &= ", NULL "                                                            'Conciliação
            Sql &= ", NULL "                                                            'Data Conciliação
        End If

        Sql &= ", '" & Raz_Historico & "'"          'Histórico
        Sql &= ", 'P'"                              'Previsto/Realizado
        Sql &= ", '" & HTabela.Value & "'"                   'Processo
        Sql &= ", '" & Session("ssNomeUsuario") & "'"  'Usuario que Baixou
        Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')"           'Data da Baixa

        SqlArray.Add(Sql)

        Return True

    End Function

    Private Function IsVariacao(ByVal Empresa As String, ByVal Carteira As String, ByVal campo As String) As Boolean
        Dim sql As String = "   SELECT CASE                                                                                                        " & vbCrLf &
                            "            WHEN cart.conta" & campo & " in (empDed.contaVariacaoMonetariaAtiva, empDed.contaVariacaoMonetariaPassiva) " & vbCrLf &
                            "              THEN 1                                                                                                  " & vbCrLf &
                            "              ELSE 0                                                                                                  " & vbCrLf &
                            "          END Variacao                                                                                               " & vbCrLf &
                            "     FROM ComprasXProdutos cart,                                                                                      " & vbCrLf &
                            "         clientesxempresas empDed                                                                                     " & vbCrLf &
                            "    Where PRODUTO_ID           = '" & Carteira & "'                                                                   " & vbCrLf &
                            "      AND empDed.empresa_id    = '" & Empresa.Split("-")(0) & "'" & vbCrLf &
                            "      AND empDed.endempresa_id =  " & Empresa.Split("-")(1)

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Deducoes")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                If row("Variacao") = 1 Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function

    Private Function IsVariacao(ByVal Empresa As String, ByVal Conta As String) As Boolean
        Dim sql As String = "SELECT CASE" & vbCrLf &
                            "        WHEN '" & Conta & "' in (empDed.contaVariacaoMonetariaAtiva, empDed.contaVariacaoMonetariaPassiva, empDed.ContaVariacaoAtivaEfetiva, empDed.ContaVariacaoPassivaEfetiva) " & vbCrLf &
                            "          THEN 1" & vbCrLf &
                            "          ELSE 0" & vbCrLf &
                            "       END Variacao" & vbCrLf &
                            "  FROM clientesxempresas empDed" & vbCrLf &
                            " Where empDed.empresa_id    = '" & Empresa.Split("-")(0) & "'" & vbCrLf &
                            "   AND empDed.endempresa_id =  " & Empresa.Split("-")(1)

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Deducoes")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                If row("Variacao") = 1 Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function

    Private Sub GravaTitulo()
        If ImgValores.ImageUrl.Contains("Bloquear") Then
            MsgBox(Me.Page, "Faça o bloqueio das alterações antes de gravar.")
            Exit Sub
        End If
        SessaoRecuperaTitulo()

        Dim Empresa() As String = DdlEmpresaCliente.SelectedValue.Split("-")
        Dim Cliente() As String = txtCodigoFornecedor.Value.Split("-")
        Dim EmpresaPagadora() As String = DdlEmpresaPagadora.SelectedValue.Split("-")

        Dim objPedido As New [Lib].Negocio.Pedido()
        If txtPedido.Text.Length > 0 AndAlso CInt(txtPedido.Text) > 0 Then
            objPedido = New [Lib].Negocio.Pedido(Empresa(0), Empresa(1), txtPedido.Text)
        End If

        If objPedido.Empresa.Empresa.ObrigaNavio AndAlso txtNaviosXInvoice.Text.Length > 0 Then
            objPedido.InvoiceNavio = CInt(txtNaviosXInvoice.Text)
        End If

        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)

        Dim TributoEncargo As Encargo = Nothing
        If DdlTributos.SelectedIndex > 0 Then
            TributoEncargo = New Encargo(DdlTributos.SelectedValue)
        End If

        Dim Valor As String
        ValidaValores(False)

        SqlArray.Clear()
        If Not ValidaCampos() Then Exit Sub

        Dim DtMovimento As String = CDate(txtMovimento.Text).ToString("yyyy/MM/dd")
        Dim DtProrrogacao As String = "'" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'"
        Dim DtVencimentoOriginal As String = "'" & CDate(IIf(lblVencOriginal.Text.Length = 0, txtProrrogacao.Text, lblVencOriginal.Text)).ToString("yyyy/MM/dd") & "'"

        '*********************************************
        '*******  Gera sequencia de titulos  *********
        '*********************************************
        If txtRegistro.Text = "" Then
            Sql = "exec sp_Numerador '" & Session("ssNomeServidor").ToUpper() & "',0,1"
            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Numerador").Tables(0).Rows
                Registro = Dr("Sequencia")
                txtRegistro.Text = Registro
                ObjTitulo.Codigo = Registro
            Next
        Else
            Registro = CInt(txtRegistro.Text)

            Sql = "DELETE FROM razao" & vbCrLf &
                  " WHERE Titulo = " & Registro
            SqlArray.Add(Sql)

            Sql = "DELETE AdiantamentosXBaixas" & vbCrLf &
                  " WHERE Titulo = " & Registro
            SqlArray.Add(Sql)


            Sql = "DELETE Adiantamentos" & vbCrLf &
                  " WHERE Titulo = " & Registro
            SqlArray.Add(Sql)

            Sql = "DELETE TitulosXDesdobrarFornecedor " & vbCrLf &
                  " WHERE Registro_Id = " & Registro & vbCrLf
            SqlArray.Add(Sql)

            Sql = "DELETE FinanceiroXDocumentos " &
                  " WHERE Titulo_Id      = " & Registro & vbCrLf
            SqlArray.Add(Sql)

            Sql = "DELETE FROM " & HTabela.Value & vbCrLf &
                  " WHERE Registro_Id = " & Registro
            SqlArray.Add(Sql)

        End If


        Sql = "INSERT INTO " & HTabela.Value & vbCrLf &
              "       (Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda" & vbCrLf &
              "       ,TipoPagto, Situacao, Lote" & vbCrLf &
              "       ,Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa" & vbCrLf &
              "       ,UnidadeDeNegocio, Empresa, EndEmpresa" & vbCrLf &
              "       ,Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, TipoContaCliente, ContaContabilCliente" & vbCrLf &
              "       ,EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, TipoContaPagadora, ContaContabilPagadora" & vbCrLf &
              "       ,Cheque, Slips, Recibo, Aviso, ReciboDeposito" & vbCrLf &
              "       ,EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao" & vbCrLf &
              "       ,ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido" & vbCrLf &
              "       ,MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos, MoedaValorLiquido" & vbCrLf &
              "       ,Historico" & vbCrLf &
              "       ,CodigoDeBarras, CodigoDigitado, CodigoDeBarraPreImpresso" & vbCrLf &
              "       ,Destinatario, EndDestinatario, solicitacao" & vbCrLf &
              "       ,UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioBaixa, UsuarioBaixaData, UsuarioLiberacao, UsuarioLiberacaoData" & vbCrLf &
              "       ,Grupado" & vbCrLf &
              "       ,Observacoes" & vbCrLf &
              "       ,SituacaoBancaria, NumeroDoCheque" & vbCrLf &
              "       ,VencimentoAdto, TaxaAdto" & vbCrLf &
              "       ,UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, UsuarioLiberacaoCheque, UsuarioLiberacaoChequeDate" & vbCrLf &
              "       ,CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento, ContratoBancario)" & vbCrLf &
              "VALUES ( " & Registro & ", 0, " & DdlProvisoes.SelectedValue & ", '" & ddlCarteiras.SelectedValue & "'" & IIf(DdlTributos.Text <> "", ", '" & DdlTributos.SelectedValue & "'", ", ''") & ", " & ddlIndexador.SelectedValue & ", " & ddlMoeda.SelectedValue & vbCrLf &
              IIf(DdlTiposDePagamentos.Text <> "", ", " & DdlTiposDePagamentos.SelectedValue, ", 0") & ", 1, 70," & vbCrLf &
              "'" & CDate(txtDataEntradaSistema.Text).ToString("yyyy/MM/dd") & "'" 'Movimento

        If lblVencOriginal.Text.Length = 0 Then
            Sql &= "," & DtProrrogacao 'Vencimento
        Else
            Sql &= "," & DtVencimentoOriginal 'vencimento
        End If
        Sql &= ", " & DtProrrogacao 'prorrogacao
        Sql &= ", " & DtProrrogacao 'data moeda
        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'" 'data baixa

        Sql &= ", '" & DdlUnidadeDeNegocioEmpresaCliente.SelectedValue & "'"

        Sql &= ", '" & Empresa(0) & "', " & Empresa(1)                         'EmpresaCliente 'Endereco Empresa Cliente
        Sql &= ", '" & Cliente(0) & "', " & Cliente(1)                          'Cliente 'Endereco Cliente


        If ObjTitulo.CodigoBancoCliente > 0 Then
            Sql &= "," & ObjTitulo.CodigoBancoCliente & ", '" & ObjTitulo.CodigoAgenciaCliente & "', '" & ObjTitulo.DigitoAgenciaCliente & "', '" & ObjTitulo.ContaCliente & "', '" & ObjTitulo.DigitoContaCliente & "', '" & ObjTitulo.TipoDaContaCliente & "'" 'Banco Cliente, Agencia do Destinatario, Digito da Agencia do Destinatário, Conta Corrente do Destinatário 'Digito da Conta Corrente do Destinatário 'Tipo da Conta do Destinatário
        Else
            Sql &= ", 0, '', '', '', '', ''" 'Banco Cliente, Agencia do Destinatario, Digito da Agencia do Destinatário, Conta Corrente do Destinatário 'Digito da Conta Corrente do Destinatário 'Tipo da Conta do Destinatário
        End If


        If TributoEncargo IsNot Nothing Then
            Sql &= ", '" & TributoEncargo.ContaCredito & "'"  'Grupo de Conta Corrente do Fornecedor
        Else
            Sql &= ", '" & objCarteira.CodigoContaCliente & "'"  'Grupo de Conta Corrente do Fornecedor
        End If

        'Empresa Pagadora------------------
        If Not String.IsNullOrWhiteSpace(DdlEmpresaPagadora.SelectedValue) Then
            Sql &= ", '" & EmpresaPagadora(0) & "'"                 'Empresa Pagadora
            Sql &= ", " & EmpresaPagadora(1)                        'Endereco Empresa Pagadora
        Else
            Sql &= ", ''"                                           'Empresa Pagadora
            Sql &= ", 0"                                            'Endereco Empresa Pagadora
        End If

        If Not String.IsNullOrWhiteSpace(DdlBancoPagador.SelectedValue) Then
            Sql &= ", " & DdlBancoPagador.SelectedValue             'Banco Pagadora
        Else
            Sql &= ", 0"                                            'Banco Pagadora
        End If

        If Not String.IsNullOrWhiteSpace(DdlContaPagadora.SelectedValue) Then
            Dim conta() As String = DdlContaPagadora.SelectedValue.Split("-")
            Sql &= ", '" & conta(0) & "'"                           'Agencia Pagadora
            Sql &= ", '" & conta(1) & "'"                           'Digito Agencia Pagadora
            Sql &= ", '" & conta(2) & "'"                           'Conta Pagadora
            Sql &= ", '" & conta(3) & "'"                           'Digito Conta Pagadora
            Sql &= ", '" & conta(4) & "'"                           'Tipo da Conta Pagadora
            Sql &= ", '" & conta(5) & "'"                           'Conta Contabil
        Else
            Sql &= ", '', '', '', '', '', ''"
        End If

        '-----------------------------------------------
        If txtEmiteCheque.Value.ToString.Length = 0 Then
            Sql &= ", 'N'"                                          'Emite Cheque
        Else
            Sql &= ", '" & txtEmiteCheque.Value & "'"               'Emite Cheque
        End If

        If DdlProvisoes.SelectedValue = 1 AndAlso txtSlip.Value = "S" Then
            Sql &= ", 'S'"                                          'Emite Slips
        Else
            Sql &= ", 'N'"                                          'Emite Slips
        End If

        Sql &= ", 'N'"                                          'Emite Recibo
        Sql &= ", 'N'"                                          'Emite Aviso
        Sql &= ", 'N'"                                          'Emite Recibo De Deposito

        If String.IsNullOrWhiteSpace(txtPedido.Text) OrElse CInt(txtPedido.Text) = "0" Then                       'Pedido
            Sql &= ", NULL, NULL, NULL, NULL"
        Else
            Sql &= ", '" & Empresa(0) & "', " & Empresa(1) & ", " & txtPedido.Text

            If txtPedidoFixacao.Value.ToString.Length = 0 OrElse txtPedidoFixacao.Value = 0 Then  'PedidoFixacao
                Sql &= ", NULL"
            Else
                Sql &= ", " & txtPedidoFixacao.Value
            End If
        End If

        If String.IsNullOrWhiteSpace(txtCessaoDeCredito.Text) OrElse CInt(txtCessaoDeCredito.Text) = "0" Then
            Sql &= ", NULL"                              'Procuracao
        Else
            Sql &= ", " & txtCessaoDeCredito.Text             'Procuracao
        End If

        '---Valores em Reais ----------------------------
        Sql &= ", " & Str(txtValorDoDocumento.Text) 'Valor do Documento
        Sql &= ", " & Str(txtDescontos.Text)        'Descontos
        Sql &= ", " & Str(txtDeducoes.Text)         'Deducoes
        Sql &= ", " & Str(txtJuros.Text)            'Juros
        Sql &= ", " & Str(txtAcrescimos.Text)       'Acrescimos
        Sql &= ", " & Str(txtValorCobrado.Text)     'Liquido

        '---Valores em Dolar------------------
        Sql &= ", " & Str(txtValorEmMoeda.Text)      'Valor do Documento
        Sql &= ", " & Str(txtDescontosMoeda.Text)    'Descontos
        Sql &= ", " & Str(txtDeducoesMoeda.Text)     'Deducoes
        Sql &= ", " & Str(txtJurosMoeda.Text)        'Juros
        Sql &= ", " & Str(txtAcrescimosMoeda.Text)   'Acrescimos
        Sql &= ", " & Str(txtValorCobradoMoeda.Text) 'Liquido


        '-------------------------------------
        If txtPedido.Text.Length > 0 And txtPedido.Text <> "0" And (DdlProvisoes.SelectedValue = 2 Or DdlProvisoes.SelectedValue = 3) And Session("ssNomeLiberacao" & HID.Value).ToString.Length = 0 Then
            If ddlMoeda.SelectedValue = 1 Then
                txtHistorico.Text = PesoPago(txtPedido.Text, txtValorDoDocumento.Text, txtHistorico.Text)
            Else
                txtHistorico.Text = PesoPago(txtPedido.Text, txtValorEmMoeda.Text, txtHistorico.Text)
            End If
        End If

        Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text).ToUpper & "'"                             'Historico
        Sql &= ", '" & txtCodigoDeBarras.Text.Replace(".", "").Replace("-", "").Replace(" ", "").Replace(",", "") & "'" 'Codigo De Barras

        Sql &= "," & IIf(CkbCodigoDeBarras.Checked, "'S'", "'N'") 'Codigo De Barras 
        Sql &= "," & IIf(ckPreImpresso.Checked, "1", "0")         'Codigo De Barras PreImpresso

        campo = txtCodigoFavorecido.Value.Split("-")
        If campo(0) <> "" Then
            Sql &= ", '" & campo(0) & "'"                          'Cliente Dados Bancarios
            Sql &= ", " & campo(1)                                 'Endereco Cliente Dados Bancarios
        Else
            Sql &= ", ''"
            Sql &= ", 0"
        End If

        '--------------------------------------------------------

        If txtSolicitacao.Text = "" Then
            Sql &= ", 0"                                                            'Solicitacao
        Else
            Sql &= ", " & txtSolicitacao.Text                                       'Solicitacao
        End If

        If String.IsNullOrWhiteSpace(ddlUsuarios.SelectedValue) Then
            Sql &= ", '" & Session("ssNomeUsuario") & "'"             'Usuario Que Esta Incluindo
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"  'Data da desta Inclusao
            Sql &= ", ''"                                             'UsuarioAlteracao
            Sql &= ", ''"                                             'UsuarioData
        Else
            Dim Usu As Array = ddlUsuarios.SelectedItem.Text.Trim.Split("-")
            Sql &= ", '" & Trim(Usu(1)) & "'"                                'Usuario que Incluiu
            Sql &= ", '" & CDate(Trim(Usu(2))).ToString("yyyy/MM/dd") & "'"  'Data De Quando Ocorreu a Inclusao
            Sql &= ", '" & Session("ssNomeUsuario") & "'"              'Usuario Que Esta alterando
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"   'Data Desta Alteracao
        End If

        If DdlProvisoes.Text <> "" Then
            If DdlProvisoes.SelectedValue = 1 Then
                Sql &= ", '" & Session("ssNomeUsuario") & "'"             'Usuario que Baixou
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"  'Data da Baixa
            Else
                Sql &= ", ''"
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"             'Data da Baixa
            End If
        Else
            Sql &= ", ''"
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"                 'Data da Baixa
        End If

        If ChkLiberado.Checked = False Then
            If Funcoes.VerificaPermissao("AutorizacoesDePagamentos", "GRAVAR") Then
                Sql &= ", '" & Session("ssNomeUsuario") & "'"  'Usuario que Baixou
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"            'Data da Baixa
            Else
                Sql &= ", ''"  'Usuario que Baixou
                Sql &= ", NULL"            'Data da Baixa
            End If
        Else
            Sql &= ", '" & Session("ssNomeLiberacao" & HID.Value) & "'"                                   'Usuario que Baixou
            Sql &= ", '" & CDate(Session("ssNomeLiberacaoData" & HID.Value)).ToString("yyyy/MM/dd") & "'"  'Data da Baixa
        End If

        If lblAgrupar.Text = "AP" Then
            Sql &= ", 'M'"                                                          'Registro Grupado
            Sql &= ", '" & Session("ssObservacoes" & HID.Value) & "'"      'Observaçoes de Agrupamento
        Else
            Sql &= ", 'N'"                                                          'Registro Grupado
            Sql &= ", '" & txtObservacoes.Text & " '"                               'Observacoes
        End If

        Sql &= ", 0"                                                                'situacao bancaria 

        If txtNumeroCheque.Text.Length = 0 Then
            Sql &= ", " & 0                                                         'Número do Cheque
        Else
            Sql &= ", " & txtNumeroCheque.Text                                      'Número do Cheque
        End If

        If objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento Then
            Sql &= ",'" & CDate(txtVencimentoAdto.Text).ToSqlDate & "'," & Str(txtTaxaAdto.Text)
        Else
            Sql &= ",NULL, 0"
        End If

        If txtLiberarTitulo.Value = "S" Then
            Sql &= ", '" & Session("ssNomeUsuario") & "'"        'Usuario que Liberou Titulo
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"                  'Data da Liberação do Titulo
        Else
            If txtUsuarioLiberarTitulo.Value = "" Then
                Sql &= ", ''"                                                        'Usuario que Liberou Titulo
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"              'Data da Liberação do Titulo
            Else
                Sql &= ", '" & txtUsuarioLiberarTitulo.Value & "'"                                  'Usuario que Liberou Titulo
                Sql &= ", '" & CDate(txtUsuarioLiberarTituloData.Value).ToString("yyyy/MM/dd") & "'" 'Data da Liberação do Titulo
            End If
        End If

        If txtLiberarPedido.Value = "S" Then
            Sql &= ", '" & Session("ssNomeUsuario") & "'"        'Usuario que Liberou Pedido
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"                  'Data da Liberação do Pedido
        Else
            If txtUsuarioLiberarPedido.Value = "" Then
                Sql &= ", ''"                                                        'Usuario que Liberou Pedido
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"              'Data da Liberação do Pedido
            Else
                Sql &= ", '" & txtUsuarioLiberarPedido.Value & "'"                                  'Usuario que Liberou Pedido
                Sql &= ", '" & CDate(txtUsuarioLiberarPedidoData.Value).ToString("yyyy/MM/dd") & "'" 'Data da Liberação do Pedido
            End If
        End If

        If txtLiberarCheque.Value = "S" Then
            Sql &= ", '" & Session("ssNomeUsuario") & "'"        'Usuario que Liberou Cheque
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"                  'Data da Liberação do Cheque
        Else
            If txtUsuarioLiberarCheque.Value = "" Then
                Sql &= ", ''"                                                        'Usuario que Liberou Pedido
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"              'Data da Liberação do Pedido
            Else
                Sql &= ", '" & txtUsuarioLiberarCheque.Value & "'"                                  'Usuario que Liberou Pedido
                Sql &= ", '" & CDate(txtUsuarioLiberarChequeData.Value).ToString("yyyy/MM/dd") & "'" 'Data da Liberação do Pedido
            End If
        End If

        Sql &= ", ''"                                 'Carteira de Adiantamento nao existe mais

        Sql &= ", " & ddlCarteiraDoTitulo.SelectedValue & ""                         'Carteira do Titulo 
        Sql &= ", '" & Trim(txtContratoFinanceiro.Text) & "'"                        'Contrato De Financiamento 
        Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(Trim(txtContratoBanco.Text)) & "')" 'Contrato Bancário

        SqlArray.Add(Sql)

        If CDbl(txtValorCobrado.Text) <> txtValorLiquido.Value Then
            chkConciliado.Checked = False
            txtDataConciliacao.Value = ""
        End If


        '*********************************************************************************************
        '****************   GRAVA NOVO TITULO DE PAGAMENTOS PARCIAIS   *******************************
        '*********************************************************************************************
        Dim MensagemParcial As String = ""
        If objPedido.Codigo > 0 OrElse CInt(txtCodigoFaturaDeFrete.Text) > 0 Then
            If ddlMoeda.SelectedValue = 1 AndAlso HDValorOriginalOficial.Value > 0 AndAlso CDec(txtValorCobrado.Text) <> HDValorOriginalOficial.Value Then
                If CDbl(txtValorDoDocumento.Text) < HDValorOriginalOficial.Value And (HDValorOriginalOficial.Value - CDbl(txtValorDoDocumento.Text) > 0) Then
                    GravaTituloParcial(MensagemParcial, txtRegistro.Text, objPedido)
                End If
            End If

            If ddlMoeda.SelectedValue = 3 AndAlso HDValorOriginalMoeda.Value > 0 And CDec(txtValorEmMoeda.Text) <> HDValorOriginalMoeda.Value Then
                If CDbl(txtValorEmMoeda.Text) < HDValorOriginalMoeda.Value And (HDValorOriginalMoeda.Value - CDbl(txtValorEmMoeda.Text) > 0) Then
                    GravaTituloParcial(MensagemParcial, txtRegistro.Text, objPedido)
                End If
            End If
        End If
        '*********************************************************************************************
        '*********************************************************************************************
        '*********************************************************************************************

        If DdlProvisoes.SelectedValue = 1 AndAlso Not lblAgrupar.Text = "AP" Then  'Adiantamento e AdiantamentoXBaixa
            Dim VlrBrutoOficial As Decimal = CDec(txtValorDoDocumento.Text)
            Dim VlrLiquidoOficial As Decimal = CDec(txtValorCobrado.Text)

            Dim VlrBrutoMoeda As Decimal = CDec(txtValorEmMoeda.Text)
            Dim VlrLiquidoMoeda As Decimal = CDec(txtValorCobradoMoeda.Text)

            If objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento AndAlso Not lblAgrupar.Text = "AP" Then
                If Not Adiantamento(objPedido, txtRegistro.Text, CDate(txtVencimentoAdto.Text)) Then
                    MsgBox(Me.Page, Mensagem)
                    Exit Sub
                End If
            End If

            If objCarteira.BaixaAdiantamento Or (ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento) Then
                If Left(ObjTitulo.ContaContabilPagadora, 1) = IIf(HTabela.Value = "ContasAPagar", "2", "1") AndAlso Not objCarteira.BaixaAdiantamento Then
                    'Continua sem fazer nada pois está dando adto ao invés de consumir - Furlan - 09/05/2020
                Else
                    For Each row In ObjTitulo.Adiantamentos.Where(Function(s) s.VlrABaixarOficial > 0 OrElse (s.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso s.ValorMoeda > 0))
                        If Not AdiantamentoAmortizacao(txtRegistro.Text, row.CodigoTitulo, row.Codigo, row.VlrABaixarOficial, row.VlrABaixarMoeda, row.VlrCalcVariacao) Then
                            MsgBox(Me.Page, Mensagem)
                            Exit Sub
                        End If
                    Next
                End If
            End If
        End If

        'AGRUPAMENTO DE TITULOS
        If lblAgrupar.Text = "AP" Then

            Dim vtAgruado As Decimal

            For index = 0 To Session("ssRegistros" & HID.Value).Count - 1

                Dim tit As New Titulo(Session("ssRegistros" & HID.Value).Item(index))

                Sql = "SELECT CP.Registro_Id, CP.Sequencia_Id, CP.Provisao, CP.Carteira, CP.Tributo, CP.Indexador, " & vbCrLf &
                      "       CP.Moeda, CP.TipoPagto, CP.Situacao, CP.Lote, CP.Movimento, CP.Vencimento, CP.Prorrogacao, " & vbCrLf &
                      "       CP.DataMoeda, isnull(CP.Baixa,CP.Prorrogacao) as Baixa, CP.UnidadeDeNegocio, CP.Empresa, CP.EndEmpresa, CP.Cliente, " & vbCrLf &
                      "       CP.EndCliente, CP.BancoCliente, CP.AgenciaCliente, CP.DigitoAgenciaCliente, CP.ContaCliente, " & vbCrLf &
                      "       CP.DigitoContaCliente, Isnull(CP.TipoContaCliente,'C') AS TipoContaCliente, CP.ContaContabilCliente, " & vbCrLf &
                      "       CP.EmpresaPagadora, CP.EndEmpresaPagadora, CP.BancoPagador, CP.AgenciaPagadora, CP.DigitoAgenciaPagadora, " & vbCrLf &
                      "       CP.ContaPagadora, CP.DigitoContaPagadora, CP.ContaContabilPagadora, CP.Cheque, isnull(CP.Slips,'N') AS Slips, " & vbCrLf &
                      "       CP.Recibo, CP.Aviso, CP.ReciboDeposito, isnull(CP.EmpresaPedido,'') AS EmpresaPedido, isnull(CP.EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(CP.Pedido, 0) AS Pedido, " & vbCrLf &
                      "       isnull(CP.PedidoFixacao,0) AS PedidoFixacao, isnull(CP.Procuracao,0) AS Procuracao, CP.ValorDoDocumento, CP.Descontos, CP.Deducoes, " & vbCrLf &
                      "       CP.Juros, CP.Acrescimos, CP.ValorLiquido, ISNULL(CP.MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, " & vbCrLf &
                      "       ISNULL(CP.MoedaDescontos, 0) AS MoedaDescontos, ISNULL(CP.MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(CP.MoedaJuros, 0) AS MoedaJuros, " & vbCrLf &
                      "       ISNULL(CP.MoedaAcrescimos, 0) AS MoedaAcrescimos, ISNULL(CP.MoedaValorLiquido, 0) AS MoedaValorLiquido, CP.Historico, " & vbCrLf &
                      "       CP.CodigoDeBarras, CP.CodigoDigitado, CP.Destinatario, CP.EndDestinatario, CP.NomeDoDestinatario, CP.Destinacao, " & vbCrLf &
                      "       CP.Solicitacao, CP.UsuarioInclusao, CP.UsuarioInclusaoData, CP.UsuarioAlteracao, CP.UsuarioAlteracaoData, " & vbCrLf &
                      "       CP.UsuarioCancelamento, CP.UsuarioCancelamentoData, isnull(CP.UsuarioLiberacao,'') AS UsuarioLiberacao, CP.UsuarioLiberacaoData, " & vbCrLf &
                      "       CP.UsuarioBaixa, CP.UsuarioBaixaData, isnull(CP.Grupado,'N') AS Grupado, isnull(CP.RegistroMestre, 0) as RegistroMestre, CP.Observacoes, " & vbCrLf &
                      "       CP.SituacaoBancaria, ISNULL(CP.NumeroDoCheque,0) AS NumeroDoCheque, isnull(CP.Adiantamento,0) AS Adiantamento, CP.VencimentoAdto, CP.TaxaAdto, " & vbCrLf &
                      "       isnull(CP.UsuarioLiberacaoBloqueio, '') AS UsuarioLiberacaoBloqueio, CP.UsuarioLiberacaoBloqueioDate, isnull(CP.UsuarioLiberacaoPedido, '') AS UsuarioLiberacaoPedido, " & vbCrLf &
                      "       CP.UsuarioLiberacaoPedidoDate, isnull(CP.UsuarioLiberacaoCheque, '') AS UsuarioLiberacaoCheque, CP.UsuarioLiberacaoChequeDate, " & vbCrLf &
                      "       isnull(CP.CarteiraAdto,'') AS CarteiraAdto, isnull(CP.CarteiraDoTitulo,0) AS CarteiraDoTitulo, isnull(CP.ContratoDeFinanciamento,'') AS ContratoDeFinanciamento, " & vbCrLf &
                      "       ISNULL(NF.TipoDeDocumento,0) AS TipoDeDocumento, ISNULL(FFxT.Fatura_Id, 0) AS FaturaDeFrete" & vbCrLf &
                      "  FROM " & HTabela.Value & " CP " & vbCrLf &
                      "  LEFT JOIN NotaFiscalXTitulo NFxT " & vbCrLf &
                      "	   ON CP.Registro_Id = NFxT.Titulo_Id " & vbCrLf &
                      "	   And CP.Empresa = NFxT.Empresa_Id " & vbCrLf &
                      "	 LEFT JOIN NotasFiscais NF " & vbCrLf &
                      "	   ON NFxT.Empresa_Id      = NF.Empresa_Id " & vbCrLf &
                      "   AND NFxT.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf &
                      "   AND NFxT.Cliente_Id      = NF.Cliente_Id " & vbCrLf &
                      "   AND NFxT.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf &
                      "   AND NFxT.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
                      "   AND NFxT.Serie_Id        = NF.Serie_Id " & vbCrLf &
                      "   AND NFxT.Nota_Id         = NF.Nota_Id" & vbCrLf &
                      "  LEFT JOIN FaturaDeFreteXTitulo FFxT " & vbCrLf &
                      "    ON CP.Registro_Id = FFxT.Titulo_Id " & vbCrLf &
                      " WHERE CP.Registro_Id = " & CStr(Session("ssRegistros" & HID.Value).Item(index)) & vbCrLf &
                      "   and CP.Situacao not in (2,3,4,5,6,10) " & vbCrLf

                Dim dsFilho As New DataSet
                dsFilho = Banco.ConsultaDataSet(Sql, "TituloXFilho")

                If Not dsFilho Is Nothing AndAlso dsFilho.Tables(0).Rows.Count > 0 Then

                    'ACRESCENTEI PARA RESOLVER PROBLEMAS QUE VEM APARECENDO DE MESTRE DIFERENTE DO VALOR DOS FILHOS - FURLAN - 03/11/2022.

                    vtAgruado += dsFilho.Tables(0).Compute("SUM(ValorLiquido)", "")

                    For Each drFilho As DataRow In dsFilho.Tables(0).Rows
                        Sql = " DELETE FROM razao" & vbCrLf &
                              "  WHERE Titulo = " & CStr(Session("ssRegistros" & HID.Value).Item(index)) & vbCrLf
                        SqlArray.Add(Sql)

                        Sql = " DELETE FROM Adiantamentos" & vbCrLf &
                              "  WHERE Titulo = " & CStr(Session("ssRegistros" & HID.Value).Item(index)) & vbCrLf
                        SqlArray.Add(Sql)

                        Sql = " DELETE FROM AdiantamentosXBaixas" & vbCrLf &
                              "  WHERE Titulo = " & CStr(Session("ssRegistros" & HID.Value).Item(index)) & vbCrLf

                        SqlArray.Add(Sql)

                        Sql = " UPDATE " & HTabela.Value & vbCrLf &
                              "    SET Provisao = " & DdlProvisoes.SelectedValue & vbCrLf &
                              "        ,TipoPagto   = '" & DdlTiposDePagamentos.SelectedValue & "'" & vbCrLf &
                              "        ,DataMoeda   = '" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf &
                              "        ,Baixa       = '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf &
                              "        ,Prorrogacao = '" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf

                        'If DdlTributos.SelectedIndex > 0 Then
                        '    Sql &= ", Tributo = '" & DdlTributos.SelectedValue & "'" & vbCrLf
                        'End If

                        If CInt(drFilho("FaturaDeFrete")) > 0 Then
                            If Not ddlIndexador.SelectedValue = 99 Then
                                drFilho("MoedaValorDoDocumento") = DolarizaBaixa(DtMovimento, drFilho("ValorDoDocumento"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                                If CDec(drFilho("Descontos")) > 0 Then drFilho("MoedaDescontos") = DolarizaBaixa(DtMovimento, drFilho("Descontos"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                                If CDec(drFilho("Deducoes")) > 0 Then drFilho("MoedaDeducoes") = DolarizaBaixa(DtMovimento, drFilho("Deducoes"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                                If CDec(drFilho("Juros")) > 0 Then drFilho("MoedaJuros") = DolarizaBaixa(DtMovimento, drFilho("Juros"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                                If CDec(drFilho("Acrescimos")) > 0 Then drFilho("MoedaAcrescimos") = DolarizaBaixa(DtMovimento, drFilho("Acrescimos"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                                drFilho("MoedaValorLiquido") = drFilho("MoedaValorDoDocumento") + drFilho("MoedaJuros") + drFilho("MoedaAcrescimos") - drFilho("MoedaDescontos") - drFilho("MoedaDeducoes")
                            End If
                        ElseIf drFilho("Moeda") = 1 Then
                            If Not ddlIndexador.SelectedValue = 99 Then
                                If CDec(drFilho("MoedaValorDoDocumento")) = 0 Then drFilho("MoedaValorDoDocumento") = DolarizaBaixa(DtMovimento, drFilho("ValorDoDocumento"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                                If CDec(drFilho("Descontos")) > 0 Then drFilho("MoedaDescontos") = DolarizaBaixa(DtMovimento, drFilho("Descontos"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                                If CDec(drFilho("Deducoes")) > 0 Then drFilho("MoedaDeducoes") = DolarizaBaixa(DtMovimento, drFilho("Deducoes"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                                If CDec(drFilho("Juros")) > 0 Then drFilho("MoedaJuros") = DolarizaBaixa(DtMovimento, drFilho("Juros"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                                If CDec(drFilho("Acrescimos")) > 0 Then drFilho("MoedaAcrescimos") = DolarizaBaixa(DtMovimento, drFilho("Acrescimos"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                                drFilho("MoedaValorLiquido") = drFilho("MoedaValorDoDocumento") + drFilho("MoedaJuros") + drFilho("MoedaAcrescimos") - drFilho("MoedaDescontos") - drFilho("MoedaDeducoes")
                            End If
                        End If
                        'Valor Documento Moeda
                        Sql &= ", MoedaValorDoDocumento = " & Replace(drFilho("MoedaValorDoDocumento"), ",", ".")
                        'Descontos
                        Sql &= ", MoedaDescontos = " & Replace(drFilho("MoedaDescontos"), ",", ".")
                        'Deducoes
                        Sql &= ", MoedaDeducoes = " & Replace(drFilho("MoedaDeducoes"), ",", ".")
                        'Juros
                        Sql &= ", MoedaJuros = " & Replace(drFilho("MoedaJuros"), ",", ".")
                        'Acrescimos
                        Sql &= ", MoedaAcrescimos = " & Replace(drFilho("MoedaAcrescimos"), ",", ".")
                        'Liquido
                        Sql &= ", MoedaValorLiquido = " & Replace(drFilho("MoedaValorLiquido"), ",", ".")

                        Sql &= ", RegistroMestre = " & txtRegistro.Text
                        Sql &= ", Grupado = 'S'"

                        If Not String.IsNullOrWhiteSpace(DdlEmpresaPagadora.SelectedValue) Then
                            Sql &= ", EmpresaPagadora = '" & EmpresaPagadora(0) & "', EndEmpresaPagadora = " & EmpresaPagadora(1)  'EmpresaPagadora 'Endereco Empresa Pagadora"
                        End If

                        If DdlBancoPagador.SelectedIndex <> 0 Then
                            Sql &= ", BancoPagador = '" & DdlBancoPagador.SelectedValue & "'"          'Banco
                            Dim Conta() As String = DdlContaPagadora.SelectedValue.Split("-")
                            '(0)Agencia Cliente (1)Digito Agencia Cliente (2)Conta Cliente (3)Digito Conta Cliente (4) Tipo conta Pagadora (5)Conta Contabil
                            Sql &= ", AgenciaPagadora       = '" & Conta(0) & "', DigitoAgenciaPagadora = '" & Conta(1) & "', ContaPagadora         = '" & Conta(2) & "', DigitoContaPagadora   = '" & Conta(3) & "', TipoContaPagadora = '" & Conta(4) & "', ContaContabilPagadora = '" & Conta(5) & "'"
                        End If

                        'If DdlBancos.Text <> "" Then
                        '    Sql &= ", BancoCliente         = " & DdlBancos.SelectedValue               'Banco Cliente
                        '    Sql &= ", AgenciaCliente       ='" & txtAgencia.Text.Trim & "'"            'Agencia do Destinatario
                        '    Sql &= ", DigitoAgenciaCliente ='" & txtDigitoAgencia.Text.Trim & "'"      'Digito da Agencia do Destinatário
                        '    Sql &= ", ContaCliente         ='" & txtContaCorrente.Text.Trim & "'"      'Conta Corrente do Destinatário
                        '    Sql &= ", DigitoContaCliente   ='" & txtDigitoDaConta.Text.Trim & "'"      'Digito da Conta Corrente do Destinatário
                        '    Sql &= ", TipoContaCliente     ='" & ddlTipoConta.SelectedValue & "'"      'Tipo da Conta do Destinatário
                        'End If

                        Sql &= ", UsuarioAlteracao = '" & Session("ssNomeUsuario") & "'"                 'Usuario que Incluiu
                        Sql &= ", UsuarioAlteracaoData = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"  'Data da Inclusao
                        Sql &= ", ContratoBancario = '" & Funcoes.EliminarCaracteresEspeciais(Trim(txtContratoBanco.Text)) & "'" 'Contrato Bancário

                        Sql &= " WHERE Registro_ID = " & CStr(Session("ssRegistros" & HID.Value).Item(index))

                        SqlArray.Add(Sql)

                        If DdlProvisoes.SelectedValue = 1 Then
                            Dim objPedidoAP As New [Lib].Negocio.Pedido()
                            Dim objCarteiraAP As New [Lib].Negocio.CarteiraFinanceira(drFilho("Carteira"))
                            Dim objCarteiraAdtoAP As New [Lib].Negocio.CarteiraFinanceira(drFilho("CarteiraAdto"))

                            If drFilho("Pedido") > 0 Then
                                objPedidoAP = New [Lib].Negocio.Pedido(drFilho("EmpresaPedido"), drFilho("EndEmpresaPedido"), drFilho("Pedido"))
                            End If

                            If DdlProvisoes.SelectedValue = 1 Then  'Adiantamento e AdiantamentoXBaixa
                                Dim VlrBrutoOficial As Decimal = CDec(drFilho("ValorDoDocumento"))
                                Dim VlrLiquidoOficial As Decimal = CDec(drFilho("ValorDoDocumento")) + CDec(drFilho("Juros")) + CDec(drFilho("Acrescimos")) - CDec(drFilho("Descontos")) - CDec(drFilho("Deducoes"))

                                Dim VlrBrutoMoeda As Decimal = CDec(drFilho("MoedaValorDoDocumento"))
                                Dim VlrLiquidoMoeda As Decimal = CDec(drFilho("MoedaValorDoDocumento")) + CDec(drFilho("MoedaJuros")) + CDec(drFilho("MoedaAcrescimos")) - CDec(drFilho("MoedaDescontos")) - CDec(drFilho("MoedaDeducoes"))

                                If objCarteiraAP.isAdiantamento And Not objCarteiraAP.BaixaAdiantamento Then
                                    If Not Adiantamento(objPedidoAP, drFilho("Registro_Id"), drFilho("VencimentoAdto")) Then
                                        MsgBox(Me.Page, Mensagem)
                                        Exit Sub
                                    End If
                                End If

                                '#01062016 - Verificar Baixa de Adiantamento de titulos agrupados
                                If objCarteiraAP.BaixaAdiantamento Or objCarteiraAdtoAP.BaixaAdiantamento Then
                                    If Left(ObjTitulo.ContaContabilPagadora, 1) = IIf(HTabela.Value = "ContasAPagar", "2", "1") Then
                                        'Continua sem fazer nada pois está dando adto ao invés de consumir - Furlan - 09/05/2020
                                    Else
                                        For Each row In ObjTitulo.Adiantamentos.Where(Function(s) s.VlrABaixarOficial > 0)
                                            If Not AdiantamentoAmortizacao(drFilho("Registro_Id"), row.CodigoTitulo, row.Codigo, row.VlrABaixarOficial, row.VlrABaixarMoeda, row.VlrCalcVariacao) Then
                                                MsgBox(Me.Page, Mensagem)
                                                Exit Sub
                                            End If
                                        Next
                                    End If
                                End If
                            End If

                            '****************************************************************************************************************************************************************************************
                            '********************************************* Gravacao do Razão do Titulos Filho ******************************************************************************************************
                            '****************************************************************************************************************************************************************************************

                            Dim Carteira As New CarteiraFinanceira(drFilho("Carteira"))

                            '**************************************************************************************************
                            '***************** Valor Do Documento *************************************************************
                            '**************************************************************************************************
                            Registro = drFilho("Registro_Id")
                            Raz_Empresa = drFilho("Empresa")                     'EmpresaCliente
                            Raz_EndEmpresa = drFilho("EndEmpresa")               'Endereco Empresa Cliente

                            Raz_Conta = Carteira.CodigoContaCliente                            'Conta sem tributo
                            Dim objPlaConta As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                            If objPlaConta.TemCliente Then
                                Raz_Cliente = drFilho("Cliente")                'Cliente
                                Raz_EndCliente = drFilho("EndCliente")
                            Else
                                Raz_Cliente = ""                                'Cliente
                                Raz_EndCliente = 0                              'Endereco do Cliente
                            End If

                            If drFilho("Tributo") <> "" Then
                                Dim Tributo As New Encargo(drFilho("Tributo"))
                                Raz_Conta = Tributo.ContaDebito                       'Conta com tributo
                                Dim objPlaContaTributo As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                                If objPlaContaTributo.TemCliente Then

                                    Raz_Cliente = drFilho("Cliente")                'Cliente
                                    Raz_EndCliente = drFilho("EndCliente")
                                Else
                                    Raz_Cliente = ""                                'Cliente
                                    Raz_EndCliente = 0                              'Endereco do Cliente
                                End If
                            End If

                            Raz_UnidadeDeNegocio = drFilho("UnidadeDeNegocio")   'Unidade De Negócio

                            Raz_ValorOficial = drFilho("ValorDoDocumento")       'ValorDoDocumento

                            If drFilho("MoedaValorDoDocumento") > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                                Raz_ValorMoeda = drFilho("MoedaValorDoDocumento")
                            Else
                                Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, drFilho("ValorDoDocumento"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                            End If

                            Raz_Historico = Funcoes.EliminarCaracteresEspeciais((drFilho("Historico").ToString.Trim & ". " & drFilho("Observacoes").ToString.Trim))                 'Historico
                            Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "D", "C")                         'Debito/Credito

                            LanctosContabeis(Registro)

                            '**************************************************************************************************
                            '********************************  Descontos ******************************************************
                            '**************************************************************************************************

                            If drFilho("Descontos") > 0 Then
                                Registro = drFilho("Registro_Id")
                                Raz_Empresa = drFilho("Empresa")                     'EmpresaCliente
                                Raz_EndEmpresa = drFilho("EndEmpresa")               'Endereco Empresa Cliente
                                Raz_Conta = Carteira.CodigoContaDesconto             'Grupo de Contas

                                Dim objPlaContaDescontos As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                                If objPlaContaDescontos.TemCliente Then
                                    Raz_Cliente = drFilho("Cliente")                'Cliente
                                    Raz_EndCliente = drFilho("EndCliente")
                                Else
                                    Raz_Cliente = ""                                'Cliente
                                    Raz_EndCliente = 0                              'Endereco do Cliente
                                End If

                                Raz_UnidadeDeNegocio = drFilho("UnidadeDeNegocio")   'Unidade De Negócio

                                Raz_ValorOficial = drFilho("Descontos")                  'Descontos

                                If drFilho("MoedaDescontos") > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                                    Raz_ValorMoeda = drFilho("MoedaDescontos")
                                Else
                                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, drFilho("Descontos"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                                End If

                                Raz_Historico = Funcoes.EliminarCaracteresEspeciais((drFilho("Historico").ToString.Trim & ". " & drFilho("Observacoes").ToString.Trim))                'Historico
                                Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "C", "D")                        'Debito/Credito

                                LanctosContabeis(Registro)
                            End If

                            '**************************************************************************************************
                            '*********************************   Deducoes  ****************************************************
                            '**************************************************************************************************
                            If drFilho("Deducoes") > 0 Then
                                Registro = drFilho("Registro_Id")
                                Raz_Empresa = drFilho("Empresa")                     'EmpresaCliente
                                Raz_EndEmpresa = drFilho("EndEmpresa")               'Endereco Empresa Cliente
                                Raz_Conta = Carteira.CodigoContaDeducao              'Grupo de Contas

                                Dim objPlaContaDeducoes As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                                If objPlaContaDeducoes.TemCliente Then
                                    Raz_Cliente = drFilho("Cliente")                'Cliente
                                    Raz_EndCliente = drFilho("EndCliente")
                                Else
                                    Raz_Cliente = ""                                'Cliente
                                    Raz_EndCliente = 0                              'Endereco do Cliente
                                End If

                                Raz_UnidadeDeNegocio = drFilho("UnidadeDeNegocio")   'Unidade De Negócio
                                Raz_ValorOficial = drFilho("Deducoes")               'Deducoes

                                If drFilho("MoedaDeducoes") > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                                    Raz_ValorMoeda = drFilho("MoedaDeducoes")
                                Else
                                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, drFilho("Deducoes"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                                End If

                                Raz_Historico = Funcoes.EliminarCaracteresEspeciais((drFilho("Historico").ToString.Trim & ". " & drFilho("Observacoes").ToString.Trim))                 'Historico
                                Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "C", "D")                         'Debito/Credito
                                LanctosContabeis(Registro)
                            End If

                            '**************************************************************************************************
                            '***************************************   Juros   ************************************************
                            '**************************************************************************************************

                            If drFilho("Juros") > 0 Then
                                Registro = drFilho("Registro_Id")
                                Raz_Empresa = drFilho("Empresa")                     'EmpresaCliente
                                Raz_EndEmpresa = drFilho("EndEmpresa")               'Endereco Empresa Cliente
                                Raz_Conta = Carteira.CodigoContaJuro                       'Grupo de Contas

                                Dim objPlaContaJuros As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                                If objPlaContaJuros.TemCliente Then
                                    Raz_Cliente = drFilho("Cliente")                'Cliente
                                    Raz_EndCliente = drFilho("EndCliente")
                                Else
                                    Raz_Cliente = ""                                'Cliente
                                    Raz_EndCliente = 0                              'Endereco do Cliente
                                End If

                                Raz_UnidadeDeNegocio = drFilho("UnidadeDeNegocio")   'Unidade De Negócio

                                Raz_ValorOficial = drFilho("Juros")                     'Juros

                                If drFilho("MoedaJuros") > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                                    Raz_ValorMoeda = drFilho("MoedaJuros")
                                Else
                                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, drFilho("Juros"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                                End If

                                Raz_Historico = Funcoes.EliminarCaracteresEspeciais((drFilho("Historico").ToString.Trim & ". " & drFilho("Observacoes").ToString.Trim))                 'Historico
                                Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "D", "C")                          'Debito/Credito
                                LanctosContabeis(Registro)
                            End If

                            '**************************************************************************************************
                            '*************************************  Acrescimos  ***********************************************
                            '**************************************************************************************************

                            If drFilho("Acrescimos") <> 0 Then
                                Registro = drFilho("Registro_Id")
                                Raz_Empresa = drFilho("Empresa")                     'EmpresaCliente
                                Raz_EndEmpresa = drFilho("EndEmpresa")               'Endereco Empresa Cliente
                                Raz_Conta = Carteira.CodigoContaAcrescimo            'Grupo de Contas

                                Dim objPlaContaAcrescimos As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                                If objPlaContaAcrescimos.TemCliente Then
                                    Raz_Cliente = drFilho("Cliente")                'Cliente
                                    Raz_EndCliente = drFilho("EndCliente")
                                Else
                                    Raz_Cliente = ""                                'Cliente
                                    Raz_EndCliente = 0                              'Endereco do Cliente
                                End If

                                Raz_UnidadeDeNegocio = drFilho("UnidadeDeNegocio")   'Unidade De Negócio
                                Raz_ValorOficial = drFilho("Acrescimos")                    'Acrescimos

                                If drFilho("MoedaAcrescimos") > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                                    Raz_ValorMoeda = drFilho("MoedaAcrescimos")
                                Else
                                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, drFilho("Acrescimos"), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                                End If

                                Raz_Historico = Funcoes.EliminarCaracteresEspeciais((drFilho("Historico").ToString.Trim & ". " & drFilho("Observacoes").ToString.Trim))                'Historico
                                Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "D", "C")                          'Debito/Credito

                                LanctosContabeis(Registro)
                            End If
                        End If
                    Next
                End If
            Next

            'Não deixar total do agrupamento ficar diferente do total de títulos agrupados - Furlan - 16/11/2022
            If Not vtAgruado = CDec(txtValorCobrado.Text) AndAlso ddlMoeda.SelectedValue = 1 Then
                MsgBox(Me.Page, "O valor do agrupamento " & txtValorCobrado.Text & " está divergente dos titulos agrupados " & vtAgruado.ToString & ". Limpe a tela e refaça o processo.")
                Exit Sub
            End If

            'Devolve titulo da tela para o Registro - Furlan - 21/02/2022
            Registro = txtRegistro.Text
        End If

        '****************************************************************************************************************************************************************************************
        '********************************************* Gravacao do Razão do Titulo / Mestre *****************************************************************************************************
        '****************************************************************************************************************************************************************************************

        Dim Carteira2 As New CarteiraFinanceira(ddlCarteiras.SelectedValue)

        If DdlProvisoes.SelectedValue = 1 And lblAgrupar.Text = "" Then
            '**************************************************************************************************
            '*******************************  Valor do Documento  *********************************************
            '**************************************************************************************************
            If Carteira2.BaixaAdiantamento Then
                ContabilizaBaixaAdiantamento(Carteira2.CodigoContaCliente, Registro)
            Else
                'Registro = CInt(txtRegistro.Text)
                Raz_Empresa = Empresa(0)                                  'EmpresaCliente
                Raz_EndEmpresa = Empresa(1)                               'Endereco Empresa Cliente

                Raz_Conta = Carteira2.CodigoContaCliente            'Conta sem tributo
                Dim objPlaConta As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                If objPlaConta.TemCliente Then
                    campo = txtCodigoFornecedor.Value.Split("-")
                    Raz_Cliente = campo(0)                              'Cliente
                    Raz_EndCliente = campo(1)
                Else
                    Raz_Cliente = ""                                'Cliente
                    Raz_EndCliente = 0                              'Endereco do Cliente
                End If

                If DdlTributos.Text <> "" Then
                    'ConsultaTributos(DdlTributos.SelectedValue)    'Consulta Contas Contabeis dos Tributos
                    Dim Encargo2 As New Encargo(DdlTributos.SelectedValue)
                    Raz_Conta = Encargo2.ContaDebito                 'Conta com tributo
                    Dim objPlaContaTributo As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                    If objPlaContaTributo.TemCliente Then
                        campo = txtCodigoFornecedor.Value.Split("-")
                        Raz_Cliente = campo(0)                              'Cliente
                        Raz_EndCliente = campo(1)
                    Else
                        Raz_Cliente = ""                                'Cliente
                        Raz_EndCliente = 0                              'Endereco do Cliente
                    End If
                End If

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue
                Raz_ValorOficial = Replace(txtValorDoDocumento.Text, ".", "")

                If CDec(txtValorEmMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtValorEmMoeda.Text, ".", "")
                Else
                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, txtValorEmMoeda.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                End If

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)

                Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "D", "C")


                LanctosContabeis(Registro)
            End If


            '**************************************************************************************************
            '**************************************  Descontos  ***********************************************
            '**************************************************************************************************

            If CDec(txtDescontos.Text) > 0 Then
                'Registro = CInt(txtRegistro.Text)
                Raz_Empresa = Empresa(0)                  'EmpresaCliente
                Raz_EndEmpresa = Empresa(1)               'Endereco Empresa Cliente

                Raz_Conta = Carteira2.CodigoContaDesconto             'Conta de Descontos Obtidos

                Dim objPlaContaDesconto As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                If objPlaContaDesconto.TemCliente Then
                    campo = txtCodigoFornecedor.Value.Split("-")
                    Raz_Cliente = campo(0)                              'Cliente
                    Raz_EndCliente = campo(1)
                Else
                    Raz_Cliente = ""                                'Cliente
                    Raz_EndCliente = 0                              'Endereco do Cliente
                End If

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'Unidade de Negocio do Titulo

                Raz_ValorOficial = Replace(txtDescontos.Text, ".", "")  'Descontos

                If CDec(txtDescontosMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtDescontosMoeda.Text, ".", "")
                Else
                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, txtDescontos.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                End If

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)
                'Históricos
                Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "C", "D")

                LanctosContabeis(Registro)

            End If


            '**************************************************************************************************
            '**************************************  Deducoes  ************************************************
            '**************************************************************************************************

            If CDec(txtDeducoes.Text) > 0 Then
                'Registro = CInt(txtRegistro.Text)
                Raz_Empresa = Empresa(0)                  'EmpresaCliente
                Raz_EndEmpresa = Empresa(1)               'Endereco Empresa Cliente

                Raz_Conta = Carteira2.CodigoContaDeducao   'Conta de Descontos Obtidos

                Dim objPlaContaDeducoes As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                If objPlaContaDeducoes.TemCliente Then
                    campo = txtCodigoFornecedor.Value.Split("-")
                    Raz_Cliente = campo(0)                              'Cliente
                    Raz_EndCliente = campo(1)
                Else
                    Raz_Cliente = ""                                'Cliente
                    Raz_EndCliente = 0                              'Endereco do Cliente
                End If

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'Unidade de Negocio do Titulo

                Raz_ValorOficial = Replace(txtDeducoes.Text, ".", "")  'Deducoes

                If CDec(txtDeducoesMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtDeducoesMoeda.Text, ".", "")
                Else
                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, txtDeducoes.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                End If

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)
                'Históricos
                Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "C", "D")

                LanctosContabeis(Registro)
            End If


            '**************************************************************************************************
            '*************************************  Juros  ****************************************************
            '**************************************************************************************************

            If CDec(txtJuros.Text) > 0 Then
                'Registro = CInt(txtRegistro.Text)
                Raz_Empresa = Empresa(0)                  'EmpresaCliente
                Raz_EndEmpresa = Empresa(1)               'Endereco Empresa Cliente

                Raz_Conta = Carteira2.CodigoContaJuro   'Conta de Descontos Obtidos

                Dim objPlaContaJuros As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                If objPlaContaJuros.TemCliente Then
                    campo = txtCodigoFornecedor.Value.Split("-")
                    Raz_Cliente = campo(0)                              'Cliente
                    Raz_EndCliente = campo(1)
                Else
                    Raz_Cliente = ""                                'Cliente
                    Raz_EndCliente = 0                              'Endereco do Cliente
                End If

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'Unidade de Negocio do Titulo

                Raz_ValorOficial = Replace(txtJuros.Text, ".", "")  'Juros

                If CDec(txtJurosMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtJurosMoeda.Text, ".", "")
                Else
                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, txtJuros.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                End If

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)
                'Históricos
                Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "D", "C")

                LanctosContabeis(Registro)
            End If


            '**************************************************************************************************
            '*************************************  Acrescimos  ***********************************************
            '**************************************************************************************************

            If CDec(txtAcrescimos.Text) > 0 Then
                'Registro = CInt(txtRegistro.Text)
                Raz_Empresa = Empresa(0)                  'EmpresaCliente
                Raz_EndEmpresa = Empresa(1)               'Endereco Empresa Cliente

                Raz_Conta = Carteira2.CodigoContaAcrescimo    'Conta de Descontos Obtidos

                Dim objPlaContaAcrescimo As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                If objPlaContaAcrescimo.TemCliente Then
                    campo = txtCodigoFornecedor.Value.Split("-")
                    Raz_Cliente = campo(0)                              'Cliente
                    Raz_EndCliente = campo(1)
                Else
                    Raz_Cliente = ""                                'Cliente
                    Raz_EndCliente = 0                              'Endereco do Cliente
                End If

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'Unidade de Negocio do Titulo

                Raz_ValorOficial = Replace(txtAcrescimos.Text, ".", "")  'Acrescimos

                If CDec(txtAcrescimosMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtAcrescimosMoeda.Text, ".", "")
                Else
                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, txtAcrescimos.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                End If

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)
                'Históricos
                Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "D", "C")

                LanctosContabeis(Registro)
            End If

        End If

        '************************************************************************************************************************
        '****************************  LIQUIDO DO DOCUMENTO *********************************************************************
        '************************************************************************************************************************
        If DdlProvisoes.SelectedValue = 1 And (lblAgrupar.Text = "" Or lblAgrupar.Text = "AP") Then
            If (ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso Not ObjTitulo.ObjContaContabilPagadora.Adiantamento) OrElse
                (Left(ObjTitulo.ContaContabilPagadora, 1) = IIf(HTabela.Value = "ContasAPagar", "2", "1")) Then
                'Registro = CInt(txtRegistro.Text)
                Raz_Empresa = EmpresaPagadora(0)                  'Empresa Pagadora
                Raz_EndEmpresa = EmpresaPagadora(1)               'Endereco Empresa Pagadora

                Dim conta() As String = DdlContaPagadora.SelectedValue.Split("-")
                Raz_Conta = conta(5)                    'Conta Contabil


                Dim objPlaConta As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                If objPlaConta.TemCliente Then
                    Raz_Cliente = Cliente(0)                              'Cliente
                    Raz_EndCliente = Cliente(1)
                Else
                    Raz_Cliente = ""                                'Cliente
                    Raz_EndCliente = 0                              'Endereco do Cliente
                End If

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'Pegar Unidadde de Negocio da EmpresaPagadora

                Raz_ValorOficial = Replace(txtValorCobrado.Text, ".", "") 'Valor Liquido

                If CDec(txtValorCobradoMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtValorCobradoMoeda.Text, ".", "")
                Else
                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, txtValorCobrado.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                End If

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)

                Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "C", "D")
                LanctosContabeis(Registro)
            Else
                ContabilizaBaixaAdiantamento(ObjTitulo.ContaContabilPagadora, Registro)
            End If

            '-------------------------------------------
            'Transferencias Financeiras
            '-------------------------------------------
            Sql = " SELECT EmpresaDebito, EnderecoDebito, EmpresaCredito, EnderecoCredito, EmpresaContabil, EnderecoContabil, " & vbCrLf &
                  "        ContaContabil, ClienteContabil,EndClienteContabil, " & vbCrLf &
                  IIf(HTabela.Value = "ContasAPagar", "DebitoCredito", "case when DebitoCredito='D'then 'C' else 'D' end DebitoCredito ") & vbCrLf &
                  "   FROM TransferenciasFinanceiras " & vbCrLf &
                  "  WHERE EmpresaDebito   ='" & Empresa(0) & "'" & vbCrLf &
                  "    and EnderecoDebito  = " & Empresa(1) & vbCrLf &
                  "    and EmpresaCredito  ='" & EmpresaPagadora(0) & "'" & vbCrLf &
                  "    and EnderecoCredito = " & EmpresaPagadora(1)

            For Each DrT As DataRow In Banco.ConsultaDataSet(Sql, "Transferencias").Tables(0).Rows
                Raz_Empresa = DrT("EmpresaContabil")                'EmpresaCliente
                Raz_EndEmpresa = DrT("EnderecoContabil")            'Endereco Empresa Cliente
                Raz_Conta = DrT("ContaContabil")                    'Grupo de Contas
                Raz_Cliente = DrT("ClienteContabil")                'Cliente
                Raz_EndCliente = DrT("EndClienteContabil")          'Endereco do Cliente
                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue
                Raz_ValorOficial = CDec(txtValorDoDocumento.Text) + CDec(txtJuros.Text) + CDec(txtAcrescimos.Text) - CDec(txtDescontos.Text) - CDec(txtDeducoes.Text)

                If ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = CDec(txtValorEmMoeda.Text) + CDec(txtJurosMoeda.Text) + CDec(txtAcrescimosMoeda.Text) - CDec(txtDescontosMoeda.Text) - CDec(txtDeducoesMoeda.Text)
                Else
                    Valor = Replace(DolarizaBaixa(DtMovimento, Raz_ValorOficial, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                    Raz_ValorMoeda = Valor                                      'ValorDoDocumento
                End If

                Raz_ValorMoeda = Replace(Raz_ValorMoeda, ".", "")
                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)
                'Historico
                Raz_DebitoCredito = DrT("DebitoCredito")            'Debito/Credito
                LanctosContabeis(Registro)
            Next
        End If

        If Not Session("objDestinoContabil" & HID.Value) Is Nothing AndAlso DdlProvisoes.SelectedValue = 1 Then
            Dim objDestinoContabil() As String = Session("objDestinoContabil" & HID.Value).ToString.Split("-")
            Sql = " INSERT INTO TitulosXDesdobrarFornecedor (Registro_Id, Cliente, EndCliente, Pedido, Carteira) " & vbCrLf &
                  " VALUES (" & Registro & ",'" & objDestinoContabil(0) & "'," & objDestinoContabil(1) & "," & 0 & ",'" & objDestinoContabil(2) & "')" & vbCrLf
            SqlArray.Add(Sql)
        End If

        '-------------------------------------------
        'Grava Documentos
        '-------------------------------------------
        If ObjTitulo.FinanceiroXDocumentos.Count > 0 Then
            For Each fXd In ObjTitulo.FinanceiroXDocumentos
                fXd.IUD = "I"
            Next

            ObjTitulo.FinanceiroXDocumentos.SalvarSql(SqlArray)
        End If

        If Banco.GravaBanco(SqlArray) = False Then
            MsgBox(Me.Page, Session("ssMessage"))
        Else
            If Not GravarPamcard(Registro) Then
                SqlArray.Clear()

                Sql = " DELETE Razao " & vbCrLf &
                      "  WHERE Titulo = " & Registro
                SqlArray.Add(Sql)

                Sql = "UPDATE " & HTabela.Value & " SET " & vbCrLf &
                      "   Provisao = 2 " & vbCrLf &
                      " WHERE Registro_Id = " & Registro
                SqlArray.Add(Sql)

                If Not Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, Session("ssMessage"))
                    Exit Sub
                End If

                MsgBox(Me.Page, Mensagem)
                Exit Sub
            Else
                GridConsultaTitulos.DataSource = Nothing
                GridConsultaTitulos.DataBind()
                Mensagem = "Registro <" & Registro & "> salvo/atualizado com sucesso! " & MensagemParcial
                MsgBox(Me.Page, Mensagem)
                txtRegistro.Text = Registro
                If Not String.IsNullOrWhiteSpace(Registro) AndAlso lblAgrupar.Text = "AP" Then
                    viewAgrupamento(Registro)
                End If
                If chkEmitirRecibo.Checked Then
                    EmitirRecibo()
                End If
                Limpar(True)
            End If
        End If

    End Sub

    Public Sub ContabilizaBaixaAdiantamento(pContaAdiantamento As String, pSequenciaTitulo As Integer)
        SessaoRecuperaTitulo()
        Dim pEmpresa() As String = DdlEmpresaCliente.SelectedValue.Split("-")
        Dim Empresa As New Cliente(pEmpresa(0), pEmpresa(1))
        Dim pCliente() As String = txtCodigoFornecedor.Value.Split("-")
        Dim objPlaContaAdiantamento As New [Lib].Negocio.PlanoDeConta("", 0, pContaAdiantamento)
        Dim pSeq As Integer

        Dim varAdtos As Boolean = False
        If ObjTitulo.Adiantamentos.Count > 1 Then
            varAdtos = True
            Dim ds As DataSet
            ds = Banco.ConsultaDataSet("select isnull(max(sequencia_id),0) from razao where Lote_Id = 70 and Movimento_Id  = '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'  and Empresa_Id ='" & Empresa.Codigo & "' and EndEmpresa_Id = " & Empresa.CodigoEndereco, "Sequencia")
            pSeq = ds.Tables(0).Rows(0)(0)
        Else
            pSeq = pSequenciaTitulo
        End If

        For Each row In ObjTitulo.Adiantamentos.Where(Function(s) s.VlrABaixarMoeda + s.VlrABaixarOficial > 0)
            '**************************************************
            '** Contabilizacao da conta de Adiantamento *******
            '**************************************************
            Registro = CInt(txtRegistro.Text)
            Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'Unidade de Negocio do Titulo
            Raz_Empresa = Empresa.Codigo                'EmpresaCliente
            Raz_EndEmpresa = Empresa.CodigoEndereco     'Endereco Empresa Cliente
            Raz_Conta = pContaAdiantamento              'Conta de Adiantamento

            If objPlaContaAdiantamento.TemCliente Then
                Raz_Cliente = pCliente(0)               'Cliente
                Raz_EndCliente = pCliente(1)            'Endereco do Cliente
            Else
                Raz_Cliente = ""                        'Cliente
                Raz_EndCliente = 0                      'Endereco do Cliente
            End If

            Raz_ValorOficial = Str(row.VlrABaixarOficial) 'Valor Oficial
            Raz_ValorMoeda = Str(row.VlrABaixarMoeda)     'Valor Moeda
            'Históricos
            Raz_Historico = Funcoes.EliminarCaracteresEspeciais("Baixa do Adiantamento/Titulo " & row.Codigo & "/" & row.CodigoTitulo & " gerado pelo titulo " & Registro & IIf(row.RegistroPedido = 0, ".", ". Referente ao pedido" & row.RegistroPedido) & " / " & row.CodigoCliente & "-" & row.EndCliente & " " & row.Cliente.Nome & IIf(row.RegistroPedido <> CInt(txtPedido.Text) And row.RegistroPedido > 0 And CInt(txtPedido.Text) > 0, " / Compensado pelo Pedido " & txtPedido.Text, ""))
            Raz_DebitoCredito = IIf(Left(pContaAdiantamento, 1) = "1", "C", "D")

            If varAdtos Then pSeq += 1

            LanctosContabeis(pSeq)

            '**************************************************
            '****** Contabilizacao da Conta de Variacao *******
            '**************************************************
            If row.VlrCalcVariacao <> 0 Then
                Registro = CInt(txtRegistro.Text)
                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'Unidade de Negocio do Titulo
                Raz_Empresa = Empresa.Codigo                                           'EmpresaCliente
                Raz_EndEmpresa = Empresa.CodigoEndereco                                'Endereco Empresa Cliente

                If Left(pContaAdiantamento, 1) = "1" Then
                    Raz_Conta = IIf(row.VlrCalcVariacao > 0, Empresa.Empresa.CodigoContaVariacaoMonetariaAtiva, Empresa.Empresa.CodigoContaVariacaoMonetariaPassiva)
                    Raz_DebitoCredito = IIf(row.VlrCalcVariacao > 0, "C", "D")
                Else
                    Raz_Conta = IIf(row.VlrCalcVariacao > 0, Empresa.Empresa.CodigoContaVariacaoMonetariaPassiva, Empresa.Empresa.CodigoContaVariacaoMonetariaAtiva)
                    Raz_DebitoCredito = IIf(row.VlrCalcVariacao > 0, "D", "C")
                End If

                Raz_Cliente = ""                        'Cliente
                Raz_EndCliente = 0                      'Endereco do Cliente

                Raz_ValorOficial = Str(Math.Abs(row.VlrCalcVariacao)) 'Valor Variacao
                Raz_ValorMoeda = Str(0)                 'Valor Variacao
                'Históricos
                Raz_Historico = Funcoes.EliminarCaracteresEspeciais("Variacao Baixa do Adiantamento/Titulo " & row.Codigo & "/" & row.CodigoTitulo & " gerado pelo titulo " & Registro & IIf(row.RegistroPedido = 0, ".", ". Referente ao pedido" & row.RegistroPedido) & " / " & row.CodigoCliente & "-" & row.EndCliente & " " & row.Cliente.Nome & IIf(row.RegistroPedido <> CInt(txtPedido.Text) And row.RegistroPedido > 0 And CInt(txtPedido.Text) > 0, " / Compensado pelo Pedido " & txtPedido.Text, ""))
                'pSeq += 1
                LanctosContabeis(pSeq)
            End If
        Next
    End Sub

    Function GravarPamcard(ByVal titulo As String) As Boolean
        If txtTipoDoDocumento.Value = 4 AndAlso DdlProvisoes.SelectedValue = 1 Then
            Dim sql As String = ""
            sql = "SELECT NFxT.Empresa_Id, NFxT.EndEmpresa_Id, NFxT.Cliente_Id, " & vbCrLf &
                  "       NFxT.EndCliente_Id, NFxT.EntradaSaida_Id, NFxT.Serie_Id, " & vbCrLf &
                  "       NFxT.Nota_Id " & vbCrLf &
                  "  FROM  " & HTabela.Value & " CP " & vbCrLf &
                  " INNER JOIN NotaFiscalXTitulo NFxT " & vbCrLf &
                  "    ON CP.Registro_Id = NFxT.Titulo_Id " & vbCrLf &
                  " WHERE (CP.Registro_Id = " & titulo & ")  " & vbCrLf

            Dim ds As New DataSet
            ds = Banco.ConsultaDataSet(sql, "Nota Fiscal")

            If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                Mensagem = "Contrato de Frete não foi localizado, verifique."
                Return False
            Else
                Dim objNotaFiscal = New [Lib].Negocio.NotaFiscal()
                objNotaFiscal.CodigoEmpresa = ds.Tables(0).Rows(0).Item("Empresa_Id")
                objNotaFiscal.EnderecoEmpresa = ds.Tables(0).Rows(0).Item("EndEmpresa_Id")
                objNotaFiscal.CodigoCliente = ds.Tables(0).Rows(0).Item("Cliente_Id")
                objNotaFiscal.EnderecoCliente = ds.Tables(0).Rows(0).Item("EndCliente_Id")
                objNotaFiscal.Serie = ds.Tables(0).Rows(0).Item("Serie_Id")
                objNotaFiscal.Codigo = ds.Tables(0).Rows(0).Item("Nota_Id")
                objNotaFiscal.EntradaSaida = IIf(ds.Tables(0).Rows(0).Item("EntradaSaida_Id") = "S", eEntradaSaida.Saida, eEntradaSaida.Entrada)
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNotaFiscal)

                'Baixar Contrato Pamcard
                Dim x As New Pamcard()
                ds = x.AtualizarValoresSaldoContratoDeFrete("04854422000185", objNotaFiscal, txtProrrogacao.Text, txtDescontos.Text, txtDeducoes.Text, txtJuros.Text, txtAcrescimos.Text, txtValorCobrado.Text)

                If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                    Mensagem = "Erro inesperado gravando Contrato: " & Funcoes.EliminarCaracteresEspeciais(RTrim(Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI"
                    Return False
                Else
                    If ds.Tables(0).Rows(0).Item("erro") = 0 Then
                        Return True
                    Else
                        Mensagem = "Erro: Código " & ds.Tables(0).Rows(0).Item("erro") & " - " & ds.Tables(0).Rows(0).Item("errodescricao")
                        Return False
                    End If
                End If
            End If
        Else
            Return True
        End If
    End Function

    Private Sub GravaTituloParcial(ByRef MensagemParcial As String, ByVal TituloOrigem As Integer, ByRef objpedido As Pedido)
        '******************************************************
        '*******  Gera sequencia de titulos Parcial  **********
        '******************************************************
        Dim NovaSeq As Integer = 0
        Sql = "Exec sp_Numerador '" & Session("ssNomeServidor").ToUpper() & "',0,1"
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Numerador").Tables(0).Rows
            NovaSeq = Dr("Sequencia")
        Next

        Dim VlrOficial As Decimal = CDec(HDValorOriginalOficial.Value) - CDec(txtValorDoDocumento.Text)
        Dim VlrMoeda As Decimal = CDec(HDValorOriginalMoeda.Value) - CDec(txtValorEmMoeda.Text)
        VlrOficial = IIf(VlrOficial <= 0, 0, VlrOficial)
        VlrMoeda = IIf(VlrMoeda <= 0, 0, VlrMoeda)

        If ddlMoeda.SelectedValue = 3 Then
            If objpedido IsNot Nothing AndAlso objpedido.Codigo > 0 AndAlso objpedido.IndexadorFixo Then
                VlrOficial = Math.Round(VlrMoeda * objpedido.IndiceFixado, 2)
            Else
                VlrOficial = Funcoes.ConverteParaMoedaOficial(VlrMoeda, 3, Now.Date, True, False, 2)
            End If
        End If

        Sql = "INSERT INTO " & HTabela.Value & vbCrLf &
              " (Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa," & vbCrLf &
              "  UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente," & vbCrLf &
              "  ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora," & vbCrLf &
              "  ContaContabilPagadora, cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao," & vbCrLf &
              "  ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros," & vbCrLf &
              "  MoedaAcrescimos, MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, solicitacao," & vbCrLf &
              "  UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao," & vbCrLf &
              "  UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, RegistroMestre, Observacoes, SituacaoBancaria," & vbCrLf &
              "  UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate," & vbCrLf &
              "  UsuarioLiberacaoCheque, UsuarioLiberacaoChequeDate, CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento)" & vbCrLf &
              "SELECT " & NovaSeq & ", Sequencia_Id, 2, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa," & vbCrLf &
              "       UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente," & vbCrLf &
              "       ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora," & vbCrLf &
              "       ContaContabilPagadora, cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao," & vbCrLf &
              "       " & Str(VlrOficial) & ", 0, 0, 0, 0, " & Str(VlrOficial) & "," & vbCrLf &
              "       " & Str(VlrMoeda) & ", 0, 0, 0, 0, " & Str(VlrMoeda) & "," & vbCrLf &
              "       Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, solicitacao," & vbCrLf &
              "       '" & Session("ssNomeUsuario") & "', '" & CDate(Today).ToString("yyyy/MM/dd") & "', '', NULL, '', NULL, ''," & vbCrLf &
              "       NULL, '', NULL, Grupado, RegistroMestre, Observacoes, SituacaoBancaria," & vbCrLf &
              "       UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate," & vbCrLf &
              "       '', NULL, CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento" & vbCrLf &
              "  FROM " & HTabela.Value & vbCrLf &
              " WHERE Registro_Id = " & TituloOrigem

        SqlArray.Add(Sql)

        Sql = " INSERT INTO NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id) " & vbCrLf &
              " Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, " & NovaSeq & vbCrLf &
              "   From NotaFiscalXTitulo" & vbCrLf &
              "  Where Titulo_Id = " & TituloOrigem
        SqlArray.Add(Sql)

        Sql = " INSERT INTO FaturaDeFreteXTitulo(Empresa_Id, EndEmpresa_Id, Conveniado_Id, EndConveniado_Id, Fatura_Id, Titulo_Id) " & vbCrLf &
              " Select Empresa_Id, EndEmpresa_Id, Conveniado_Id, EndConveniado_Id, Fatura_Id, " & NovaSeq & vbCrLf &
              "   From FaturaDeFreteXTitulo" & vbCrLf &
              "  Where Titulo_Id = " & TituloOrigem
        SqlArray.Add(Sql)

        MensagemParcial = " E Registro Parcial de Número <" & NovaSeq & ">..."
    End Sub

    Private Sub Limpar(ByVal LimparConsulta As Boolean)
        Try
            BloquearValores(False)

            Session.Remove("objBancoFRMTIT" & HID.Value)
            Session.Remove("objDestinoContabil" & HID.Value)
            Session.Remove("objPedidoSelecionadoCTAPAG" & HID.Value)
            Session.Remove("ssObservacoes" & HID.Value)
            Session.Remove("ssNomeLiberacao" & HID.Value)
            Session.Remove("ssNomeLiberacaoData" & HID.Value)
            Session.Remove("ssRegistros" & HID.Value)
            Session.Remove("ssGrupado" & HID.Value)
            Session.Remove("objFornecedorCP" & HID.Value)
            Session.Remove("objFavorecidoCP" & HID.Value)
            Session.Remove("objTitulo" & HID.Value)
            Session.Remove("ControleCP" & HID.Value)
            Session.Remove("ssRetornoDs" & HID.Value)
            Session.Remove("ssRetorno" & HID.Value)
            Session.Remove("objPedCons" & HID.Value)
            Session.Remove("objRegistrosSelecionados" & HID.Value)
            Session.Remove("objNotaFiscal" & HID.Value)
            Session.Remove("objConsultarNaviosXInvoice" & HID.Value)
            Session.Remove("objCarregarTitulo" & HID.Value)

            gridAdiantamentosDisponiveis.DataSource = Nothing
            gridAdiantamentosDisponiveis.DataBind()
            gridBaixasAdiantamentos.DataSource = Nothing
            gridBaixasAdiantamentos.DataBind()
            divSelecaoAdBX.Visible = False
            divAdBaixados.Visible = False
            divCheque.Visible = False

            divEndosso.Visible = False
            gridEndosso.DataSource = Nothing
            gridEndosso.DataBind()

            idRetornoBancario.Visible = False
            txtObsRetornoBancario.Text = ""

            HID.Value = Guid.NewGuid().ToString
            ucConsultaClientes.SetarHID(HID.Value)
            ucConsultaAdiantamentos.SetarHID(HID.Value)
            ucConsultaPedidos.SetarHID(HID.Value)
            ucDestinoContabil.SetarHID(HID.Value)
            ucConsultaTitulo.SetarHID(HID.Value)

            ucFile.Clear()

            If LimparConsulta AndAlso lblAgrupar.Text = "AP" Then Limpar_ConsultaTitulos(False)

            If chkManterLancamento.Checked Then
                DdlProvisoes.Enabled = True
                txtMovimento.Enabled = True
                txtProrrogacao.Enabled = True

                ObjTitulo = New NGS.Lib.Negocio.TituloV

                If DdlContaPagadora.SelectedValue.Length > 0 Then
                    'Guardar conta da Empresa pagadora
                    Dim dadosBanc As String() = DdlContaPagadora.SelectedValue.Split("-")
                    ObjTitulo.CodigoBancoPagador = DdlBancoPagador.SelectedValue
                    ObjTitulo.CodigoAgenciaPagadora = dadosBanc(0)
                    ObjTitulo.DigitoAgenciaPagadora = dadosBanc(1)
                    ObjTitulo.ContaPagadora = dadosBanc(2)
                    ObjTitulo.DigitoContaPagadora = dadosBanc(3)
                    ObjTitulo.TipoDaContaPagadora = dadosBanc(4)
                    ObjTitulo.ContaContabilPagadora = dadosBanc(5)
                End If

                SessaoSalvaTitulo()
            Else

                divAd.Visible = False

                Carteiras()

                SqlArray.Clear()

                txtTipoDoDocumento.Value = 0

                divArquivos.Visible = False
                rowDolar.Visible = False
                imgBloqueio.Visible = False
                imgExtratoPedido.Visible = False
                txtCodigoFaturaDeFrete.Text = 0
                txtLiberarTitulo.Value = "N"
                txtLiberarPedido.Value = "N"
                txtLiberarCheque.Value = "N"
                txtUsuarioLiberarTitulo.Value = ""
                txtUsuarioLiberarPedido.Value = ""
                txtUsuarioLiberarCheque.Value = ""
                txtUsuarioLiberarTituloData.Value = ""
                txtUsuarioLiberarPedidoData.Value = ""
                txtUsuarioLiberarChequeData.Value = ""
                txtSlip.Value = ""
                txtPedidoFixacao.Value = ""
                txtDataConciliacao.Value = ""
                chkConciliado.Checked = False
                txtMestre.Text = String.Empty

                txtContratoBanco.Text = String.Empty

                ContratoBanco.Visible = Funcoes.VerificaPermissao("ContratoBancario", "GRAVAR")

                txtFornecedor.Text = ""
                txtCodigoFornecedor.Value = ""
                cmdClientesTitulo.Enabled = True

                txtFavorecido.Text = ""
                txtCodigoFavorecido.Value = ""

                lblBanco.Text = "Banco"
                lblAgencia.Text = "Agência"
                lblContaCorrente.Text = "Conta"
                lblTipoConta.Text = "Tipo Da Conta"
                imgContaObs.ToolTip = "Observações"

                DdlTributos.Items.Clear()

                DdlEmpresaPagadora.SelectedIndex = 0
                DdlBancoPagador.Items.Clear()
                DdlContaPagadora.Items.Clear()

                lblCotacao.Text = String.Empty
                lblDescCotacao.Text = String.Empty

                divlCot.Visible = False

                ddlCarteiras.Enabled = True
                'ddlCarteiras.SelectedIndex = 0
                DdlTiposDePagamentos.SelectedIndex = 0
                ddlMoeda.SelectedValue = 1
                ddlIndexador.SelectedIndex = 3
                DdlProvisoes.SelectedIndex = 0

                ddlCarteiraDoTitulo.SelectedIndex = 0
                txtContratoFinanceiro.Text = ""
                ViewState.Clear()
                txtHistorico.Text = ""
                txtHistoricoAutoComplete.Text = String.Empty
                txtHistorico.Enabled = True
                txtCodigoDeBarras.Text = ""

                txtPedido.Text = "0"
                cmdPedido.Enabled = True
                txtCessaoDeCredito.Text = "0"
                txtCessaoDeCredito.Parent.Visible = False

                txtSolicitacao.Text = ""
                CkbCodigoDeBarras.Checked = False

                txtNumeroCheque.Text = ""

                txtVencimentoAdto.Text = ""
                txtTaxaAdto.Text = 0
                ddlSafraAdto.Enabled = True
                ddlSafraAdto.SelectedIndex = 0

                GridPedidos.DataSource = Nothing
                GridPedidos.DataBind()

                gridRazao.DataSource = Nothing
                gridRazao.DataBind()

                txtRealDolar.Value = ""
                HiddenIndexador.Value = ""
                txtObservacoes.Text = ""
                lblAgrupar.Text = ""
                ChkLiberado.Checked = False
                txtObservacoes.Text = ""

                For Each objLinha As GridViewRow In GridConsultaTitulos.Rows
                    Dim chkTitulo As CheckBox = objLinha.Cells(0).FindControl("ChkGridTitulos")
                    If chkTitulo.Checked Then chkTitulo.Checked = False
                Next

                DdlUnidadeDeNegocioEmpresaCliente.Enabled = True
                DdlEmpresaCliente.Enabled = True

                lblDeducoes.Parent.Visible = True
                lblDescontos.Parent.Visible = True
                lblJuros.Parent.Visible = True
                lblAcrescimos.Parent.Visible = True

                txtValorDoDocumento.Enabled = True
                txtDescontos.Enabled = True
                txtDeducoes.Enabled = True
                txtJuros.Enabled = True
                txtAcrescimos.Enabled = True
                txtValorCobrado.Enabled = True

                txtValorEmMoeda.Enabled = True
                txtDescontosMoeda.Enabled = True
                txtDeducoesMoeda.Enabled = True
                txtJurosMoeda.Enabled = True
                txtAcrescimosMoeda.Enabled = True
                txtValorCobradoMoeda.Enabled = True

                txtSolicitacao.Enabled = True
                txtCodigoDeBarras.Enabled = True
                ddlMoeda.Enabled = True
                ddlIndexador.Enabled = True
                ddlCarteiras.Enabled = True
                DdlTributos.Enabled = True
                txtHistorico.Enabled = True
                DdlProvisoes.Enabled = True
                lnkRelatorioAgrupamento.Parent.Visible = False
                pnlReprogramaVencimentos.Visible = False
                lnkReprogramar.Enabled = False
                lblTotalRegistroAgrupado.Text = ""

                txtMovimento.Enabled = True
                txtProrrogacao.Enabled = True

                If Not DdlEmpresaCliente.SelectedIndex > 0 Then VerificaUnidade()
            End If

            imgBoletoPDF.Visible = False
            DdlEmpresaPagadora.Enabled = True
            DdlTiposDePagamentos.Enabled = True
            DdlBancoPagador.Enabled = True
            DdlContaPagadora.Enabled = True
            DdlUnidadeDeNegocioEmpresaCliente.Enabled = False
            DdlEmpresaCliente.Enabled = False

            lnkNovo.Parent.Visible = False
            lnkConsultar.Parent.Visible = True
            lnkExcluir.Parent.Visible = False
            lnkRelatorio.Parent.Visible = False
            lnkSituacaoBancaria.Parent.Visible = False

            ChkLiberado.Parent.Visible = False
            chkConciliado.Parent.Visible = False

            txtRegistro.Text = ""
            txtRegistro.Enabled = False
            txtProrrogacao.Text = ""
            hdnProrrogacaoOriginal.Value = String.Empty
            lblVencOriginal.Text = ""

            txtDataEntradaSistema.Text = CDate(Today).ToString("dd/MM/yyyy")
            DdlProvisoes.SelectedValue = 2

            txtDescontos.Text = ""
            txtDeducoes.Text = ""
            txtJuros.Text = ""
            txtAcrescimos.Text = ""
            txtValorCobrado.Text = ""

            txtDescontosMoeda.Text = ""
            txtDeducoesMoeda.Text = ""
            txtJurosMoeda.Text = ""
            txtAcrescimosMoeda.Text = ""
            txtValorCobradoMoeda.Text = ""

            lblDeducoes.Text = "Deduções:"
            lblDescontos.Text = "Descontos:"
            lblJuros.Text = "Juros:"
            lblAcrescimos.Text = "Acréscimos:"

            Session("ssObservacoes" & HID.Value) = 0
            Session("ssNomeLiberacao" & HID.Value) = ""
            Session("ssNomeLiberacaoData" & HID.Value) = Today
            Session("ssRegistros" & HID.Value) = ""
            Session("ssGrupado" & HID.Value) = ""
            Session("ssNomeUsuario" & HID.Value) = ""

            txtValorDoDocumento.Text = ""
            HDValorOriginalOficial.Value = 0
            HDValorOriginalMoeda.Value = 0
            txtValorEmMoeda.Text = ""
            txtValorLiquido.Value = 0
            txtValorLiquidoMoeda.Value = 0

            txtNumeroAdto.Text = "0"

            txtPedidoConsultaTitulos.Text = ""

            ddlUsuarios.Items.Clear()

            ImgValores.Enabled = True

            TabPanel4.Visible = True

            txtValorDoDocumento.ForeColor = System.Drawing.Color.Black
            txtValorEmMoeda.ForeColor = System.Drawing.Color.Black
            txtOficial.ForeColor = System.Drawing.Color.Black
            txtMoeda.ForeColor = System.Drawing.Color.Black
            cmdPedido.Enabled = True

            'objPedido = New [Lib].Negocio.Pedido(objEmpresa(0), objEmpresa(1), "")
            'objPedido.Empresa = New [Lib].Negocio.Cliente(objEmpresa(0), objEmpresa(1))

            Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresaCliente.SelectedValue.ToString().Split("-")(0), DdlEmpresaCliente.SelectedValue.ToString().Split("-")(1))

            If objEmpresa.Empresa.ObrigaNavio Then
                divNaviosXInvoice.Visible = True
            Else
                divNaviosXInvoice.Visible = False
            End If
            txtNaviosXInvoice.Text = String.Empty

            ucConsultarNaviosXInvoice.SetarHID(HID.Value)

            '*************************************************************************
            '**************************** Tab Contratos ******************************
            '*************************************************************************
            txtDescricaoDocumento.Text = String.Empty
            txtNomeDoArquivo.Text = String.Empty
            gridDocumentos.DataSource = Nothing
            gridDocumentos.DataBind()

            LiberaEmpresa()

        Catch ex As Exception
            MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(ex.Message))
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeDeNegocioEmpresaCliente.Enabled = False
            DdlEmpresaCliente.Enabled = False

            DdlUnidadeConsultaTitulos.Enabled = False
            DdlEmpresaConsultaTitulos.Enabled = False
        End If
    End Sub

    Protected Sub GridConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Registro = GridConsultaTitulos.SelectedRow.Cells(3).Text()
            Limpar(False)
            txtRegistro.Text = Registro
            ConsultaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message))
        End Try
    End Sub

    Public Function ConsultaCLientes(ByVal Cli As String, ByVal EndCli As Integer) As String
        Dim Cliente As String = ""
        Dim objCliente As New [Lib].Negocio.Cliente(Cli, EndCli)

        txtCodigoSituacaoCliente.Value = objCliente.CodigoSituacao
        txtSituacaoCliente.Value = objCliente.Situacao.Descricao
        Cliente = objCliente.Nome & " - " & objCliente.Cidade & " - " & objCliente.CodigoEstado & "  " & Funcoes.FormatarCpfCnpj(objCliente.Codigo) & " - " & objCliente.CodigoEndereco & " - " & objCliente.Situacao.Descricao

        Return Cliente
    End Function

    Private Sub ConsultaTitulo()
        Dim Pedido As Integer = 0
        Dim SaldoParcelas As Decimal
        Dim conta As String = ""
        Dim ConferenciaNF As Boolean = False
        TemRegistro = ""

        Try
            If txtRegistro.Text <> "" Then
                Registro = txtRegistro.Text
                txtRegistro.Text = Registro

                Sql = " SELECT ISNULL(NF.Nota_Id,0) AS Nota, CP.Registro_Id, CP.Sequencia_Id, CP.Provisao, CP.Carteira, CP.Tributo, CP.Indexador, " & vbCrLf &
                      "        CP.Moeda, CP.TipoPagto, CP.Situacao, CP.Lote, CP.Movimento, CP.Vencimento, CP.Prorrogacao, " & vbCrLf &
                      "        CP.DataMoeda, isnull(CP.Baixa,CP.Prorrogacao) as Baixa, CP.UnidadeDeNegocio, CP.Empresa, CP.EndEmpresa, CP.Cliente, CP.EndCliente, CP.ContaContabilCliente," & vbCrLf &
                      "        CP.BancoCliente, CP.AgenciaCliente, CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente, Isnull(CP.TipoContaCliente,'C') AS TipoContaCliente," & vbCrLf &
                      "        CP.EmpresaPagadora, CP.EndEmpresaPagadora, CP.BancoPagador, isnull(CP.AgenciaPagadora,'') AS AgenciaPagadora, CP.DigitoAgenciaPagadora, " & vbCrLf &
                      "        CP.ContaPagadora, CP.DigitoContaPagadora, Isnull(CP.TipoContaPagadora,'C') AS TipoContaPagadora, CP.ContaContabilPagadora, CP.Cheque, isnull(CP.Slips,'N') AS Slips, " & vbCrLf &
                      "        CP.Recibo, CP.Aviso, CP.ReciboDeposito, isnull(CP.EmpresaPedido,'') AS EmpresaPedido, isnull(CP.EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(CP.Pedido, 0) AS Pedido, " & vbCrLf &
                      "        isnull(CP.PedidoFixacao,0) AS PedidoFixacao, isnull(CP.Procuracao,0) AS Procuracao, CP.ValorDoDocumento, CP.Descontos, CP.Deducoes, " & vbCrLf &
                      "        CP.Juros, CP.Acrescimos, CP.ValorLiquido, ISNULL(CP.MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, " & vbCrLf &
                      "        ISNULL(CP.MoedaDescontos, 0) AS MoedaDescontos, ISNULL(CP.MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(CP.MoedaJuros, 0) AS MoedaJuros, " & vbCrLf &
                      "        ISNULL(CP.MoedaAcrescimos, 0) AS MoedaAcrescimos, ISNULL(CP.MoedaValorLiquido, 0) AS MoedaValorLiquido, CP.Historico, " & vbCrLf &
                      "        CP.CodigoDeBarras, CP.CodigoDigitado, isnull(CP.CodigoDeBarraPreImpresso,0) AS CodigoDeBarraPreImpresso, CP.Destinatario, CP.EndDestinatario, CP.NomeDoDestinatario, CP.Destinacao, " & vbCrLf &
                      "        CP.Solicitacao, CP.UsuarioInclusao, CP.UsuarioInclusaoData, CP.UsuarioAlteracao, CP.UsuarioAlteracaoData, " & vbCrLf &
                      "        CP.UsuarioCancelamento, CP.UsuarioCancelamentoData, isnull(CP.UsuarioLiberacao,'') AS UsuarioLiberacao, CP.UsuarioLiberacaoData, " & vbCrLf &
                      "        CP.UsuarioBaixa, CP.UsuarioBaixaData, isnull(CP.Grupado,'N') AS Grupado, isnull(CP.RegistroMestre, 0) as RegistroMestre, CP.Observacoes, " & vbCrLf &
                      "        CP.SituacaoBancaria, ISNULL(CP.NumeroDoCheque,0) AS NumeroDoCheque, isnull(ad.adiantamento_id,isnull(CP.Adiantamento,0)) as Adiantamento, CP.VencimentoAdto, ISNULL(CP.TaxaAdto,0) TaxaAdto, " & vbCrLf &
                      "        isnull(CP.UsuarioLiberacaoBloqueio, '') AS UsuarioLiberacaoBloqueio, CP.UsuarioLiberacaoBloqueioDate, isnull(CP.UsuarioLiberacaoPedido, '') AS UsuarioLiberacaoPedido, " & vbCrLf &
                      "        CP.UsuarioLiberacaoPedidoDate, isnull(CP.UsuarioLiberacaoCheque, '') AS UsuarioLiberacaoCheque, CP.UsuarioLiberacaoChequeDate, " & vbCrLf &
                      "        isnull(CP.CarteiraAdto,'') AS CarteiraAdto, isnull(CP.CarteiraDoTitulo,0) AS CarteiraDoTitulo, isnull(CP.ContratoDeFinanciamento,'') AS ContratoDeFinanciamento, " & vbCrLf &
                      "        ISNULL(NF.TipoDeDocumento,0) AS TipoDeDocumento, isnull(ContratoANTT,'') AS ContratoANTT, isnull(Ped.FinanceiroAberto,1) AS FinanceiroAberto, isnull(Ped.MomentoFinanceiro,0) AS MomentoFinanceiro, " & vbCrLf &
                      "        ISNULL(NF.NFG,0) AS NFG, ISNULL(NF.Conferencia, 1) AS Conferencia, " & vbCrLf &
                      "        ISNULL(FFxT.Fatura_Id, 0) AS FaturaDeFrete, ISNULL(CP.ContratoBancario, '') AS ContratoBanco, " & vbCrLf &
                      "        CASE " & vbCrLf &
                      "			   WHEN ISNULL(BaixaAdiantamento.ValorOficial,0) = 0 " & vbCrLf &
                      "				    THEN 'S' " & vbCrLf &
                      "				    ELSE 'N' " & vbCrLf &
                      "		   END AS LiberaAdiantamento," & vbCrLf &
                      "        isnull(ped.safra,isnull(ad.safra,'NENHUMA')) as Safra, ISNULL(CP.BoletoBancario, 0) AS BoletoBancario, ISNULL(CP.HistoricoRemessa, '') AS HistoricoRemessa, ISNULL(eTxT.Codigo_Id, 0) AS SequenciaEndosso, " & vbCrLf &
                      "        ISNULL(CP.FolhaDePagamento, 0) AS FolhaDePagamento" & vbCrLf &
                      "   FROM " & HTabela.Value & " CP" & vbCrLf &
                      "	  LEFT JOIN NotaFiscalXTitulo NFxT " & vbCrLf &
                      "     ON CP.Registro_Id = NFxT.Titulo_Id " & vbCrLf &
                      "     AND CP.Empresa       = NFxT.Empresa_Id " & vbCrLf &
                      "     AND CP.EndEmpresa   = NFxT.EndEmpresa_Id " & vbCrLf &
                      "	  LEFT JOIN NotasFiscais NF " & vbCrLf &
                      "     ON NFxT.Empresa_Id       = NF.Empresa_Id " & vbCrLf &
                      "	   AND NFxT.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf &
                      "    AND NFxT.Cliente_Id      = NF.Cliente_Id " & vbCrLf &
                      "    AND NFxT.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf &
                      "    AND NFxT.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
                      "    AND NFxT.Serie_Id        = NF.Serie_Id " & vbCrLf &
                      "    AND NFxT.Nota_Id         = NF.Nota_Id" & vbCrLf &
                      "   LEFT JOIN FaturaDeFreteXTitulo FFxT " & vbCrLf &
                      "     ON CP.Registro_Id = FFxT.Titulo_Id" & vbCrLf &
                      "   LEFT JOIN Pedidos Ped" & vbCrLf &
                      "     ON Ped.Empresa_Id    = CP.EmpresaPedido " & vbCrLf &
                      "    AND Ped.EndEmpresa_Id = CP.EndEmpresaPedido " & vbCrLf &
                      "    AND Ped.Pedido_Id     = CP.Pedido" & vbCrLf &
                      "   Left Join VW_Adiantamento ad" & vbCrLf &
                      "     on ad.Titulo = CP.Registro_id" & vbCrLf &
                      "   LEFT JOIN (SELECT aXb.Empresa_id, aXb.EndEmpresa_Id, aXb.RegistroPedido, sum(aXb.ValorOficial) AS ValorOficial" & vbCrLf &
                      "	               FROM vw_AdiantamentosXBaixas aXb" & vbCrLf &
                      "				  GROUP BY aXb.Empresa_id, aXb.EndEmpresa_Id, aXb.RegistroPedido) BaixaAdiantamento" & vbCrLf &
                      "	    ON BaixaAdiantamento.Empresa_id      = CP.Empresa" & vbCrLf &
                      "	   AND BaixaAdiantamento.EndEmpresa_id  = CP.EndEmpresa" & vbCrLf &
                      "	   AND BaixaAdiantamento.RegistroPedido = CP.Pedido" & vbCrLf &
                      "   LEFT JOIN EndossoXTitulo eTxT" & vbCrLf &
                      "     ON eTxT.Empresa_Id    = CP.Empresa " & vbCrLf &
                      "    AND eTxT.EndEmpresa_Id = CP.EndEmpresa " & vbCrLf &
                      "    AND eTxT.Titulo_Id     = CP.Registro_Id" & vbCrLf &
                      "  WHERE CP.Registro_Id = " & Registro

                Dim dsTitulo As New DataSet
                dsTitulo = Banco.ConsultaDataSet(Sql, "Titulo")

                If dsTitulo Is Nothing OrElse dsTitulo.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Registro não encontrado")
                    Exit Sub
                End If

                Dim Dr As DataRow = dsTitulo.Tables(0).Rows(0)

                SessaoRecuperaTitulo()

                'TEM QUE TER PERMISSÃO PARA ACESSAR TÍTULO PAGAMENTO DE FUNCIONÁRIO - FURLAN 05/08/2025
                'BLOQUEADO PELA INTEGRAÇÃO E ENVIO DE ARQUIVO PARA BANCO
                If Dr("Situacao") = 105 AndAlso Not Funcoes.VerificaPermissao("TituloDeFuncionario", "LEITURA") Then
                    MsgBox(Me.Page, "Você não ter permissão para consultar esse Registro.", eTitulo.Info)
                    Limpar(True)
                    Exit Sub
                End If

                ObjTitulo.Codigo = Dr("Registro_id")

                Session("ControleCP" & HID.Value) = Dr("Registro_Id").ToString + ";" + Dr("ValorDoDocumento").ToString + ";" + Dr("MoedaValorDoDocumento").ToString + ";" + Dr("Moeda").ToString

                DdlUnidadeDeNegocioEmpresaCliente.SelectedIndex = DdlUnidadeDeNegocioEmpresaCliente.Items.IndexOf(DdlUnidadeDeNegocioEmpresaCliente.Items.FindByValue(Dr("UnidadeDeNegocio")))
                CargaEmpresaCliente()
                DdlEmpresaCliente.SelectedIndex = DdlEmpresaCliente.Items.IndexOf(DdlEmpresaCliente.Items.FindByValue(Dr("Empresa") & "-" & CStr(Dr("EndEmpresa"))))

                ObjTitulo.CodigoEmpresa = Dr("Empresa")
                ObjTitulo.EnderecoEmpresa = Dr("EndEmpresa")

                ObjTitulo.CodigoEmpresaPedido = Dr("EmpresaPedido")
                ObjTitulo.EndEmpresaPedido = Dr("EndEmpresaPedido")

                Dim strCliente() As String = ConsultaCLientes(Dr("Cliente"), Dr("EndCliente")).Split("-")

                txtFornecedor.Text = ConsultaCLientes(Dr("Cliente"), Dr("EndCliente")) '
                txtCodigoFornecedor.Value = Dr("Cliente") & "-" & CStr(Dr("EndCliente"))

                ObjTitulo.CodigoCliente = Dr("Cliente")
                ObjTitulo.EndCliente = Dr("EndCliente")

                If Dr("Destinatario") = "" Then
                    txtCodigoFavorecido.Value = Dr("Cliente") & "-" & CStr(Dr("EndCliente"))
                    txtFavorecido.Text = ConsultaCLientes(Dr("Cliente"), Dr("EndCliente"))
                Else
                    txtCodigoFavorecido.Value = Dr("Destinatario") & "-" & Dr("EndDestinatario")
                    txtFavorecido.Text = ConsultaCLientes(Dr("Destinatario"), Dr("EndDestinatario"))
                End If

                ObjTitulo.CodigoBancoCliente = Dr("BancoCliente")
                ObjTitulo.CodigoAgenciaCliente = Dr("AgenciaCliente")
                ObjTitulo.DigitoAgenciaCliente = Dr("DigitoAgenciaCliente")
                ObjTitulo.ContaCliente = Dr("ContaCliente")
                ObjTitulo.DigitoContaCliente = Dr("DigitoContaCliente")
                ObjTitulo.TipoDaContaCliente = Dr("TipoContaCliente")
                ObjTitulo.ContaContabilCliente = Dr("ContaContabilCliente")
                AtualizaDadosBancariosNoForm(ObjTitulo)

                DdlEmpresaPagadora.SelectedIndex = DdlEmpresaPagadora.Items.IndexOf(DdlEmpresaPagadora.Items.FindByValue(Dr("EmpresaPagadora") & "-" & CStr(Dr("EndEmpresaPagadora"))))
                DdlTiposDePagamentos.SelectedIndex = DdlTiposDePagamentos.Items.IndexOf(DdlTiposDePagamentos.Items.FindByValue(Dr("TipoPagto")))

                If DdlEmpresaPagadora.Text <> "" Then
                    CargaBancoPagador()
                End If

                If Dr("BancoPagador") > 0 AndAlso Dr("AgenciaPagadora").ToString.Length > 0 Then
                    ObjTitulo.CodigoBancoPagador = Dr("BancoPagador")
                    ObjTitulo.CodigoAgenciaPagadora = Dr("AgenciaPagadora")
                    ObjTitulo.DigitoAgenciaPagadora = Dr("DigitoAgenciaPagadora")
                    ObjTitulo.ContaPagadora = Dr("ContaPagadora")
                    ObjTitulo.DigitoContaPagadora = Dr("DigitoContaPagadora")
                    ObjTitulo.TipoDaContaPagadora = Dr("TipoContaPagadora")
                    ObjTitulo.ContaContabilPagadora = Dr("ContaContabilPagadora")

                    DdlBancoPagador.SelectedIndex = DdlBancoPagador.Items.IndexOf(DdlBancoPagador.Items.FindByValue(Dr("BancoPagador")))
                    BancoPagador()
                    conta = Dr("AgenciaPagadora") & "-" & Dr("DigitoAgenciaPagadora") & "-" & Dr("ContaPagadora") & "-" & Dr("DigitoContaPagadora") & "-" & Dr("TipoContaPagadora") & "-" & Dr("ContaContabilPagadora")
                    DdlContaPagadora.SelectedIndex = DdlContaPagadora.Items.IndexOf(DdlContaPagadora.Items.FindByValue(conta))
                End If

                ddlMoeda.SelectedValue = Dr("Moeda")
                ddlIndexador.SelectedValue = Dr("Indexador")

                ddlCarteiraDoTitulo.SelectedValue = Dr("CarteiraDoTitulo")
                DdlProvisoes.SelectedValue = Dr("Provisao")

                Carteiras(Dr("Pedido"), Dr("Carteira"))
                ddlCarteiras.SelectedIndex = ddlCarteiras.Items.IndexOf(ddlCarteiras.Items.FindByValue(Dr("Carteira")))
                ConfigurarContas(Dr("Empresa"), Dr("EndEmpresa"), Dr("Moeda"), ddlCarteiras.SelectedValue)

                ObjTitulo.CodigoCarteira = Dr("Carteira")

                '*************************************************************************************************
                '*************************************  ADIANTAMENTO  ********************************************
                '*************************************************************************************************
                Dim cart As New CarteiraFinanceira(Dr("Carteira"))

                divSelecaoAdBX.Visible = False
                divAdBaixados.Visible = False
                divAd.Visible = False
                Dim contaadto As String = ""

                If cart.isAdiantamento Then
                    If Dr("Pedido") > 0 Then
                        ddlCarteiras.Enabled = False
                    Else
                        'LIBERAR TROCA DA EMPRESA DO CLIENTE CASO SEJA ADTO, NÃO TENHA PEDIDO E SEJA PREVISÃO - FURLAN - 29/08/2024
                        If Dr("Provisao") = 2 AndAlso Funcoes.VerificaPermissao("LiberaEmpresaAdiantamento", "LEITURA") Then
                            DdlEmpresaCliente.Enabled = True
                        End If
                    End If

                    If cart.BaixaAdiantamento Then
                        contaadto = cart.CodigoContaCliente
                        If Dr("Provisao") = 1 Then
                            divAdBaixados.Visible = True
                        Else
                            divSelecaoAdBX.Visible = True
                        End If
                    Else
                        divAd.Visible = True
                    End If
                End If

                If ObjTitulo.ContaContabilPagadora.Length > 0 AndAlso ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento Then
                    contaadto = ObjTitulo.ContaContabilPagadora
                    If Dr("Provisao") = 1 Then
                        divAdBaixados.Visible = True
                    ElseIf Left(ObjTitulo.ContaContabilPagadora, 1) = IIf(HTabela.Value = "ContasAPagar", "2", "1") Then
                        divSelecaoAdBX.Visible = False
                    Else
                        divSelecaoAdBX.Visible = True
                    End If
                End If

                If divAdBaixados.Visible = True Then
                    gridBaixasAdiantamentos.DataSource = ObjTitulo.BaixasAdiantamento.ToArray
                    gridBaixasAdiantamentos.DataBind()
                    SessaoSalvaTitulo()
                End If

                If divSelecaoAdBX.Visible Then
                    Dim objEmpresa As New [Lib].Negocio.Cliente(Dr("Empresa"), Dr("EndEmpresa"))
                    Dim objCliente As New [Lib].Negocio.Cliente(Dr("Cliente"), Dr("EndCliente"))
                    Dim Moeda As New Moeda(Dr("Moeda"))

                    Dim listadto As New ListAdiantamento(objEmpresa, objCliente, 1, Moeda.Classificacao, IIf(Not IsNumeric(Dr("Pedido")), 0, Dr("Pedido")), contaadto)
                    ObjTitulo.Adiantamentos = listadto

                    If ddlMoeda.SelectedValue = 3 Then
                        For Each adto In ObjTitulo.Adiantamentos
                            adto.VlrSaldoOficial = adto.ValorOficial - adto.VlrBaixaOficial
                        Next
                    End If

                    SessaoSalvaTitulo()

                    gridAdiantamentosDisponiveis.DataSource = listadto.ToArray
                    gridAdiantamentosDisponiveis.DataBind()
                End If


                If cart.isAdiantamento And Not cart.BaixaAdiantamento Then
                    If (Dr("Adiantamento") > 0) Then
                        txtNumeroAdto.Text = Dr("Adiantamento")
                    ElseIf Dr("Provisao") <> 1 Then
                        ConsultaSequenciaDeAdiantamento()
                    End If
                    ddlSafraAdto.Enabled = Not Dr("Pedido") > 0
                    ddlSafraAdto.Parent.Visible = True
                    ddlSafraAdto.SelectedValue = Dr("Safra")
                Else
                    txtNumeroAdto.Text = ""
                    ddlSafraAdto.Parent.Visible = False
                    ddlSafraAdto.SelectedIndex = 0
                End If

                If Not IsDBNull(Dr("VencimentoAdto")) Then
                    txtVencimentoAdto.Text = Dr("VencimentoAdto")
                Else
                    txtVencimentoAdto.Text = ""
                End If

                txtTaxaAdto.Text = Dr("TaxaAdto")

                '*************************************************************************************************
                '*************************************************************************************************
                '*************************************************************************************************
                Session("ssObservacoes" & HID.Value) = Dr("Provisao")

                txtContratoFinanceiro.Text = Dr("ContratoDeFinanciamento")
                txtContratoBanco.Text = Dr("ContratoBanco")

                CargaTributos()

                If Dr("Tributo").ToString.Length > 0 Then
                    DdlTributos.SelectedIndex = DdlTributos.Items.IndexOf(DdlTributos.Items.FindByValue(Dr("Tributo")))
                End If

                ddlUsuarios.Items.Clear()
                ddlUsuarios.Items.Add("Inc.- " & Dr("UsuarioInclusao") & " - " & CDate(Dr("UsuarioInclusaoData")).ToString("dd/MM/yyyy"))
                If Dr("UsuarioLiberacao").ToString.Length > 0 Then ddlUsuarios.Items.Add("Lib.- " & Dr("UsuarioLiberacao") & " - " & CDate(Dr("UsuarioLiberacaoData")).ToString("dd/MM/yyyy"))
                If Dr("UsuarioAlteracao").ToString.Length > 0 Then ddlUsuarios.Items.Add("Alt.- " & Dr("UsuarioAlteracao") & " - " & CDate(Dr("UsuarioAlteracaoData")).ToString("dd/MM/yyyy"))
                If Dr("UsuarioCancelamento").ToString.Length > 0 Then ddlUsuarios.Items.Add("Exc.- " & Dr("UsuarioCancelamento") & " - " & CDate(Dr("UsuarioCancelamentoData")).ToString("dd/MM/yyyy"))

                txtRegistro.Text = Dr("Registro_Id")
                txtMestre.Text = "Mestre : " & Dr("RegistroMestre")
                txtHistorico.Text = Dr("Historico")
                txtCodigoDeBarras.Text = Dr("CodigoDeBarras")

                '****************** Movimento *******************
                txtMovimento.Text = CDate(Dr("Baixa")).ToString("dd/MM/yyyy")
                hdnMovimentoOriginal.Value = CDate(Dr("Baixa")).ToString("dd/MM/yyyy")
                '*********** Vencimento/Prorrogacao *************
                txtProrrogacao.Text = CDate(Dr("Prorrogacao")).ToString("dd/MM/yyyy")
                hdnProrrogacaoOriginal.Value = CDate(Dr("Prorrogacao")).ToString("dd/MM/yyyy")
                '************ Vencimento Original ***************
                lblVencOriginal.Text = CDate(Dr("Vencimento")).ToString("dd/MM/yyyy")
                '************ Data da Baixa ***************
                txtDataEntradaSistema.Text = CDate(Dr("Movimento")).ToString("dd/MM/yyyy")

                txtSolicitacao.Text = Dr("Solicitacao")

                txtCodigoDeBarras.Text = Dr("CodigoDeBarras")

                If Dr("CodigoDigitado") = "S" Then
                    CkbCodigoDeBarras.Checked = True
                Else
                    CkbCodigoDeBarras.Checked = False
                End If

                ckPreImpresso.Checked = Dr("CodigoDeBarraPreImpresso")

                txtEmiteCheque.Value = Dr("Cheque")
                txtSlip.Value = Dr("Slips")

                divCheque.Visible = Dr("TipoPagto") = 2
                txtNumeroCheque.Text = Dr("NumeroDoCheque")

                txtPedido.Text = Dr("Pedido")
                txtPedidoFixacao.Value = Dr("PedidoFixacao")
                Pedido = Dr("Pedido")

                TemRegistro = "S"

                If Not IsDBNull(Dr("Procuracao")) AndAlso Dr("Procuracao") > 0 Then
                    txtCessaoDeCredito.Parent.Visible = True
                    txtCessaoDeCredito.Text = Dr("Procuracao")
                End If

                Session("ssNomeLiberacao" & HID.Value) = Dr("UsuarioLiberacao")                             'Usuario que Liberou

                If IsDBNull(Dr("UsuarioLiberacaoData")) Then
                    Session("ssNomeLiberacaoData" & HID.Value) = CDate(Today).ToString("yyyy/MM/dd")         'Data da Liberacao
                Else
                    Session("ssNomeLiberacaoData" & HID.Value) = Dr("UsuarioLiberacaoData")                 'Data da Liberacao
                End If

                txtUsuarioLiberarTitulo.Value = Dr("UsuarioLiberacaoBloqueio")
                If txtUsuarioLiberarTitulo.Value = "" Then
                    txtUsuarioLiberarTituloData.Value = CDate(Today).ToString("yyyy/MM/dd")
                Else
                    txtUsuarioLiberarTituloData.Value = Dr("UsuarioLiberacaoBloqueioDate")
                End If

                txtUsuarioLiberarPedido.Value = Dr("UsuarioLiberacaoPedido")
                If txtUsuarioLiberarPedido.Value = "" Then
                    txtUsuarioLiberarPedidoData.Value = CDate(Today).ToString("yyyy/MM/dd")
                Else
                    txtUsuarioLiberarPedidoData.Value = Dr("UsuarioLiberacaoPedidoDate")
                End If

                txtUsuarioLiberarCheque.Value = Dr("UsuarioLiberacaoCheque")
                If txtUsuarioLiberarCheque.Value = "" Then
                    txtUsuarioLiberarChequeData.Value = CDate(Today).ToString("yyyy/MM/dd")
                Else
                    txtUsuarioLiberarChequeData.Value = Dr("UsuarioLiberacaoChequeDate")
                End If

                If Not IsDBNull(Dr("Observacoes")) Then
                    txtObservacoes.Text = Dr("Observacoes")
                Else
                    txtObservacoes.Text = ""
                End If

                If Dr("UsuarioLiberacao") <> "" Then
                    txtValorDoDocumento.Enabled = False
                    ChkLiberado.Checked = True
                End If

                Dim objEmpresaCliente = DdlEmpresaCliente.SelectedValue.ToString().Split("-")
                Sql = "Select isnull(IndiceFixado,0) as IndiceFixado, DataPedido From Pedidos where Pedido_id = " & Pedido & "and Empresa_Id = '" & objEmpresaCliente(0) & "'"
                Dim dsPedido As DataSet = Banco.ConsultaDataSet(Sql, "Pedidos")

                If Dr("Moeda") = 1 Then
                    txtValorDoDocumento.ForeColor = System.Drawing.Color.Red
                    txtOficial.ForeColor = System.Drawing.Color.Blue
                End If

                If Dr("Moeda") = 3 AndAlso Not Dr("Indexador") = 99 AndAlso Not Dr("Grupado") = "M" AndAlso Not Left(Session("ssEmpresa"), 8) = "24450490" Then
                    txtValorEmMoeda.ForeColor = System.Drawing.Color.Red
                    txtMoeda.ForeColor = System.Drawing.Color.Blue

                    If CType(Dr("Provisao"), eProvisao) <> eProvisao.Baixa Then
                        If Pedido > 0 Then
                            If dsPedido.Tables(0).Rows.Count > 0 AndAlso CDec(dsPedido.Tables(0).Rows(0).Item("IndiceFixado")) > 0 Then
                                Dr("ValorLiquido") = Math.Round(Dr("MoedaValorDoDocumento") * dsPedido.Tables(0).Rows(0).Item("IndiceFixado"), 2, MidpointRounding.AwayFromZero)
                            Else
                                Dr("ValorLiquido") = Funcoes.ConverteParaMoedaOficial(Dr("MoedaValorDoDocumento"), Dr("Indexador"), dsPedido.Tables(0).Rows(0).Item("DataPedido"))
                            End If

                            If (Dr("ValorDoDocumento") - Dr("ValorLiquido")) < 0 Then
                                Dr("Acrescimos") = (Dr("ValorDoDocumento") - Dr("ValorLiquido")) * -1
                            Else
                                Dr("Deducoes") = (Dr("ValorDoDocumento") - Dr("ValorLiquido"))
                            End If
                        Else
                            If Dr("Indexador") <> 99 Then
                                Dr("ValorDoDocumento") = Funcoes.ConverteParaMoedaOficial(Dr("MoedaValorDoDocumento"), Dr("Indexador"), Dr("Vencimento"))
                            End If
                            Dr("ValorLiquido") = Dr("ValorDoDocumento")
                        End If
                    End If
                End If

                'Valores em Reais
                HDValorOriginalOficial.Value = CDec(Dr("ValorDoDocumento")).ToString("N2")
                txtValorDoDocumento.Text = CDec(Dr("ValorDoDocumento")).ToString("N2")
                txtDescontos.Text = CDec(Dr("Descontos")).ToString("N2")
                txtDeducoes.Text = CDec(Dr("Deducoes")).ToString("N2")
                txtJuros.Text = CDec(Dr("Juros")).ToString("N2")
                txtAcrescimos.Text = CDec(Dr("Acrescimos")).ToString("N2")
                txtValorLiquido.Value = CDec(Dr("ValorLiquido")).ToString("N2")

                If CType(Dr("Provisao"), eProvisao) = eProvisao.Baixa OrElse (CType(Dr("Provisao"), eProvisao) = eProvisao.Previsao And Pedido > 0) Then
                    txtValorCobrado.Text = CDec(Dr("ValorLiquido")).ToString("N2")
                Else
                    txtValorCobrado.Text = CDec(Dr("ValorDoDocumento") + Dr("Juros") + Dr("Acrescimos") - Dr("Descontos") - Dr("Deducoes")).ToString("N2")
                End If

                'Valores em Dólar
                HDValorOriginalMoeda.Value = CDec(Dr("MoedaValorDoDocumento")).ToString("N2")
                txtValorEmMoeda.Text = CDec(Dr("MoedaValorDoDocumento")).ToString("N2")
                txtDescontosMoeda.Text = CDec(Dr("MoedaDescontos")).ToString("N2")
                txtDeducoesMoeda.Text = CDec(Dr("MoedaDeducoes")).ToString("N2")
                txtJurosMoeda.Text = CDec(Dr("MoedaJuros")).ToString("N2")
                txtAcrescimosMoeda.Text = CDec(Dr("MoedaAcrescimos")).ToString("N2")
                txtValorLiquidoMoeda.Value = CDec(Dr("MoedaValorLiquido")).ToString("N2")

                If CType(Dr("Provisao"), eProvisao) = eProvisao.Baixa OrElse (CType(Dr("Provisao"), eProvisao) = eProvisao.Previsao And Pedido > 0) Then
                    txtValorCobradoMoeda.Text = CDec(Dr("MoedaValorLiquido")).ToString("N2")
                Else
                    txtValorCobradoMoeda.Text = CDec(Dr("MoedaValorDoDocumento") + Dr("MoedaJuros") + Dr("MoedaAcrescimos") - Dr("MoedaDescontos") - Dr("MoedaDeducoes")).ToString("N2")
                End If

                If Dr("Moeda") = 1 AndAlso CDec(txtValorEmMoeda.Text) = 0 Then ValidaValores(True)

                lnkNovo.Parent.Visible = True
                lnkExcluir.Parent.Visible = True
                lnkRelatorio.Parent.Visible = True

                txtTipoDoDocumento.Value = Dr("TipoDeDocumento")

                'Titulos gerados pela notas gerais não pode ser alterado o valor do documento pelo financeiro - Furlan - 21/03/2014
                If Dr("NFG") = 1 Then
                    txtValorDoDocumento.Enabled = False
                    txtValorEmMoeda.Enabled = False
                    ddlCarteiras.Enabled = False
                End If

                DivFaturaDeFrete.Style.Add("Display", "none")

                If Dr("Grupado") = "M" Then
                    lblAgrupar.Text = "AP"
                    lnkRelatorioAgrupamento.Parent.Visible = True
                End If

                If CInt(Dr("FaturaDeFrete")) > 0 Then
                    txtCodigoFaturaDeFrete.Text = CInt(Dr("FaturaDeFrete"))
                    DivFaturaDeFrete.Style.Clear()
                    lnkExcluir.Parent.Visible = False
                    cmdPedido.Enabled = False
                    ddlMoeda.Enabled = False
                    ddlIndexador.Enabled = False
                    imgBloqueio.Visible = False
                    ddlCarteiras.Enabled = False
                    txtDeducoes.Enabled = False
                    txtDeducoesMoeda.Enabled = False
                    txtAcrescimos.Enabled = False
                    txtAcrescimosMoeda.Enabled = False

                    If Dr("Grupado") = "S" Then
                        lnkNovo.Parent.Visible = False
                        ImgValores.Enabled = False
                        DdlProvisoes.Enabled = False
                    ElseIf CType(Dr("Provisao"), eProvisao) = eProvisao.Provisao Then
                        DdlProvisoes.Enabled = False
                        txtValorDoDocumento.Enabled = False
                        txtValorEmMoeda.Enabled = False
                    End If
                ElseIf Dr("TipoDeDocumento") = 2 AndAlso Not String.IsNullOrWhiteSpace(Dr("ContratoANTT")) Then
                    lnkNovo.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False
                ElseIf Pedido > 0 Then
                    cmdPedido.Enabled = False
                    imgExtratoPedido.Visible = True
                    ddlMoeda.Enabled = False
                    ddlIndexador.Enabled = False

                    Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(Dr("Carteira"))

                    If objCarteira.isAdiantamento AndAlso Dr("LiberaAdiantamento") = "S" Then
                        lnkExcluir.Parent.Visible = True
                    Else
                        lnkExcluir.Parent.Visible = False
                    End If

                    If Dr("Grupado") = "S" Then
                        lnkNovo.Parent.Visible = False
                        lnkExcluir.Parent.Visible = False
                        ImgValores.Enabled = False
                        DdlProvisoes.Enabled = False
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False

                    ElseIf Dr("Grupado") = "M" Then
                        DdlProvisoes.Enabled = False
                        lnkNovo.Parent.Visible = False
                        ImgValores.Enabled = False
                        imgBloqueio.Visible = True
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                        If CType(Dr("Provisao"), eProvisao) = eProvisao.Previsao Then lnkExcluir.Parent.Visible = True

                    ElseIf CType(Dr("Provisao"), eProvisao) = eProvisao.Baixa Then
                        DdlProvisoes.Enabled = False
                        lnkNovo.Parent.Visible = False
                        ImgValores.Enabled = False
                        imgBloqueio.Visible = True
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                    ElseIf CType(Dr("Provisao"), eProvisao) = eProvisao.Provisao AndAlso Dr("PedidoFixacao") > 0 Then
                        lnkNovo.Parent.Visible = False
                        ImgValores.Enabled = False
                        DdlProvisoes.Enabled = False
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                        MsgBox(Me.Page, "Fixação Pendente Emissão de Nota Fiscal")
                    ElseIf CType(Dr("Provisao"), eProvisao) = eProvisao.Provisao AndAlso Dr("MomentoFinanceiro") = 3 Then
                        lnkNovo.Parent.Visible = False
                        lnkExcluir.Parent.Visible = True
                        ImgValores.Enabled = False
                        DdlProvisoes.Enabled = False
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                        MsgBox(Me.Page, "Titulo para Consumo na Nota Fiscal não pode ser alterado. A Exclusão fica por sua conta e risco.")
                    Else
                        DdlProvisoes.Enabled = True
                        imgBloqueio.Visible = False
                        txtProrrogacao.Enabled = True
                        txtMovimento.Enabled = True

                        'Liberei exclusão da Nota Fiscal - Furlan - 21/10/2020
                        If Dr("NFG") = 0 Then lnkExcluir.Parent.Visible = True
                    End If
                Else
                    cmdPedido.Enabled = True

                    If Dr("Grupado") = "S" Then
                        lnkNovo.Parent.Visible = False
                        ImgValores.Enabled = False
                        lnkExcluir.Parent.Visible = False
                        DdlProvisoes.Enabled = False
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                    ElseIf Dr("Grupado") = "M" Then
                        DdlProvisoes.Enabled = False
                        lnkNovo.Parent.Visible = False
                        ImgValores.Enabled = False
                        imgBloqueio.Visible = True
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                        If CType(Dr("Provisao"), eProvisao) = eProvisao.Previsao Then lnkExcluir.Parent.Visible = True
                    ElseIf CType(Dr("Provisao"), eProvisao) = eProvisao.Baixa Then
                        DdlProvisoes.Enabled = False
                        lnkNovo.Parent.Visible = False
                        ImgValores.Enabled = False
                        imgBloqueio.Visible = True
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                    Else
                        DdlProvisoes.Enabled = True
                        imgBloqueio.Visible = False
                        txtProrrogacao.Enabled = True
                        txtMovimento.Enabled = True
                    End If
                End If

                cmdClientesTitulo.Enabled = True

                If Dr("Pedido") > 0 AndAlso Dr("FinanceiroAberto") = "False" Then
                    lnkNovo.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False
                    ImgValores.Enabled = False
                    MsgBox(Me.Page, "Título com Financeiro Fechado no Pedido não pode ser alterado")
                End If

                ConferenciaNF = CBool(Dr("Conferencia")) AndAlso CBool(Dr("NFG"))

                If lblAgrupar.Text = "AP" Then
                    cmdPedido.Enabled = False

                    Dim Filhos As New ArrayList
                    Dim Mensagem As String = "Agrupamento dos Registros"
                    Dim pedidoAberto As Boolean = True

                    Sql = "Select Registro_id, isnull(EmpresaPedido,'') AS EmpresaPedido, isnull(EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(Pedido, 0) AS Pedido " & vbCrLf &
                          "From " & HTabela.Value & vbCrLf &
                          "Where RegistroMestre = " & txtRegistro.Text

                    Dim dsFilhos As New DataSet
                    dsFilhos = Banco.ConsultaDataSet(Sql, "RegistrosFilhos")

                    If Not dsFilhos Is Nothing AndAlso dsFilhos.Tables(0).Rows.Count > 0 Then
                        For Each drFilho As DataRow In dsFilhos.Tables(0).Rows
                            Filhos.Add(drFilho("Registro_id"))
                            Mensagem &= " - " & drFilho("Registro_id")

                            If drFilho("Pedido") > 0 AndAlso pedidoAberto Then

                                Sql = "Select FinanceiroAberto " & vbCrLf &
                                      "From Pedidos " & HTabela.Value & vbCrLf &
                                      "Where Empresa_id    = '" & drFilho("EmpresaPedido") & "'" & vbCrLf &
                                      "  and EndEmpresa_Id = " & drFilho("EndEmpresaPedido") & vbCrLf &
                                      "  and Pedido_Id     = " & drFilho("Pedido")

                                Dim dsPed As New DataSet
                                dsPed = Banco.ConsultaDataSet(Sql, "PedidosFilhos")

                                If Not dsPed Is Nothing AndAlso dsPed.Tables(0).Rows.Count > 0 Then
                                    For Each drPed As DataRow In dsPed.Tables(0).Rows
                                        pedidoAberto = drPed("FinanceiroAberto")

                                        If Not pedidoAberto Then Exit For
                                    Next
                                End If
                            End If
                        Next

                        If Not pedidoAberto Then
                            lnkNovo.Parent.Visible = False
                            lnkExcluir.Parent.Visible = False
                            imgBloqueio.Visible = False
                            ImgValores.Enabled = False
                            MsgBox(Me.Page, "Título com Financeiro Fechado no Pedido não pode ser alterado")
                        End If

                        Session("ssRegistros" & HID.Value) = Filhos
                        Session("ssObservacoes" & HID.Value) = Mensagem

                        If txtObservacoes.Text <> Mensagem Then
                            txtObservacoes.Text = Mensagem
                        End If
                    End If
                End If

                ObjTitulo.CodigoEmpresaPedido = Dr("EmpresaPedido")
                ObjTitulo.EndEmpresaPedido = Dr("EndEmpresaPedido")

                If Pedido > 0 Then
                    cmdClientesTitulo.Enabled = False

                    Sql = "     SELECT Registro_Id AS Registro, Vencimento, Baixa, Historico, ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, 0.0 as Saldo, Provisao " & vbCrLf &
                          "     FROM " & HTabela.Value & vbCrLf &
                          "     WHERE EmpresaPedido                 = '" & ObjTitulo.CodigoEmpresaPedido & "'" & vbCrLf &
                          "         AND EndEmpresaPedido            = " & ObjTitulo.EndEmpresaPedido & "" & vbCrLf &
                          "         AND Pedido                      = '" & Pedido & "'" & vbCrLf &
                          "         AND Pedido                      <> 0 " & vbCrLf &
                          "         AND Situacao                    = 1 " & vbCrLf &
                          "     ORDER BY Vencimento;"

                    Dim ds As New DataSet
                    Dim dra As DataRow
                    ds = Banco.ConsultaDataSet(Sql, "Pedidos")

                    For Each dra In ds.Tables(0).Rows
                        If dra("Provisao") = 2 Then
                            dra("ValorLiquido") = 0
                        End If

                        SaldoParcelas += dra("ValorDoDocumento") - (dra("ValorLiquido") + dra("Descontos") + dra("Deducoes") - dra("Juros") - dra("Acrescimos"))
                        dra("Saldo") = SaldoParcelas
                    Next

                    GridPedidos.DataSource = ds
                    GridPedidos.DataBind()

                    DdlUnidadeDeNegocioEmpresaCliente.Enabled = False
                    DdlEmpresaCliente.Enabled = False
                End If

                Sql = " SELECT R.Empresa_Id + '-' + cast(R.EndEmpresa_Id as varchar) as Empresa, R.Conta_Id, R.Cliente_Id, R.EndCliente_Id, " & vbCrLf &
                      "        PC.Titulo, R.Movimento_Id, R.Lote_Id, ISNULL(R.Produto, '') AS Produto, R.Custo, R.Historico, " & vbCrLf &
                      "        R.DebitoOficial, R.CreditoOficial, isnull(Conciliacao,'') AS Conciliacao, isnull(DataDaBaixa,CURRENT_TIMESTAMP) AS DataDaBaixa " & vbCrLf &
                      "   FROM Razao R" & vbCrLf &
                      "  INNER JOIN PlanoDeContas PC " & vbCrLf &
                      "     ON R.Conta_Id = PC.Conta_Id " & vbCrLf &
                      "  WHERE (R.Titulo = " & Registro & ") " & vbCrLf

                Dim dsRazao As New DataSet
                dsRazao = Banco.ConsultaDataSet(Sql, "Razao")

                If Not dsRazao Is Nothing AndAlso dsRazao.Tables(0).Rows.Count > 0 Then
                    Dim dtRazao As New DataTable("razao")
                    Dim saldo As Double = 0

                    dtRazao.Columns.Add("Empresa", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Conta", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Cliente", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Movimento", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Titulo", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Lote", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Produto", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Custo", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Historico", Type.GetType("System.String"))
                    dtRazao.Columns.Add("Debito", Type.GetType("System.Double"))
                    dtRazao.Columns.Add("Credito", Type.GetType("System.Double"))
                    dtRazao.Columns.Add("Saldo", Type.GetType("System.Double"))

                    For Each row As DataRow In dsRazao.Tables(0).Rows
                        Dim drRazao As DataRow = dtRazao.NewRow()

                        drRazao("Empresa") = row("Empresa")
                        drRazao("Conta") = row("Conta_Id")
                        drRazao("Cliente") = row("Cliente_Id") & "-" & row("EndCliente_Id")
                        drRazao("Movimento") = row("Movimento_Id")
                        drRazao("Titulo") = row("Titulo")
                        drRazao("Lote") = row("Lote_Id")
                        drRazao("Produto") = row("Produto")
                        drRazao("Custo") = row("Custo")
                        drRazao("Historico") = row("Historico")
                        drRazao("Debito") = row("DebitoOficial")
                        drRazao("Credito") = row("CreditoOficial")
                        saldo = Math.Round(saldo + (row("DebitoOficial") - row("CreditoOficial")), 2)
                        drRazao("Saldo") = saldo
                        dtRazao.Rows.Add(drRazao)

                        If row("Conciliacao") = "B" Then
                            chkConciliado.Checked = True
                            txtDataConciliacao.Value = row("DataDaBaixa")
                        End If
                    Next
                    gridRazao.DataSource = dtRazao
                    gridRazao.DataBind()
                End If

                If txtFornecedor.Text.Length > 0 Then
                    Sql = "SELECT Registro_Id, Cliente, EndCliente, Carteira " & vbCrLf &
                          "  FROM TitulosXDesdobrarFornecedor " & vbCrLf &
                          " WHERE Registro_Id = " & txtRegistro.Text & vbCrLf

                    Dim dsDesdobrarFornecedor As New DataSet
                    dsDesdobrarFornecedor = Banco.ConsultaDataSet(Sql, "DesdobrarFornecedor")

                    If Not dsDesdobrarFornecedor Is Nothing AndAlso dsDesdobrarFornecedor.Tables(0).Rows.Count > 0 Then
                        Session("objDestinoContabil" & HID.Value) = dsDesdobrarFornecedor.Tables(0).Rows(0).Item("Cliente") & "-" &
                                                        dsDesdobrarFornecedor.Tables(0).Rows(0).Item("EndCliente") & "-" &
                                                        dsDesdobrarFornecedor.Tables(0).Rows(0).Item("Carteira")
                    End If
                End If

                txtRegistro.Enabled = False

                MostrarCotacao()

                'Registro liberado para pagamento 
                If ChkLiberado.Checked = True Then
                    If DdlProvisoes.SelectedValue = 2 AndAlso ddlMoeda.SelectedValue = 3 Then
                        'txtDataBaixa.Enabled = True
                        txtValorDoDocumento.Enabled = True
                        txtDescontos.Enabled = True
                        txtDeducoes.Enabled = True
                        txtJuros.Enabled = True
                        txtAcrescimos.Enabled = True
                    Else
                        'txtDataBaixa.Enabled = False
                        txtValorDoDocumento.Enabled = False
                        txtDescontos.Enabled = False
                        txtDeducoes.Enabled = False
                        txtJuros.Enabled = False
                        txtAcrescimos.Enabled = False
                    End If

                    txtValorCobrado.Enabled = False
                    txtSolicitacao.Enabled = False
                    txtValorEmMoeda.Enabled = False
                    txtDescontosMoeda.Enabled = False
                    txtDeducoesMoeda.Enabled = False
                    txtJurosMoeda.Enabled = False
                    txtAcrescimosMoeda.Enabled = False
                    txtValorCobradoMoeda.Enabled = False

                    If (txtPedido.Text.Length > 0 AndAlso CInt(txtPedido.Text) > 0) And Funcoes.VerificaPermissao("AlterarCarteiraCP", "ALTERAR") = False Then
                        ddlCarteiras.Enabled = False
                        DdlTributos.Enabled = False
                        txtHistorico.Enabled = False
                    Else
                        If cart Is Nothing Then cart = New CarteiraFinanceira(Dr("Carteira"))
                        If (txtPedido.Text.Length > 0 AndAlso CInt(txtPedido.Text) > 0) AndAlso Not cart.isAdiantamento Then
                            ddlCarteiras.Enabled = True
                        End If

                        DdlTributos.Enabled = True
                        txtHistorico.Enabled = True
                    End If
                End If

                If DdlProvisoes.SelectedValue = 3 AndAlso Not Funcoes.VerificaPermissao("AutorizacoesDePagamentos", "GRAVAR") Then
                    lnkNovo.Parent.Visible = False 'Provisao Bloqueada P/ Alteracao.
                End If

                idRetornoBancario.Visible = False
                txtObsRetornoBancario.Text = ""

                If Not dsTitulo Is Nothing AndAlso dsTitulo.Tables(0).Rows.Count > 0 AndAlso Not dsTitulo.Tables(0).Rows(0).Item("Situacao") = 1 Then
                    lnkNovo.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False
                    DdlProvisoes.Enabled = False
                    'txtDataBaixa.Enabled = False
                    lnkRelatorio.Parent.Visible = False
                    lnkRelatorioAgrupamento.Parent.Visible = False
                    pnlReprogramaVencimentos.Visible = False
                    lnkReprogramar.Enabled = False
                    imgBloqueio.Visible = False
                    txtMovimento.Enabled = False
                    txtProrrogacao.Enabled = False
                    ImgValores.Enabled = False
                    If dsTitulo.Tables(0).Rows(0).Item("Situacao") = 101 Then
                        txtMestre.Text = "TÍTULO EM COBRANÇA"

                        idRetornoBancario.Visible = True
                        txtObsRetornoBancario.Text = dsTitulo.Tables(0).Rows(0).Item("HistoricoRemessa")


                        If Funcoes.VerificaPermissao("LiberarCobranca", "GRAVAR") Then
                            lnkSituacaoBancaria.Parent.Visible = True
                        Else
                            lnkSituacaoBancaria.Parent.Visible = False
                        End If
                    ElseIf dsTitulo.Tables(0).Rows(0).Item("Situacao") = 102 Then
                        txtMestre.Text = "TÍTULO DE ENDOSSO"
                        lnkSituacaoBancaria.Parent.Visible = False

                        Sql = "Codigo_Id = " & Dr("SequenciaEndosso")

                        Dim endossos As New ListEndosso(Dr("Empresa"), Dr("EndEmpresa"), Sql)

                        If endossos.Count > 0 Then
                            For Each eD In endossos
                                'If eD.Cliente.CodigoFormatado.Length = 0 Then Dim erro = "erro - só para instância da propriedade ClienteDescricao READ ONLY"

                                If eD.ClienteEndosso.CodigoFormatado.Length = 0 Then Dim erro = "erro - só para instância da propriedade ClienteEndossoDescricao READ ONLY"
                            Next

                            divEndosso.Visible = True
                            gridEndosso.DataSource = endossos.ToArray
                            gridEndosso.DataBind()
                        End If
                    ElseIf dsTitulo.Tables(0).Rows(0).Item("Situacao") = 105 Then
                        txtMestre.Text = "TÍTULO EM FOLHA DE PAGAMENTO"

                        idRetornoBancario.Visible = True
                        txtObsRetornoBancario.Text = dsTitulo.Tables(0).Rows(0).Item("HistoricoRemessa")

                        'POR HORA VOU DEIXAR BLOQUEADO PARA VERMOS COMO VAI FICAR O PROCESSO - FURLAN - 05/11/2025
                        'If Funcoes.VerificaPermissao("LiberarCobranca", "GRAVAR") Then
                        '    lnkSituacaoBancaria.Parent.Visible = True
                        'Else
                        '    lnkSituacaoBancaria.Parent.Visible = False
                        'End If
                    Else
                        txtMestre.Text = "CANCELADO"
                    End If
                End If

                'NF conferida pelo fiscal
                If Not ConferenciaNF Then
                    ViewState.Add("CONTROLE_FISCAL", "BLOQUEADO")
                End If

                lnkConsultar.Parent.Visible = False

                ChkLiberado.Parent.Visible = True
                chkConciliado.Parent.Visible = True

                TabContainer1.ActiveTabIndex = 0
                BloquearValores(True)

                'CARREGAR DOCUMENTOS
                If ObjTitulo.FinanceiroXDocumentos.Count > 0 Then
                    For Each fXd In ObjTitulo.FinanceiroXDocumentos
                        fXd.IUD = "I"
                    Next
                    CarregarDocumentos()
                End If

                SessaoSalvaTitulo()

                ConsultaArquivoNFE(ObjTitulo.Codigo)

                'Boleto bancário
                Dim boletoBancario = dsTitulo.Tables(0).Rows(0).Item("BoletoBancario")
                If boletoBancario AndAlso Not dsTitulo.Tables(0).Rows(0).Item("Situacao") = 105 Then
                    imgBoletoPDF.Visible = True
                    DdlEmpresaPagadora.Enabled = False
                    DdlTiposDePagamentos.Enabled = False
                    DdlBancoPagador.Enabled = False
                    DdlContaPagadora.Enabled = False
                Else
                    imgBoletoPDF.Visible = False
                End If
            Else
                MsgBox(Me.Page, "Informe o Numero do Registro para consulta")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao consultar registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(ex.Message.ToString())) & ". Entre em contato com o Suporte.")
        End Try
    End Sub

    Private Sub ConsultaArquivoNFE(codigo As String)

        objNotaFiscal = New [Lib].Negocio.NotaFiscal(codigo)

        If objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.Arquivos.Count > 0 Then
            ucFile.Parent.Visible = True
            ucFile.Bind(objNotaFiscal.Arquivos)
            divArquivos.Visible = True
        End If

    End Sub

    Public Property SessaoDsXML() As DataSet
        Get
            Return CType(Session("dsXML" & HID.Value), DataSet)
        End Get
        Set(ByVal value As DataSet)
            If value Is Nothing Then
                Session.Remove("dsXml" & HID.Value)
            Else
                Session("dsXml" & HID.Value) = value
            End If
        End Set
    End Property

    Protected Sub lnkConsultarNaviosXInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkConsultarNaviosXInvoice.Click
        Try
            Limpar(True)
            Popup.ConsultarNaviosXInvoice(Me.Page, "objConsultarNaviosXInvoice")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlUnidadeConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresaConsultaTitulos.Items.Clear()

        Sql = " SELECT Cli.Cliente_Id as Codigo, Cli.Endereco_Id, Cli.Reduzido, Cli.Nome, Cli.Cidade, Cli.Estado " & vbCrLf &
              "   FROM GruposXEmpresas GxE" & vbCrLf &
              "  INNER JOIN Clientes Cli " & vbCrLf &
              "     ON GxE.Cliente_Id = Cli.Cliente_Id " & vbCrLf &
              "    AND GxE.EndCliente_Id = Cli.Endereco_Id" & vbCrLf &
              "  WHERE GxE.Empresa_Id = '" & DdlUnidadeConsultaTitulos.SelectedValue & "' " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

            DdlEmpresaConsultaTitulos.Items.Add(New ListItem(Descricao, Codigo))

        Next

        DdlEmpresaConsultaTitulos.Items.Insert(0, "")
        DdlEmpresaConsultaTitulos.SelectedIndex = 0

    End Sub

    Protected Sub cmdClientesTitulo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo" & HID.Value) = "LivreClasse"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objFornecedorCP" & HID.Value, "txtNome")
    End Sub

    Protected Sub cmdConsultaClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo" & HID.Value) = "LivreClasse"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteCTAXPC" & HID.Value, "txtNome")
    End Sub

    Protected Sub chkClienteEndosso_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If chkClienteEndosso.Checked Then
                divClienteConsultaEndosso.Visible = True
            Else
                divClienteConsultaEndosso.Visible = False
            End If

            txtClienteConsultaEndosso.Text = String.Empty
            txtCodigoClienteConsultaEndosso.Value = String.Empty

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaClientesEndosso_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo" & HID.Value) = "LivreClasse"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteCTAXPCXED" & HID.Value, "txtNome")
    End Sub

    Private Sub Limpar_ConsultaTitulos(ByVal limparTudo As Boolean)
        If limparTudo Then
            If Not DdlEmpresaConsultaTitulos.SelectedIndex > 0 Then Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeConsultaTitulos, DdlEmpresaConsultaTitulos, True)
            txtClienteConsulta.Text = String.Empty
            txtCodigoClienteConsulta.Value = String.Empty
            chkClienteEndosso.Checked = False
            txtClienteConsultaEndosso.Text = String.Empty
            txtCodigoClienteConsultaEndosso.Value = String.Empty
            divClienteConsultaEndosso.Visible = False
            txtPeriodoInicialConsultaTitulos.Text = String.Format("{0}/{1}/{2}", "01", Month(Now).ToString.PadLeft(2, "0"), Year(Now))
            txtPeriodoFinalConsultaTitulos.Text = Now.ToString("dd/MM/yyyy")
        End If

        lblTotalRegistroAgrupado.Text = String.Empty
        lblTotalRegistroAgrupado.Parent.Visible = False

        rowDolar.Visible = False
        txtRealDolar.Value = String.Empty
        txtPedido.Text = String.Empty
        txtNumNota.Text = String.Empty
        txtNovoVencimento.Text = String.Empty
        HiddenIndexador.Value = String.Empty
        pnlReprogramaVencimentos.Visible = False
        lnkReprogramar.Parent.Visible = False

        lnkSlip.Parent.Visible = True
        lnkRecibo.Parent.Visible = False
        lnkAgruparPagamento.Parent.Visible = False
        lnkBaixarSelecionados.Parent.Visible = False

        GridConsultaTitulos.DataSource = Nothing
        GridConsultaTitulos.DataBind()
    End Sub

    Protected Sub DdlBancoPagador_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlBancoPagador.SelectedIndexChanged
        BancoPagador(True)
    End Sub

    Private Sub BancoPagador(Optional ByVal SomenteAtivas As Boolean = False)
        Dim Cliente As String
        Dim Campo() As String
        Dim Conta As String

        Cliente = DdlEmpresaPagadora.SelectedValue
        Campo = Cliente.Split("-")
        DdlContaPagadora.Items.Clear()

        If Campo(0) <> "" Then
            Sql = " SELECT bxc.Agencia_Id, bxc.DigitoAgencia_Id, bxc.Conta_Id, bxc.DigitoConta_Id, isnull(bxc.TipoConta,'C') as TipoConta, bxc.ContaContabil, bxc.Observacoes, Pc.Titulo" & vbCrLf &
                  "   FROM BancosXContas bxc" & vbCrLf &
                  "  Inner Join PlanoDeContas PC" & vbCrLf &
                  "     on Bxc.ContaContabil = Pc.Conta_Id" & vbCrLf &
                  "  WHERE bxc.Empresa_Id    = '" & Campo(0) & "'" & vbCrLf &
                  "    AND bxc.EndEmpresa_Id = " & Campo(1) & vbCrLf &
                  "    AND bxc.Banco_Id      =  " & DdlBancoPagador.SelectedValue & vbCrLf
            '"    AND (isnull(pc.Adiantamento,0) = 0 or (pc.Adiantamento = 1 and left(pc.Conta_Id,1) = '" & IIf(HTabela.Value = "ContasAPagar", "1", "2") & "'))" & vbCrLf
            'Desabilitei para carregar todas as Agencias e tratar as excessões - furlan - 06/05/2020
            If SomenteAtivas Then
                Sql &= "   AND bxc.Ativo = 1 " & vbCrLf
            End If


            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "BancosXContas").Tables(0).Rows
                Conta = Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "-" & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & "-" & Dr("TipoConta") & "-" & Dr("ContaContabil")
                Descricao = "AG " & Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & " " & Dr("TipoConta") & " " & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & " - " & Dr("Observacoes") & " - (" & Dr("ContaContabil") & "-" & Dr("Titulo") & ")"
                DdlContaPagadora.Items.Add(New ListItem(Descricao, Conta))
            Next

            DdlContaPagadora.Items.Insert(0, "")
            DdlContaPagadora.SelectedIndex = 0
        End If
    End Sub

    Public Sub BloquearValores(Bloquear As Boolean)
        If Bloquear Then
            ImgValores.ImageUrl = ImgValores.ImageUrl.Replace("Bloquear", "liberar")
            ImgValores.ToolTip = "Liberar Valores"
            txtValorDoDocumento.Enabled = False
            txtDescontos.Enabled = False
            txtDeducoes.Enabled = False
            txtJuros.Enabled = False
            txtAcrescimos.Enabled = False
            txtValorEmMoeda.Enabled = False
            txtDescontosMoeda.Enabled = False
            txtDeducoesMoeda.Enabled = False
            txtJurosMoeda.Enabled = False
            txtAcrescimosMoeda.Enabled = False
        Else
            ImgValores.ImageUrl = ImgValores.ImageUrl.Replace("liberar", "Bloquear")
            ImgValores.ToolTip = "Bloquear Valores"
            txtValorDoDocumento.Enabled = True
            txtDescontos.Enabled = True
            txtDeducoes.Enabled = True
            txtJuros.Enabled = True
            txtAcrescimos.Enabled = True
            txtValorEmMoeda.Enabled = True
            txtDescontosMoeda.Enabled = True
            txtDeducoesMoeda.Enabled = True
            txtJurosMoeda.Enabled = True
            txtAcrescimosMoeda.Enabled = True
        End If

        If Session("ssGrupado" & HID.Value) = "S" Then
            txtValorDoDocumento.Enabled = False
            txtDescontos.Enabled = False
            txtDeducoes.Enabled = False
            txtJuros.Enabled = False
            txtAcrescimos.Enabled = False
            txtValorCobrado.Enabled = False

            txtValorEmMoeda.Enabled = False
            txtDescontosMoeda.Enabled = False
            txtDeducoesMoeda.Enabled = False
            txtJurosMoeda.Enabled = False
            txtAcrescimosMoeda.Enabled = False
            txtValorCobradoMoeda.Enabled = False
        End If
    End Sub

    Private Sub ValidaValores(ByVal dolarizar As Boolean)

        Dim Zero As Decimal = 0
        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)

        txtValorDoDocumento.Text = Funcoes.AtribuirirValorFormatado(txtValorDoDocumento.Text)
        txtDescontos.Text = Funcoes.AtribuirirValorFormatado(txtDescontos.Text)
        txtJuros.Text = Funcoes.AtribuirirValorFormatado(txtJuros.Text)
        txtAcrescimos.Text = Funcoes.AtribuirirValorFormatado(txtAcrescimos.Text)
        txtDeducoes.Text = Funcoes.AtribuirirValorFormatado(txtDeducoes.Text)
        txtValorCobrado.Text = Funcoes.AtribuirirValorFormatado(txtValorCobrado.Text)
        txtValorEmMoeda.Text = Funcoes.AtribuirirValorFormatado(txtValorEmMoeda.Text)
        txtDescontosMoeda.Text = Funcoes.AtribuirirValorFormatado(txtDescontosMoeda.Text)
        txtJurosMoeda.Text = Funcoes.AtribuirirValorFormatado(txtJurosMoeda.Text)
        txtAcrescimosMoeda.Text = Funcoes.AtribuirirValorFormatado(txtAcrescimosMoeda.Text)
        txtDeducoesMoeda.Text = Funcoes.AtribuirirValorFormatado(txtDeducoesMoeda.Text)
        txtValorCobradoMoeda.Text = Funcoes.AtribuirirValorFormatado(txtValorCobradoMoeda.Text)

        'NÃO GERAR CORREÇÃO MONETÁRIA AUTOMÁTICA PARA RTGRÃOS - FURLAN - 08/08/2024
        If Left(Session("ssEmpresa"), 8) = "24450490" Then dolarizar = False
        If Funcoes.VerificaPermissao("NaoDolarizarFinanceiro", "ACESSAR") Then dolarizar = False

        If dolarizar Then
            If ddlMoeda.SelectedValue = 1 AndAlso CDec(txtValorEmMoeda.Text) = 0 Then
                Dim vlr As Decimal = DolarizaBaixa(CDate(txtMovimento.Text).ToString("yyyy-MM-dd"), txtValorDoDocumento.Text, 2)
                txtValorEmMoeda.Text = vlr.ToString("N2")
            ElseIf ddlMoeda.SelectedValue = 3 AndAlso (DdlProvisoes.Text = 2 Or DdlProvisoes.Text = 3) And ChkLiberado.Checked = False Then
                If txtPedido.Text > 0 Then
                    Sql = "Select DataPedido, isnull(IndiceFixado,0) as IndiceFixado From Pedidos where Pedido_id = " & txtPedido.Text
                    Dim dsPedido As DataSet = Banco.ConsultaDataSet(Sql, "Pedidos")

                    If dsPedido.Tables(0).Rows.Count > 0 AndAlso CDec(dsPedido.Tables(0).Rows(0).Item("IndiceFixado")) > 0 Then
                        txtValorCobrado.Text = Math.Round(CDbl(txtValorEmMoeda.Text) * dsPedido.Tables(0).Rows(0).Item("IndiceFixado"), 2, MidpointRounding.AwayFromZero)
                    Else
                        If Not ddlIndexador.SelectedValue = 99 Then txtValorCobrado.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtValorEmMoeda.Text), ddlIndexador.SelectedValue, dsPedido.Tables(0).Rows(0).Item("DataPedido")), 2)
                    End If

                    If Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaDesconto) _
                       AndAlso Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaDeducao) _
                       AndAlso Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaAcrescimo) _
                       AndAlso Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaJuro) Then

                        If CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text) < 0 Then
                            txtAcrescimos.Text = ((CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text)) * -1).ToString("N2")
                        Else
                            txtDeducoes.Text = ((CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text))).ToString("N2")
                        End If
                    End If

                    dolarizar = False
                End If
            End If

            If ddlMoeda.SelectedValue = 1 Then
                If ddlIndexador.SelectedValue = 99 Then

                    If CDec(txtDescontos.Text) > 0 AndAlso CDec(txtValorEmMoeda.Text) > 0 Then
                        txtDescontosMoeda.Text = Math.Round(CDec(txtDescontos.Text) / (CDec(txtValorDoDocumento.Text) / CDec(txtValorEmMoeda.Text)), 2, MidpointRounding.AwayFromZero)
                    Else
                        txtDescontosMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtDeducoes.Text) > 0 AndAlso CDec(txtValorEmMoeda.Text) > 0 Then
                        txtDeducoesMoeda.Text = Math.Round(CDec(txtDeducoes.Text) / (CDec(txtValorDoDocumento.Text) / CDec(txtValorEmMoeda.Text)), 2, MidpointRounding.AwayFromZero)
                    Else
                        txtDeducoesMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtJuros.Text) > 0 AndAlso CDec(txtValorEmMoeda.Text) > 0 Then
                        txtJurosMoeda.Text = Math.Round(CDec(txtJuros.Text) / (CDec(txtValorDoDocumento.Text) / CDec(txtValorEmMoeda.Text)), 2, MidpointRounding.AwayFromZero)
                    Else
                        txtJurosMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtAcrescimos.Text) > 0 AndAlso CDec(txtValorEmMoeda.Text) > 0 Then
                        txtAcrescimosMoeda.Text = Math.Round(CDec(txtAcrescimos.Text) / (CDec(txtValorDoDocumento.Text) / CDec(txtValorEmMoeda.Text)), 2, MidpointRounding.AwayFromZero)
                    Else
                        txtAcrescimosMoeda.Text = Zero.ToString("N2")
                    End If
                Else
                    If CDec(txtDescontos.Text) > 0 Then
                        txtDescontosMoeda.Text = DolarizaBaixa(CDate(txtMovimento.Text).ToString("yyyy-MM-dd"), txtDescontos.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                    Else
                        txtDescontosMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtDeducoes.Text) > 0 Then
                        txtDeducoesMoeda.Text = DolarizaBaixa(CDate(txtMovimento.Text).ToString("yyyy-MM-dd"), txtDeducoes.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                    Else
                        txtDeducoesMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtJuros.Text) > 0 Then
                        txtJurosMoeda.Text = DolarizaBaixa(CDate(txtMovimento.Text).ToString("yyyy-MM-dd"), txtJuros.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                    Else
                        txtJurosMoeda.Text = Zero.ToString("N2")
                    End If
                    If CDec(txtAcrescimos.Text) > 0 Then
                        txtAcrescimosMoeda.Text = DolarizaBaixa(CDate(txtMovimento.Text).ToString("yyyy-MM-dd"), txtAcrescimos.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue))
                    Else
                        txtAcrescimosMoeda.Text = Zero.ToString("N2")
                    End If
                End If
            ElseIf ddlMoeda.SelectedValue = 3 Then

                If CDec(txtDescontosMoeda.Text) > 0 Then
                    txtDescontos.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtDescontosMoeda.Text), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue), CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd")), 2)
                Else
                    txtDescontos.Text = Zero.ToString("N2")
                End If

                If CDec(txtDeducoesMoeda.Text) > 0 Then
                    txtDeducoes.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtDeducoesMoeda.Text), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue), CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd")), 2)
                ElseIf Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaAcrescimo) _
                       AndAlso Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaJuro) _
                       AndAlso Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaDesconto) _
                       AndAlso Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaDeducao) Then
                    txtDeducoes.Text = Zero.ToString("N2")
                End If

                If CDec(txtJurosMoeda.Text) > 0 Then
                    txtJuros.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtJurosMoeda.Text), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue), CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd")), 2)
                Else
                    txtJuros.Text = Zero.ToString("N2")
                End If

                If CDec(txtAcrescimosMoeda.Text) > 0 Then
                    txtAcrescimos.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtAcrescimosMoeda.Text), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue), CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd")), 2)
                ElseIf Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaAcrescimo) _
                       AndAlso Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaJuro) _
                       AndAlso Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaDesconto) _
                       AndAlso Not IsVariacao(DdlEmpresaCliente.SelectedValue, objCarteira.CodigoContaDeducao) Then
                    txtAcrescimos.Text = Zero.ToString("N2")
                End If
            End If
        End If

        txtValorCobrado.Text = (CDec(txtValorDoDocumento.Text) + CDec(txtJuros.Text) + CDec(txtAcrescimos.Text) - CDec(txtDescontos.Text) - CDec(txtDeducoes.Text)).ToString("N2")
        txtValorCobradoMoeda.Text = (CDec(txtValorEmMoeda.Text) + CDec(txtJurosMoeda.Text) + CDec(txtAcrescimosMoeda.Text) - CDec(txtDescontosMoeda.Text) - CDec(txtDeducoesMoeda.Text)).ToString("N2")

        If lnkNovo.Parent.Visible = False Then lnkNovo.Parent.Visible = True
    End Sub

    Private Function TemNota(ByVal Titulo As String) As Boolean
        Dim ds As New DataSet

        Sql = "SELECT Nota_Id FROM NotaFiscalXTitulo where Titulo_Id  = '" & Titulo & "'"

        ds = Banco.ConsultaDataSet(Sql, "NotaFiscalXTitulo")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            MsgBox(Me.Page, "Registro Não pode se Excluido, Existe uma Nota " & ds.Tables(0).Rows(0).Item("Nota_Id") & " Vinculada a este.")
            Return True
        Else
            Return False
        End If
    End Function

    Protected Sub ddlCarteiras_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlCarteiras.SelectedIndexChanged
        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)
        'DdlTributos.Parent.Visible = Not objCarteira.isAdiantamento

        SessaoRecuperaTitulo()

        If ddlCarteiras.SelectedValue.Length > 0 Then
            ObjTitulo.CodigoCarteira = objCarteira.CodigoCarteira
            SessaoSalvaTitulo()

            SessaoRecuperaTitulo()
        End If

        divAd.Visible = (objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento)

        txtVencimentoAdto.Enabled = (objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento)

        DdlTributos.DataSource = Nothing
        DdlTributos.Items.Clear()

        If Not objCarteira.isAdiantamento Then
            CargaTributos()
            DdlTributos.Parent.Visible = True
        Else
            DdlTributos.Parent.Visible = False
        End If

        If objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento Then
            ConsultaSequenciaDeAdiantamento()
        End If

        If (objCarteira.isAdiantamento And objCarteira.BaixaAdiantamento) AndAlso Not (ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento) Then
            divSelecaoAdBX.Visible = True

            Dim strEmpresa() As String = DdlEmpresaCliente.SelectedValue.Split("-")
            Dim objEmpresa As New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
            Dim strCliente() As String = txtCodigoFornecedor.Value.Split("-")
            Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
            Dim Moeda As New Moeda(ddlMoeda.SelectedValue)
            Dim listadto As New ListAdiantamento(objEmpresa, objCliente, 1, Moeda.Classificacao, IIf(Not IsNumeric(txtPedido.Text), 0, txtPedido.Text), objCarteira.CodigoContaCliente)

            ObjTitulo.Adiantamentos = listadto
            SessaoSalvaTitulo()

            gridAdiantamentosDisponiveis.DataSource = listadto.ToArray
            gridAdiantamentosDisponiveis.DataBind()
        ElseIf Not (ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento) Then
            divSelecaoAdBX.Visible = False
        End If

        If Session("objDestinoContabil" & HID.Value) Is Nothing AndAlso DdlProvisoes.SelectedValue = 1 Then
            Dim strCliente() As String = txtCodigoFornecedor.Value.Split("-")
            Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
            If objCliente.DesdobrarFornecedor = True Then
                ucDestinoContabil.Limpar()
                Dim parameters As New Dictionary(Of String, Object)
                If Request.QueryString("idTabela").Equals("ContasAPagar") Then
                    parameters.Add("tipo", "P")
                Else
                    parameters.Add("tipo", "R")
                End If
                Popup.ConsultaDeDestinoContabil(Me.Page, "objDestinoContabil" & HID.Value)
                ucDestinoContabil.Carregar(parameters)
            End If
        End If

        ConfigurarContas(DdlEmpresaCliente.SelectedValue.Split("-")(0), DdlEmpresaCliente.SelectedValue.Split("-")(1), ddlMoeda.SelectedValue, ddlCarteiras.SelectedValue)
    End Sub

    Private Sub ConsultaSequenciaDeAdiantamento()
        Dim emp() As String = DdlEmpresaCliente.SelectedValue.ToString.Split("-")
        Dim NumAdiantamento As Integer

        If IsNumeric(txtRegistro.Text) AndAlso CInt(txtRegistro.Text) > 0 Then
            Sql = "SELECT isnull(Adiantamento_id,0) as Adiantamento" & vbCrLf &
                  "  FROM " & HTabela.Value & " cp" & vbCrLf &
                  "  Left join vw_Adiantamento A" & vbCrLf &
                  "    on cp.registro_id = a.titulo " & vbCrLf &
                  " WHERE cp.Registro_id = " & txtRegistro.Text

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Adto").Tables(0).Rows
                NumAdiantamento = Dr("Adiantamento")
            Next
        End If

        If NumAdiantamento = 0 Then
            Dim num As New [Lib].Negocio.Numerador(emp(0), emp(1), 15)
            NumAdiantamento = num.Sequencia + 1
        End If

        txtNumeroAdto.Text = NumAdiantamento
    End Sub

    Private Sub CargaTributos()
        Dim dsTributo As New DataSet

        Sql = " SELECT Tributo_Id as Codigo, " & vbCrLf &
              "       case" & vbCrLf &
              "			when " + IIf(HTabela.Value = "ContasAReceber", "isnull(Encargos.ContaCredito,'')", "isnull(Encargos.ContaDebito,'')") + " = ''" & vbCrLf &
              "			   then Encargos.Descricao + ' - ' + Tributo_Id" & vbCrLf &
              "			   else Encargos.Descricao + ' - ' + Tributo_Id + ' (' + " + IIf(HTabela.Value = "ContasAReceber", "Encargos.ContaCredito", "Encargos.ContaDebito") + " + '-' + p.Titulo + ')'" & vbCrLf &
              "		  end AS Descricao" & vbCrLf &
              "   FROM CarteirasXTributos " & vbCrLf &
              "  INNER JOIN Encargos " & vbCrLf &
              "     ON CarteirasXTributos.Tributo_ID = Encargos.Encargo_id " & vbCrLf &
              "  LEFT JOIN PlanoDeContas p" & vbCrLf &
              "         on p.Conta_Id = " & IIf(HTabela.Value = "ContasAReceber", "Encargos.ContaCredito", "Encargos.ContaDebito") & vbCrLf &
              "  WHERE Carteira_Id = '" & ddlCarteiras.SelectedValue & "'" & vbCrLf &
              "  ORDER BY Tributo_Id " & vbCrLf

        dsTributo = Banco.ConsultaDataSet(Sql, "Tributos")
        DdlTributos.Items.Clear()
        DdlTributos.DataValueField = "Codigo"
        DdlTributos.DataTextField = "Descricao"
        DdlTributos.DataSource = dsTributo
        DdlTributos.DataBind()
        Funcoes.InserirLinhaEmBranco(DdlTributos)
        DdlTributos.Parent.Visible = dsTributo.Tables(0).Rows.Count > 0
    End Sub

    Private Function ValidaData(ByVal pData As String, ByVal Tipo As String, ByVal Empresa As String, ByVal EndEmpresa As String) As Boolean
        If IsDate(pData) Then
        Else
            Mensagem = " Data de " & Tipo & " inválida..."
            Return False
        End If

        If CDate(pData).DayOfWeek = 6 Then
            If Left(Session("ssEmpresa").ToString, 8) = "24450490" AndAlso DdlProvisoes.SelectedValue = 1 AndAlso CDate(txtMovimento.Text) < CDate("01/07/2024") Then
                'NÃO FAZ NADA RT GRÃOS IMPLANTAÇÃO SALDO INICIAL - FURLAN - 18/07/2024
            Else
                Mensagem = "Sábado - Data Inválida para " & Tipo & "..."
                Return False
            End If
        End If

        If CDate(pData.Replace("'", "")).DayOfWeek = 0 Then
            If Left(Session("ssEmpresa").ToString, 8) = "24450490" AndAlso DdlProvisoes.SelectedValue = 1 AndAlso CDate(txtMovimento.Text) < CDate("01/07/2024") Then
                'NÃO FAZ NADA RT GRÃOS IMPLANTAÇÃO SALDO INICIAL - FURLAN - 18/07/2024
            Else
                Mensagem = "Domingo - Data Inválida para " & Tipo & "..."
                Return False
            End If
        End If

        Sql = "  SELECT Descricao" & vbCrLf &
              "    FROM DatasNaoProgramaveis" & vbCrLf &
              "   WHERE Empresa_Id = '99999999999999' " & vbCrLf &
              "     AND EndEmpresa_ID = 0 " & vbCrLf &
              "     AND Data_ID = '" & CDate(pData).ToString("yyyy/MM/dd") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Mensagem = "Data de " & Tipo & " não programável, Feriado Nacional > " & Dr("Descricao")
            Return False
        Next

        If Empresa <> "" Then
            Sql = "  SELECT Descricao " & vbCrLf &
                  "    FROM DatasNaoProgramaveis" & vbCrLf &
                  "   WHERE Empresa_Id = '" & Empresa & "' " & vbCrLf &
                  "     AND EndEmpresa_ID = " & EndEmpresa & " " & vbCrLf &
                  "     AND Data_ID = '" & CDate(pData).ToString("yyyy/MM/dd") & "'" & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                Mensagem = "Data de " & Tipo & " não programável, Feriado Municipal > " & Dr("Descricao")
                Return False
            Next
        End If

        If Tipo = "Movimento" Then
            If Not Funcoes.VerificaAcesso(Empresa, EndEmpresa, pData, "Financeiro") Then
                Mensagem = "Movimento já Fechado para esta data " & pData & ", para empresa " & Empresa
                Return False
            End If
        End If

        Return True

    End Function

    Private Function DolarizaBaixa(ByVal pData As String, ByVal Valor As String, ByVal Indexador As String) As String
        Dim SqlL As String
        Dim Calculo As Decimal

        SqlL = "SELECT Indice" &
               "  FROM Cotacoes" &
               " WHERE Data_Id      ='" & pData & "'" & vbCrLf &
               "   AND Indexador_Id =" & Indexador

        For Each Dr As DataRow In Banco.ConsultaDataSet(SqlL, "Cot").Tables(0).Rows
            Calculo = CDec(Valor) / IIf(Dr("Indice") = 0, 1, Dr("Indice"))
            'Calculo = CDec(FormatNumber(Calculo, 2))
        Next

        Return Calculo.ToString("N2")
    End Function

    Protected Sub DdlEmpresaPagadora_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Try
            'EMPRESA PAGAGORA DEVE SER DO MESMO GRUPO DA EMPRESA DO CLIENTE/FORNECEDOR - FURLAN - 13/05/2025
            If DdlProvisoes.SelectedValue = 1 AndAlso Not Left(DdlEmpresaPagadora.SelectedValue.Split("-")(0), 8) = Left(DdlEmpresaCliente.SelectedValue.Split("-")(0), 8) Then
                If HTabela.Value = "ContasAPagar" Then
                    MsgBox(Me.Page, "Empresa Pagadora deve ser do mesmo grupo da Empresa do Fornecedor.", eTitulo.Info)
                Else
                    MsgBox(Me.Page, "Empresa Recebedora deve ser do mesmo grupo da Empresa do Cliente.", eTitulo.Info)
                End If

                DdlEmpresaPagadora.SelectedIndex = 0
                DdlTiposDePagamentos.SelectedIndex = 0
                DdlBancoPagador.Items.Clear()
                DdlContaPagadora.Items.Clear()

                Exit Sub
            End If

            DdlTiposDePagamentos.SelectedIndex = 0
            CargaBancoPagador()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub cmdFavorecido_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo" & HID.Value) = "LivreClasse"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objFavorecidoCP" & HID.Value, "txtNome")
    End Sub

    Private Function Adiantamento(ByVal Pedido As [Lib].Negocio.Pedido, ByVal RegistroTitulo As String, ByVal DataVencimento As Date) As Boolean
        Dim RegistroAdiantamento As Integer = 0

        Cliente = DdlEmpresaCliente.SelectedValue
        campo = Cliente.Split("-")

        Sql = "exec sp_Numerador '" & campo(0) & "', " & campo(1) & ", 15"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Numerador").Tables(0).Rows
            RegistroAdiantamento = Dr("Sequencia")
        Next

        If RegistroAdiantamento = 0 Then
            Mensagem = "Registro <" & RegistroAdiantamento & "> Numerador 15 Adiantamentos não encontrado..."
            Return False
        End If

        Sql = " INSERT INTO Adiantamentos (Empresa_ID, EndEmpresa_ID, Cliente_ID, EndCliente_ID, Adiantamento_Id, RegistroPedido, Titulo, Safra, Vencimento, Taxa)" & vbCrLf &
              " VALUES('" & campo(0) & "'," & campo(1)

        campo = txtCodigoFornecedor.Value.Split("-")
        Sql &= ", '" & campo(0) & "', " & campo(1) & ", " & RegistroAdiantamento & vbCrLf

        If Pedido Is Nothing OrElse Pedido.Codigo = 0 Then
            Sql &= ", NULL" & vbCrLf
        Else
            Sql &= ", " & Pedido.Codigo & vbCrLf
        End If

        Sql &= ", " & RegistroTitulo & vbCrLf

        If Pedido Is Nothing OrElse Pedido.Codigo = 0 Then
            Sql &= ",'" & ddlSafraAdto.SelectedValue.ToString & "'" & vbCrLf
        Else
            Sql &= ", '" & Pedido.CodigoSafra & "'" & vbCrLf
        End If

        If String.IsNullOrWhiteSpace(DataVencimento) Then
            Throw New Exception("Titulo " & RegistroTitulo & " com carteira de Adiantamento mais sem a data do Vencimento informado")
        Else
            Sql &= ", '" & DataVencimento.ToString("yyyy/MM/dd") & "'"
        End If

        If String.IsNullOrWhiteSpace(txtTaxaAdto.Text) Then
            Sql &= ", 0 )"                                                             'Taxa de Adiantamento 
        Else
            Sql &= ", " & Replace(Replace(txtTaxaAdto.Text, ".", ""), ",", ".") & ")"  'Taxa
        End If

        SqlArray.Add(Sql)

        Return True
    End Function

    Private Function AdiantamentoAmortizacao(ByVal RegistroTitulo As String, ByVal NumeroTituloGeradorDoAdiantamento As Integer, NumeroAdto As String, ByVal ValorBxAdtoOficial As Decimal, ByVal ValorBxAdtoMoeda As Decimal, ByVal ValorVariacao As Decimal) As Boolean
        Dim objAdiantamento As New [Lib].Negocio.Adiantamento(NumeroTituloGeradorDoAdiantamento)

        If objAdiantamento.Codigo = 0 Then
            Mensagem = "Adiantamento informado não foi encontrado, verifique a lista."
            Return False
        End If

        Dim Sequencia As Integer

        Sqla = "  SELECT ISNULL(MAX(Sequencia_Id), 0) + 1 AS Sequencia  " & vbCrLf &
               "    FROM vw_AdiantamentosXBaixas" & vbCrLf &
               "   WHERE TituloAdiantamento =" & NumeroTituloGeradorDoAdiantamento & vbCrLf


        For Each Dr As DataRow In Banco.ConsultaDataSet(Sqla, "Encargos").Tables(0).Rows
            Sequencia = Dr("Sequencia")
        Next

        Sql = " INSERT INTO AdiantamentosXBaixas" &
                    " (Empresa_ID, EndEmpresa_ID, Cliente_ID, EndCliente_ID, Adiantamento_Id, Sequencia_Id, Titulo, ValorOficial, ValorMoeda, VariacaoOficial, DataBaixa)" &
                    " VALUES('" & objAdiantamento.CodigoEmpresa & "'," & objAdiantamento.EndEmpresa

        campo = txtCodigoFornecedor.Value.Split("-")
        Sql &= ", '" & objAdiantamento.CodigoCliente & "', " & objAdiantamento.EndCliente & "," & NumeroAdto & ", " & Sequencia


        Sql &= ", " & RegistroTitulo & ", " & Str(ValorBxAdtoOficial) & ", " & Str(ValorBxAdtoMoeda) & "," & Str(ValorVariacao) & ",'" & CDate(txtMovimento.Text).ToSqlDate & "')"

        SqlArray.Add(Sql)

        Return True
    End Function

    Protected Sub cmdPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdPedido.Click
        If DdlEmpresaCliente.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada!")
        ElseIf txtCodigoFornecedor.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Fornecedor não foi selecionado!")
        ElseIf (txtValorDoDocumento.Text.Length = 0 And txtValorEmMoeda.Text.Length = 0) Or txtValorCobrado.Text.Length = 0 Then
            MsgBox(Me.Page, "Valor do documento em R$ ou U$ e valor líquido para pgto são obrigatórios!")
        ElseIf ddlCarteiras.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Finalidade Financeira é obrigatório!")
        Else
            Session("ssCampo" & HID.Value) = "Pedidos"
            Cliente = DdlEmpresaCliente.SelectedValue
            campo = Cliente.Split("-")
            Dim strCliente() As String = txtCodigoFornecedor.Value.Split("-")
            Dim parameters As New Dictionary(Of String, Object)
            parameters("unidade") = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue
            parameters("empresa") = campo(0)
            parameters("enderecoEmpresa") = campo(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")

            Dim cart As New CarteiraFinanceira(ddlCarteiras.SelectedValue)
            If (Not cart.isAdiantamento Or (cart.isAdiantamento And cart.BaixaAdiantamento)) And Not chkManterLancamento.Checked Then
                'MsgBox(Me.Page, "Só é permitido associar pedido a um titulo se este for adiantamento.")
                'Exit Sub
                If Not chkManterLancamento.Checked Then parameters("ClassePedido") = IIf(HTabela.Value = "ContasAPagar", "'VENDAS'", "'COMPRAS','COMPRASGERAIS'")
            Else
                If Not chkManterLancamento.Checked Then parameters("ClassePedido") = IIf(HTabela.Value = "ContasAPagar", "'COMPRAS','COMPRASGERAIS'", "'VENDAS'")
            End If

            Popup.ConsultaDePedidos(Me.Page, "objPedidoCTAPAG" & HID.Value)
            ucConsultaPedidos.BindGridView(parameters)
        End If
    End Sub

    Private Function PesoPago(ByVal Pedido As String, ByVal Valor As String, ByVal Historico As String) As String
        Dim HistoricoParcial As String = Historico
        Dim Empresa() As String = DdlEmpresaCliente.SelectedValue.ToString.Split("-")

        Sqla = "SELECT pif.Pedido_Id," & vbCrLf &
               "       pif.Produto_Id," & vbCrLf &
               "       Produtos.Nome," & vbCrLf &
               "       Produtos.Agrupar, " & vbCrLf &
               "       pif.Fixacao_Id," & vbCrLf &
               "       pif.Quantidade," & vbCrLf &
               "       isnull(Pedidos.PedidoEfetivo,'') as PedidoEfetivo," & vbCrLf &
               "       ISNULL((SELECT SUM(T1.ValorOficial) AS Oficial " & vbCrLf &
               "                 FROM VW_PedidosXItensXFixacoesXEncargos AS T1 " & vbCrLf &
               "                WHERE T1.Encargo_Id = 'LIQUIDO'" & vbCrLf &
               "                  AND T1.Pedido_Id  = pif.Pedido_Id), 0) AS Oficial," & vbCrLf &
               "       ISNULL((SELECT SUM(T1.ValorMoeda) AS Moeda " & vbCrLf &
               "                 FROM VW_PedidosXItensXFixacoesXEncargos AS T1 " & vbCrLf &
               "                WHERE T1.Encargo_Id = 'LIQUIDO'" & vbCrLf &
               "                  AND T1.Pedido_Id  = pif.Pedido_Id), 0) AS Moeda " & vbCrLf &
               "  FROM VW_PedidosXItensXFixacoes pif" & vbCrLf &
               " INNER JOIN Produtos" & vbCrLf &
               "    ON pif.Produto_Id = Produtos.Produto_Id" & vbCrLf &
               " INNER JOIN Pedidos" & vbCrLf &
               "    ON pif.Empresa_Id    = Pedidos.Empresa_Id" & vbCrLf &
               "   AND pif.EndEmpresa_Id = Pedidos.EndEmpresa_Id" & vbCrLf &
               "   AND pif.Pedido_Id     = Pedidos.Pedido_Id " & vbCrLf &
               " WHERE Pedidos.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf &
               "   AND pif.Pedido_Id      = " & Pedido

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sqla, "Pedido").Tables(0).Rows
            If Dr("Agrupar") = "N" And Dr("Oficial") <> 0 Then
                HistoricoParcial = Left(Historico, 21)
                If ddlMoeda.SelectedValue = 1 Then
                    HistoricoParcial &= Convert.ToDecimal(Dr("Quantidade") * Valor / Dr("Oficial")).ToString("N0") & " KGS DE " & Dr("Nome") & " - PEDIDO " & Pedido & IIf(Dr("PedidoEfetivo").ToString.Length = 0, "", ", CN - " & Dr("PedidoEfetivo"))
                Else
                    HistoricoParcial &= Convert.ToDecimal(Dr("Quantidade") * Valor / Dr("Moeda")).ToString("N0") & " KGS DE " & Dr("Nome") & " - PEDIDO " & Pedido & IIf(Dr("PedidoEfetivo").ToString.Length = 0, "", ", CN - " & Dr("PedidoEfetivo"))
                End If
            End If
        Next
        Return HistoricoParcial
    End Function

    Protected Sub RbAtivo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RbAtivo.CheckedChanged
        chkPrevisao.Visible = True
        chkProvisao.Visible = True
        chkBaixado.Visible = True

        chkPrevisao.Checked = True
        chkProvisao.Checked = False
    End Sub

    Protected Sub RbGeral_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        chkPrevisao.Visible = False
        chkProvisao.Visible = False
        chkBaixado.Visible = False

        chkPrevisao.Checked = False
        chkProvisao.Checked = False
    End Sub

    Protected Sub RbCancelado_CheckedChanged(sender As Object, e As EventArgs)
        chkPrevisao.Visible = False
        chkProvisao.Visible = False
        chkBaixado.Visible = False

        chkPrevisao.Checked = False
        chkProvisao.Checked = False
    End Sub

    Private Sub DiaGeralExcel()
        Try
            Dim ds As DataSet

            Dim objEmpresa As New [Lib].Negocio.Cliente()

            If DdlEmpresaConsultaTitulos.SelectedValue.Length > 0 Then
                objEmpresa = New [Lib].Negocio.Cliente(DdlEmpresaConsultaTitulos.SelectedValue.ToString.Split("-")(0), DdlEmpresaConsultaTitulos.SelectedValue.ToString.Split("-")(1))
            End If

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            ds = getDataSet("")

            'colunas que serão consideradas na criação da dataTable
            Dim collumns() As String = {"Registro", "Pedido", "Cliente", "NomeCliente", "Movimento", "VencimentoOriginal", "Vencimento", "Baixa", "Historico", "ValorLiquido"}

            'criando linhas, passando dados do DataSet para DataTable
            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("sheet1")

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'linha informa cidade e estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "RELATORIO DE TÍTULOS")
                    worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtPeriodoInicialConsultaTitulos.Text)) & " A " & String.Format("{0:dd/MM/yyyy}", CDate(txtPeriodoFinalConsultaTitulos.Text)))
                    worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:J{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    For Each col As DataColumn In ds.Tables(0).Columns
                        For Each include As String In collumns
                            If col.ColumnName = include Then
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                columnIndex += 1
                            End If
                        Next
                    Next

                    worksheet.Cells("B5:H" & rowIndex).AutoFilter = True

                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, collumns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(0).Columns
                            For Each include As String In collumns
                                If col.ColumnName = include Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                    columnIndex += 1
                                End If
                            Next
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

                        worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("F{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

                        worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("G{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

                        worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

                        'formatando células numéricas
                        worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, collumns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If

                        rowIndex += 1
                    Next



                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Formula = String.Format("=SUM(O6:J{0})", rowIndex - 1)
                    worksheet.Cells.AutoFitColumns(0)

                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub DiaGeral(ByVal pdf As Boolean)
        Dim Parametros As String = ""
        Dim crystal As String = ""
        Dim ds As DataSet

        ds = getDataSet(Parametros)

        If String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
            crystal = "Cr_Titulos"
        Else
            crystal = "Cr_TitulosPorEmpresa"
        End If

        Dim parameters = New Dictionary(Of String, Object)()

        If HTabela.Value = "ContasAPagar" Then
            parameters.Add("Relatorio", "Relação de Títulos A Pagar")
        Else
            parameters.Add("Relatorio", "Relação de Títulos A Receber")
        End If

        Dim objEmpresa As [Lib].Negocio.Cliente = New [Lib].Negocio.Cliente(HttpContext.Current.Session("ssEmpresa"), HttpContext.Current.Session("ssEndEmpresa"))
        parameters.Add("EmpresaNome", objEmpresa.Nome)
        parameters.Add("EmpresaCidade", objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado)
        parameters.Add("EmpresaCodigo", Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
        parameters.Add("UnidadeDeNegocio", Parametros)
        parameters.Add("TipoDaCarteira", "Carteira")

        Funcoes.BindReport(Me.Page, ds, crystal, IIf(pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)

    End Sub

    Private Sub FilialGeral()
        Dim Cliente As String
        Dim Campo() As String
        Dim i As Integer = 0
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue
        'Dim dra As DataRow
        Dim sql As String
        Dim lista As String
        Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
        Dim strm As IO.StreamWriter
        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        Dim linha As String
        Dim dataproc As Date
        Dim conterro As Integer
        Dim contreg As Integer
        dataproc = Date.Today
        Dim ds As New DataSet
        'Dim Inconsist As String
        'Dim Sitban As Integer
        Dim Valdia As Decimal
        'Dim ValdiaDolar As Decimal
        Dim Valtot As Decimal
        Dim Valtotdolar As Decimal
        Dim Valtotdia As Decimal
        Dim Valtotdiadolar As Decimal
        Dim Valind As Decimal
        Dim Datvenctr As String
        Dim Empresa As String
        'Dim registro As String
        lista = "Todos"
        sql = "  SELECT CP.Registro_Id AS Registro, " & vbCrLf &
              "         convert(varchar(10),CP.Vencimento,103) as Vencimento,convert(varchar(10)," & vbCrLf &
              "         CP.Prorrogacao,103) as Posterga, " & vbCrLf &
              "         CASE " & vbCrLf &
              "             when ISNULL(ed.Codigo_Id,0) > 0" & vbCrLf &
              "                then CliED.Nome" & vbCrLf &
              "                else Cli.Nome" & vbCrLf &
              "             end AS Cliente, " & vbCrLf &
              "         Historico, " & vbCrLf &
              "         isnull(CP.MoedaValorLiquido, 0) AS Dolar, " & vbCrLf &
              "         CP.ValorLiquido AS Valor, " & vbCrLf &
              "         UsuarioLiberacao as Liberado, " & vbCrLf &
              "         CP.Carteira As Carteira, " & vbCrLf &
              "         ComprasXProdutos.Descricao As DescricaoCarteira, " & vbCrLf &
              "         CP.Situacao as Situacao, " & vbCrLf &
              "         CP.Empresa as Empresa, " & vbCrLf &
              "         Empresa.Reduzido as Reduzido, " & vbCrLf &
              "         CP.UsuarioBaixa as UsuarioBaixa" & vbCrLf &
              "    FROM " & HTabela.Value & " CP " & vbCrLf &
              "   INNER JOIN Clientes Cli " & vbCrLf &
              "      ON CP.Cliente = Cli.Cliente_Id " & vbCrLf &
              "     AND CP.EndCliente = Cli.Endereco_Id" & vbCrLf &
              "    LEFT JOIN ComprasXProdutos " & vbCrLf &
              "      ON CP.Carteira = ComprasXProdutos.Produto_id " & vbCrLf &
              "   INNER JOIN Clientes as Empresa " & vbCrLf &
              "      ON CP.Empresa = Empresa.Cliente_Id " & vbCrLf &
              "     AND CP.EndEmpresa = Empresa.Endereco_Id" & vbCrLf &
              "  LEFT JOIN EndossoXTitulo eTxT" & vbCrLf &
              "    ON eTxT.Empresa_Id    = CP.Empresa" & vbCrLf &
              "   AND eTxT.EndEmpresa_Id = CP.EndEmpresa" & vbCrLf &
              "   AND eTxT.Titulo_Id     = CP.Registro_Id" & vbCrLf &
              "  LEFT JOIN Endosso ed" & vbCrLf &
              "    ON ed.Empresa_Id    = eTxT.Empresa_Id" & vbCrLf &
              "   AND ed.EndEmpresa_Id = eTxT.EndEmpresa_Id" & vbCrLf &
              "   AND ed.Codigo_Id     = eTxT.Codigo_Id" & vbCrLf &
              "  LEFT JOIN Clientes CliED" & vbCrLf &
              "    ON CliED.Cliente_Id  = ed.ClienteEndosso_Id" & vbCrLf &
              "   AND CliED.Endereco_Id = ed.EndClienteEndosso_Id" & vbCrLf &
              "   WHERE CP.Provisao <> 1 " & vbCrLf

        If Funcoes.VerificaPermissao("TituloDeFuncionario", "LEITURA") Then
            sql &= "   AND CP.Situacao in(1,101,102,105) " & vbCrLf
        Else
            sql &= "   AND CP.Situacao in(1,101,102) " & vbCrLf
        End If

        If RbGeral.Checked = True Then
            '' Nao ira fazer nada pois vai listar todos 
            lista = "Todos"
        End If

        If RbAtivo.Checked = True Then
            sql &= " and CP.usuariobaixa = '' " & vbCrLf
            lista = "Ativos"
        End If

        If chkBaixado.Checked = True Then
            sql &= " and CP.usuariobaixa <> '' " & vbCrLf
            lista = "Baixados"
        End If

        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            sql &= " and CP.UnidadeDeNegocio = '" & Cliente & "' " & vbCrLf
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            sql &= " and CP.Empresa = '" & Campo(0) & "'" & vbCrLf &
                   " and CP.EndEmpresa = " & Campo(1) & vbCrLf
        End If

        'Campo = txtCodigoClienteConsulta.Value.Split("-")
        'If Campo(0) <> "" Then
        '    sql &= " and CP.Cliente = '" & Campo(0) & "'" & vbCrLf &
        '           " and CP.EndCliente = " & Campo(1) & vbCrLf
        'End If
        If Not String.IsNullOrWhiteSpace(txtCodigoClienteConsultaEndosso.Value) Then
            sql &= "   AND ed.ClienteEndosso_Id    = '" & txtCodigoClienteConsultaEndosso.Value.Split("-")(0) & "'" & vbCrLf &
                   "   AND ed.EndClienteEndosso_Id = " & txtCodigoClienteConsultaEndosso.Value.Split("-")(1) & vbCrLf
        ElseIf Not String.IsNullOrWhiteSpace(txtCodigoClienteConsulta.Value) Then
            Campo = txtCodigoClienteConsulta.Value.Split("-")
            sql &= " and CP.Cliente = '" & Campo(0) & "'" & vbCrLf &
                   " and CP.EndCliente = " & Campo(1) & vbCrLf
        End If

        If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
            sql &= " and Prorrogacao between '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf

        End If

        If ChkAutozizado.Checked = True Then
            sql &= " and CP.UsuarioLiberacao <> ''" & vbCrLf
        End If

        sql &= " ORDER BY CP.Empresa, CP.Prorrogacao " & vbCrLf

        dsRelatorio = Banco.ConsultaDataSet(sql, "Contas")
        linha = "<HTML>" & vbCrLf
        ''<HEAD>
        linha &= "<HEAD>" & vbCrLf
        linha &= "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf
        linha &= "<TITLE> Posicao de titulos processados em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TITLE>" & vbCrLf
        linha &= "</HEAD>" & vbCrLf
        '<BODY>
        linha &= "<BODY>" & vbCrLf
        '-----------------
        'Cabeçalho Padrao
        '-----------------
        linha &= "<TR>"
        linha &= "<TD >" & "Posicao de titulos em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TD>"
        linha &= "</TR>"
        ''**************************** Controle de cabecalho
        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Unidade : " & DdlUnidadeConsultaTitulos.SelectedItem.Text & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Unidade : " & "Todas as Unidades !!!" & " </TD>"
            linha &= "</TR>"
        End If
        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Empresa : " & DdlEmpresaConsultaTitulos.SelectedItem.Text & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Empresa : " & "Todas as Empresas !!!" & " </TD>"
            linha &= "</TR>"
        End If

        Campo = txtCodigoClienteConsulta.Value.Split("-")
        If Campo(0) <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Cliente : " & txtCodigoClienteConsulta.Value & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Cliente : " & "Todos os Clientes" & " </TD>"
            linha &= "</TR>"
        End If
        ''**************************** Controle de cabecalho fim 
        linha &= "<TR>"
        linha &= "<TD >" & " Tipo de relatorio Filial(Empresa)Geral - Totalizacao por Filial - Registros impressos : " & lista & " </TD>"
        linha &= "</TR>"
        linha &= "<table width= '370' cellpadding='0' cellspacing='0' Border=0>"
        linha &= "<TR>"
        linha &= "<TD><B>Registro</B></TD>"
        linha &= "<TD><B>Cliente/Fornecedor</B></TD>"
        linha &= "<TD><B>Historico</B></TD>"
        linha &= "<TD><B>Pagar R$</B></TD>"
        linha &= "<TD><B>Pagar US$</B></TD>"
        linha &= "<TD><B>Empresa</B></TD>"
        linha &= "<TD><B>Vencimento Original</B></TD>"
        linha &= "<TD><B>Prorrogacao</B></TD>"
        linha &= "<TD><B>Carteira</B></TD>"
        linha &= "<TD><B>Situacao</B></TD>"
        linha &= "</TR>"
        ''
        conterro = 0
        contreg = 0
        Valind = 0
        Valtot = 0
        Valdia = 0
        Datvenctr = ""
        Empresa = ""
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                contreg = contreg + 1
                If Empresa <> "" And Empresa <> dr("Empresa") Then
                    linha &= "<TR>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & " Total Filial: " & "</B></TD>"
                    linha &= "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>"
                    linha &= "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "</TR>"
                    Valtotdia = 0
                    Valtotdiadolar = 0
                End If
                Valtot = Valtot + dr("Valor")
                Valtotdolar = Valtotdolar + dr("Dolar")
                Valtotdia = Valtotdia + dr("Valor")
                Valtotdiadolar = Valtotdiadolar + dr("Dolar")
                Empresa = dr("Empresa")
                Datvenctr = dr("Vencimento")
                linha &= "<TR>"
                linha &= "<TD>" & dr("Registro") & "</TD>"
                linha &= "<TD>" & dr("Cliente") & "</TD>"
                linha &= "<TD>" & dr("Historico") & "</TD>"
                linha &= "<TD>" & Convert.ToDouble(dr("Valor")).ToString("n2") & "</TD>"
                linha &= "<TD>" & Convert.ToDouble(dr("Dolar")).ToString("n2") & "</TD>"
                linha &= "<TD>" & dr("Reduzido") & "</TD>"
                linha &= "<TD>" & dr("Vencimento") & "</TD>"
                linha &= "<TD>" & dr("Posterga") & "</TD>"
                linha &= "<TD>" & dr("DescricaoCarteira") & "</TD>"
                If dr("UsuarioBaixa") = "" Then
                    linha &= "<TD>" & "Ativo" & "</TD>"
                Else
                    linha &= "<TD>" & "Baixado" & "</TD>"
                End If
                linha &= "</TR>"
            Next
        End If
        linha &= "<TR>>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & " Total Filial: " & "</B></TD>"
        linha &= "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "</TR>"
        linha &= "<TR>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & " Total neste processamento: " & "</B></TD>"
        linha &= "<TD><B>" & Valtot.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & Valtotdolar.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "</TR>"
        If contreg = 0 Then
            MsgBox(Me.Page, "Nao existem registros corretos para o periodo.", eTitulo.Info)
            txtPeriodoFinalConsultaTitulos.Focus()
        Else
            MsgBox(Me.Page, "Movimento com registros processados.", eTitulo.Info)
            txtPeriodoFinalConsultaTitulos.Focus()
        End If

        strm = New IO.StreamWriter(arquivo, True)
        Try
            strm.WriteLine(linha)
            strm.Close()
            Funcoes.AbrirArquivo(Me.Page, NomeArquivo)
        Finally
            strm.Close()
        End Try
        '' rotina utilizada nas procuracoes fim - html 
        '' rotina de geracao do relatorio . 
    End Sub

    Private Sub CarteiraDia()
        Dim Cliente As String
        Dim Campo() As String
        Dim i As Integer = 0
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue
        'Dim dra As DataRow
        Dim sql As String
        Dim lista As String
        Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
        Dim strm As IO.StreamWriter
        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        Dim linha As String
        Dim dataproc As Date
        Dim conterro As Integer
        Dim contreg As Integer
        dataproc = Date.Today
        Dim ds As New DataSet
        'Dim Inconsist As String
        'Dim Sitban As Integer
        Dim Valdia As Decimal
        'Dim ValdiaDolar As Decimal
        Dim Valtot As Decimal
        Dim Valtotdolar As Decimal
        Dim Valtotdia As Decimal
        Dim Valtotdiadolar As Decimal
        Dim Valind As Decimal
        Dim Datvenctr As String
        Dim Empresa As String
        Dim carteira As String
        'Dim registro As String
        lista = "Geral"
        sql = " SELECT CP.Registro_Id AS Registro, " & vbCrLf &
              "        convert(varchar(10),CP.Vencimento,103) as Vencimento,convert(varchar(10)," & vbCrLf &
              "        CP.Prorrogacao,103) as Posterga, " & vbCrLf &
              "         CASE " & vbCrLf &
              "             when ISNULL(ed.Codigo_Id,0) > 0" & vbCrLf &
              "                then CliED.Nome" & vbCrLf &
              "                else Clientes.Nome" & vbCrLf &
              "             end AS Cliente, " & vbCrLf &
              "        Historico, " & vbCrLf &
              "        isnull(CP.MoedaValorLiquido, 0) AS Dolar, " & vbCrLf &
              "        CP.ValorLiquido AS Valor, " & vbCrLf &
              "        UsuarioLiberacao as Liberado, " & vbCrLf &
              "        CP.Carteira As Carteira, " & vbCrLf &
              "        ComprasXProdutos.Descricao As DescricaoCarteira, " & vbCrLf &
              "        CP.Situacao as Situacao, " & vbCrLf &
              "        CP.Empresa as Empresa, " & vbCrLf &
              "        Empresa.Reduzido as Reduzido, " & vbCrLf &
              "        CP.UsuarioBaixa as UsuarioBaixa" & vbCrLf &
              "   FROM " & HTabela.Value & " CP " & vbCrLf &
              "  INNER JOIN Clientes " & vbCrLf &
              "     ON CP.Cliente = Clientes.Cliente_Id " & vbCrLf &
              "    AND CP.EndCliente = Clientes.Endereco_Id" & vbCrLf &
              "   LEFT JOIN ComprasXProdutos " & vbCrLf &
              "     ON CP.Carteira = ComprasXProdutos.Produto_id " & vbCrLf &
              "  INNER JOIN Clientes as Empresa " & vbCrLf &
              "     ON CP.Empresa = Empresa.Cliente_Id" & vbCrLf &
              "    AND CP.EndEmpresa = Empresa.Endereco_Id" & vbCrLf &
              "  LEFT JOIN EndossoXTitulo eTxT" & vbCrLf &
              "    ON eTxT.Empresa_Id    = CP.Empresa" & vbCrLf &
              "   AND eTxT.EndEmpresa_Id = CP.EndEmpresa" & vbCrLf &
              "   AND eTxT.Titulo_Id     = CP.Registro_Id" & vbCrLf &
              "  LEFT JOIN Endosso ed" & vbCrLf &
              "    ON ed.Empresa_Id    = eTxT.Empresa_Id" & vbCrLf &
              "   AND ed.EndEmpresa_Id = eTxT.EndEmpresa_Id" & vbCrLf &
              "   AND ed.Codigo_Id     = eTxT.Codigo_Id" & vbCrLf &
              "  LEFT JOIN Clientes CliED" & vbCrLf &
              "    ON CliED.Cliente_Id  = ed.ClienteEndosso_Id" & vbCrLf &
              "   AND CliED.Endereco_Id = ed.EndClienteEndosso_Id" & vbCrLf &
              "  WHERE CP.Provisao <> 1 " & vbCrLf

        If Funcoes.VerificaPermissao("TituloDeFuncionario", "LEITURA") Then
            sql &= "   AND CP.Situacao in(1,101,102,105) " & vbCrLf
        Else
            sql &= "   AND CP.Situacao in(1,101,102) " & vbCrLf
        End If

        If RbGeral.Checked = True Then
            '' Nao ira fazer nada pois vai listar todos 
            lista = "Todos"
        End If

        If RbAtivo.Checked = True Then
            sql &= " and CP.usuariobaixa = '' " & vbCrLf
            lista = "Ativos"
        End If

        If chkBaixado.Checked = True Then
            sql &= " and CP.usuariobaixa <> '' " & vbCrLf
            lista = "Baixados"
        End If

        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            sql &= " and CP.UnidadeDeNegocio = '" & Cliente & "' " & vbCrLf
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            sql &= " and CP.Empresa = '" & Campo(0) & "'" & vbCrLf &
                   " and CP.EndEmpresa = " & Campo(1) & vbCrLf
        End If

        'Campo = txtCodigoClienteConsulta.Value.Split("-")
        'If Campo(0) <> "" Then
        '    sql &= " and CP.Cliente = '" & Campo(0) & "'" & vbCrLf &
        '           " and CP.EndCliente = " & Campo(1) & vbCrLf
        'End If
        If Not String.IsNullOrWhiteSpace(txtCodigoClienteConsultaEndosso.Value) Then
            sql &= "   AND ed.ClienteEndosso_Id    = '" & txtCodigoClienteConsultaEndosso.Value.Split("-")(0) & "'" & vbCrLf &
                   "   AND ed.EndClienteEndosso_Id = " & txtCodigoClienteConsultaEndosso.Value.Split("-")(1) & vbCrLf
        ElseIf Not String.IsNullOrWhiteSpace(txtCodigoClienteConsulta.Value) Then
            Campo = txtCodigoClienteConsulta.Value.Split("-")
            sql &= " and CP.Cliente = '" & Campo(0) & "'" & vbCrLf &
                   " and CP.EndCliente = " & Campo(1) & vbCrLf
        End If

        If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
            sql &= " and Prorrogacao between '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf

        End If

        If ChkAutozizado.Checked = True Then
            sql &= " and CP.UsuarioLiberacao <> ''" & vbCrLf   'Autorizados
        End If

        sql &= " ORDER BY CP.Carteira, CP.Prorrogacao " & vbCrLf


        dsRelatorio = Banco.ConsultaDataSet(sql, "Contas")
        linha = "<HTML>" & vbCrLf
        ''<HEAD>
        linha &= "<HEAD>" & vbCrLf
        linha &= "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf
        linha &= "<TITLE> Posicao de titulos processados em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TITLE>" & vbCrLf
        linha &= "</HEAD>" & vbCrLf
        '<BODY>
        linha &= "<BODY>" & vbCrLf
        '-----------------
        'Cabeçalho Padrao
        '-----------------
        linha &= "<TR>"
        linha &= "<TD >" & "Posicao de titulos em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TD>"
        linha &= "</TR>"
        ''**************************** Controle de cabecalho
        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Unidade : " & DdlUnidadeConsultaTitulos.SelectedItem.Text & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Unidade : " & "Todas as Unidades !!!" & " </TD>"
            linha &= "</TR>"
        End If
        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Empresa : " & DdlEmpresaConsultaTitulos.SelectedItem.Text & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Empresa : " & "Todas as Empresas !!!" & " </TD>"
            linha &= "</TR>"
        End If

        Campo = txtCodigoClienteConsulta.Value.Split("-")
        If Campo(0) <> "" Then
            linha &= "<TR>"
            linha &= "<TD >" & " Cliente : " & txtCodigoClienteConsulta.Value & " </TD>"
            linha &= "</TR>"
        Else
            linha &= "<TR>"
            linha &= "<TD >" & " Cliente : " & "Todos os Clientes" & " </TD>"
            linha &= "</TR>"
        End If
        ''**************************** Controle de cabecalho fim 
        linha &= "<TR>"
        linha &= "<TD >" & " Tipo de relatorio Filial(Empresa)Geral - Totalizacao por Filial - Registros impressos : " & lista & " </TD>"
        linha &= "</TR>"
        linha &= "<table width= '370' cellpadding='0' cellspacing='0' Border=0>"
        linha &= "<TR>"
        linha &= "<TD ><B>Registro</B></TD>"
        linha &= "<TD ><B>Cliente/Fornecedor</B></TD>"
        linha &= "<TD ><B>Historico</B></TD>"
        linha &= "<TD ><B>Pagar R$</B></TD>"
        linha &= "<TD ><B>Pagar US$</B></TD>"
        linha &= "<TD ><B>Empresa</B></TD>"
        linha &= "<TD ><B>Vencimento Original</B></TD>"
        linha &= "<TD ><B>Prorrogacao</B></TD>"
        linha &= "<TD ><B>Carteira<B></TD>"
        linha &= "<TD ><B>Situacao</B></TD>"
        linha &= "</TR>"
        ''
        conterro = 0
        contreg = 0
        Valind = 0
        Valtot = 0
        Valdia = 0
        Datvenctr = ""
        Empresa = ""
        carteira = ""
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                contreg = contreg + 1
                If carteira <> "" And carteira <> dr("DescricaoCarteira") Then
                    linha &= "<TR>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & " Total Da Carteira: " & "</B></TD>"
                    linha &= "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>"
                    linha &= "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "<TD><B>" & "." & "</B></TD>"
                    linha &= "</TR>"
                    Valtotdia = 0
                    Valtotdiadolar = 0
                End If
                Valtot = Valtot + dr("Valor")
                Valtotdolar = Valtotdolar + dr("Dolar")
                Valtotdia = Valtotdia + dr("Valor")
                Valtotdiadolar = Valtotdiadolar + dr("Dolar")
                Empresa = dr("Empresa")
                Datvenctr = dr("Vencimento")
                carteira = dr("DescricaoCarteira")
                linha &= "<TR>"
                linha &= "<TD>" & dr("Registro") & "</TD>"
                linha &= "<TD>" & dr("Cliente") & "</TD>"
                linha &= "<TD>" & dr("Historico") & "</TD>"
                linha &= "<TD>" & Convert.ToDouble(dr("Valor")).ToString("n2") & "</TD>"
                linha &= "<TD>" & Convert.ToDouble(dr("Dolar")).ToString("n2") & "</TD>"
                linha &= "<TD>" & dr("Reduzido") & "</TD>"
                linha &= "<TD>" & dr("Vencimento") & "</TD>"
                linha &= "<TD>" & dr("Posterga") & "</TD>"
                linha &= "<TD>" & dr("Carteira") & " - " & dr("DescricaoCarteira") & "</TD>"
                If dr("UsuarioBaixa") = "" Then
                    linha &= "<TD>" & "Ativo" & "</TD>"
                Else
                    linha &= "<TD>" & "Baixado" & "</TD>"
                End If
                linha &= "</TR>"
            Next
        End If
        linha &= "<TR>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & " Total Da Carteira: " & "</B></TD>"
        linha &= "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "</TR>"
        linha &= "<TR>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & " Total neste processamento: " & "</B></TD>"
        linha &= "<TD><B>" & Valtot.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & Valtotdolar.ToString("n2") & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "<TD><B>" & "." & "</B></TD>"
        linha &= "</TR>"
        If contreg = 0 Then
            MsgBox(Me.Page, "Nao existem registros corretos para o periodo.", eTitulo.Info)
            txtPeriodoFinalConsultaTitulos.Focus()
        Else
            MsgBox(Me.Page, "Movimento com registros processados.", eTitulo.Info)
            txtPeriodoFinalConsultaTitulos.Focus()
        End If

        strm = New IO.StreamWriter(arquivo, True)
        Try
            strm.WriteLine(linha)
            strm.Close()
            Funcoes.AbrirArquivo(Me.Page, NomeArquivo)
        Finally
            strm.Close()
        End Try
        '' rotina utilizada nas procuracoes fim - html 
        '' rotina de geracao do relatorio . 
    End Sub

    Private Sub EmitirRecibo()

        Try
            Dim DataBaixa As String = ""
            Dim TnumeroDoCheque As Integer

            Sql = "SELECT Baixa, ISNULL(NumeroDoCheque, 0) AS NumeroDoCheque" & vbCrLf &
                  "  FROM " & HTabela.Value & " CP " & vbCrLf &
                  " WHERE Registro_Id = " & txtRegistro.Text & vbCrLf

            For Each Dr1 As DataRow In Banco.ConsultaDataSet(Sql, "Titulo").Tables(0).Rows
                DataBaixa = IIf(Dr1("Baixa") Is Nothing, "", CDate(Dr1("Baixa")).ToString("dd/MM/yyyy"))
                TnumeroDoCheque = Dr1("NumeroDoCheque")
            Next

            '' Rotina para leitura de registro buscando a data da baixa (fim).
            '' Dados da empresa
            Dim Campo As String() = Nothing
            Dim Empresa As String = ""
            Dim EndEmpresa As String = ""
            Dim ENome As String = ""
            Dim EEndereco As String = ""
            Dim ECep As String = ""
            Dim ECidade As String = ""
            Dim EEstado As String = ""
            Dim ECnpj As String = ""
            Dim EInscricao As String = ""
            Dim Efone As String = ""
            Dim EBairro As String = ""
            Dim EComplemento As String = ""
            Dim ENumero As Integer

            'Dim GrupoEmpresa As String
            If DdlEmpresaCliente.Text <> "" Then
                Empresa = DdlEmpresaCliente.SelectedValue
                Campo = Empresa.Split("-")
                Empresa = Campo(0)                      'EmpresaCliente
                EndEmpresa = Campo(1)                   'Endereco Empresa Cliente
            Else
                Empresa = ""                            'Empresa Cliente
                EndEmpresa = 0                          'Endereco Empresa Cliente
            End If

            Dim dr As DataRow
            '' Dados da empresa - fim 
            '' Consultado empresa 
            Sql = "SELECT Emp.Cliente_Id ," & vbCrLf &
                  "       Emp.Nome, Emp.Cidade," & vbCrLf &
                  "       Emp.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf &
                  "       Emp.Endereco , Emp.Cep," & vbCrLf &
                  "       Emp.Inscricao, Emp.Telefone," & vbCrLf &
                  "       Emp.Bairro, Emp.Complemento," & vbCrLf &
                  "       Emp.Numero " & vbCrLf &
                  "  FROM Clientes Emp " & vbCrLf &
                  " WHERE Emp.Cliente_Id = '" & Empresa & "'" & vbCrLf

            DS = Banco.ConsultaDataSet(Sql, "Clientes")
            If DS.Tables(0).Rows.Count > 0 Then
                For Each dr In DS.Tables(0).Rows
                    ENome = dr("Nome")
                    EEndereco = dr("Endereco")
                    ECep = dr("Cep")
                    ECidade = dr("Cidade")
                    EEstado = dr("Estado")
                    ECnpj = dr("Cliente_id")
                    EInscricao = dr("Inscricao")
                    Efone = dr("Telefone")
                    EBairro = dr("Bairro")
                    EComplemento = dr("Complemento")
                    ENumero = dr("Numero")
                    Exit For
                Next
            End If
            '' Consultando Empresa - fim 
            ''**************************************************************
            '' Dados do cliente
            Dim Cliente As String = ""
            Dim EndCliente As String = ""
            Dim CNome As String = ""
            Dim CEndereco As String = ""
            Dim CCep As String = ""
            Dim CCidade As String = ""
            Dim CEstado As String = ""
            Dim CCnpj As String = ""
            Dim CInscricao As String = ""
            Dim Cfone As String = ""
            Dim CBairro As String = ""
            Dim CComplemento As String = ""
            Dim CNumero As Integer
            'Dim GrupoEmpresa As String
            If txtCodigoFornecedor.Value.ToString.Length > 0 Then
                Campo = txtCodigoFornecedor.Value.Split("-")
                Cliente = Campo(0)                      'Cliente
                EndCliente = Campo(1)                   'Endereco Cliente
            Else
                Cliente = ""                            'Cliente
                EndCliente = 0                          'Endereco Cliente
            End If
            '' Dados do Cliente - fim 
            '' Consultado Cliente 
            Sql = " SELECT Cli.Cliente_Id ," & vbCrLf &
                  "        Cli.Nome, Cli.Cidade," & vbCrLf &
                  "        Cli.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf &
                  "        Cli.Endereco , Cli.Cep," & vbCrLf &
                  "        Cli.Inscricao, Cli.Telefone," & vbCrLf &
                  "        Cli.Bairro, Cli.Complemento," & vbCrLf &
                  "        Cli.Numero " & vbCrLf &
                  "   FROM Clientes Cli " & vbCrLf &
                  "  WHERE Cli.Cliente_Id = '" & Cliente & "'" & vbCrLf &
                  "    AND Cli.Endereco_id = " & EndCliente & vbCrLf
            DS = Banco.ConsultaDataSet(Sql, "Clientes")
            If DS.Tables(0).Rows.Count > 0 Then
                For Each dr In DS.Tables(0).Rows
                    CNome = dr("Nome")
                    CEndereco = dr("Endereco")
                    CCep = dr("Cep")
                    CCidade = dr("Cidade")
                    CEstado = dr("Estado")
                    CCnpj = dr("Cliente_id")
                    CInscricao = dr("Inscricao")
                    Cfone = dr("Telefone")
                    CBairro = dr("Bairro")
                    CComplemento = dr("Complemento")
                    CNumero = dr("Numero")
                    Exit For
                Next
            End If
            'Dim myDataRow As DataRow
            '' Cria data Set que vai ser utilizado no relatorio 
            Dim xextenso As String
            Dim yextenso As String
            Dim mes As Integer = 0
            Dim j As Integer = 0
            Dim k As Integer = 0
            Dim Mesx As String = ""
            Dim dsRecibo As New DataSet
            Dim dtRecibo As DataTable
            Dim row As DataRow
            Dim RegistroI As String = ""
            Dim RegistroS As String = ""
            dtRecibo = New DataTable("ReciboAPagar")
            '' campos da empresa 
            dtRecibo.Columns.Add("ENome", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EEndereco", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ECep", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ECidade", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EEstado", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ECnpj", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EInscricao", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EFone", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EBairro", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EComplemento", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ENumero", Type.GetType("System.Int32")).DefaultValue = 0
            ' ''campos do cliente
            dtRecibo.Columns.Add("CNome", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CEndereco", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CCep", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CCidade", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CEstado", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CCnpj", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CInscricao", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CFone", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CBairro", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CComplemento", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CNumero", Type.GetType("System.Int32")).DefaultValue = 0
            '' campos to titulo 
            dtRecibo.Columns.Add("TNumtit", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TValor", Type.GetType("System.Decimal"))
            dtRecibo.Columns.Add("TExtenso", Type.GetType("System.String"))
            dtRecibo.Columns.Add("THistorico", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TDia", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TMes", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TAno", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TFormaPagto", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TBanco", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TAgencia", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TConta", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TVencimento", Type.GetType("System.DateTime"))
            dtRecibo.Columns.Add("TBaixa", Type.GetType("System.DateTime"))
            dtRecibo.Columns.Add("TRecibo", Type.GetType("System.Int32")).DefaultValue = 0
            dtRecibo.Columns.Add("TRegistro", Type.GetType("System.Int32")).DefaultValue = 0
            dtRecibo.Columns.Add("TNumeroDoCheque", Type.GetType("System.Int32")).DefaultValue = 0
            dtRecibo.Columns.Add("TValorDoDocumento", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TDescontos", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TDeducoes", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TJuros", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TAcrescimos", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TDigito", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TDigitoAgencia", Type.GetType("System.String"))
            dsRecibo.Tables.Add(dtRecibo)
            ' Cria Data Set que vai ser utilizado no relatorio
            ' Move campos para o Data Set.
            ' Move campos da Empresa
            row = dtRecibo.NewRow()
            row("ENome") = ENome
            row("EEndereco") = EEndereco
            row("ECep") = ECep
            row("ECidade") = ECidade
            row("EEstado") = EEstado
            row("ECnpj") = Funcoes.FormatarCpfCnpj(ECnpj)
            row("EInscricao") = EInscricao
            row("EFone") = Efone
            row("EBairro") = EBairro
            row("EComplemento") = EComplemento
            row("ENumero") = ENumero
            '' Move campos do Cliente / fornecedor
            row("CNome") = CNome
            row("CEndereco") = CEndereco
            row("CCep") = CCep
            row("CCidade") = CCidade
            row("CEstado") = CEstado
            row("CCnpj") = Funcoes.FormatarCpfCnpj(CCnpj)
            row("CInscricao") = CInscricao
            row("CFone") = Cfone
            row("CBairro") = CBairro
            row("CComplemento") = CComplemento
            row("CNumero") = CNumero
            '' Move campos do Titulo
            row("Tnumtit") = txtRegistro.Text
            row("TValor") = txtValorCobrado.Text
            row("TNumeroDoCheque") = TnumeroDoCheque
            ''* Rotina de extenso inicio
            yextenso = "("
            yextenso &= UCase(Funcoes.Extenso(txtValorCobrado.Text(), "Real", "Reais"))
            yextenso &= " *"
            xextenso = yextenso
            For j = 1 To (120 - Len(xextenso))
                xextenso &= " *"
            Next
            xextenso &= ")"
            row("TExtenso") = xextenso
            ''* Rotina de extenso fim 
            If txtObservacoes.Text.Length > 0 Then
                row("THistorico") = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim) & ". " & Funcoes.EliminarCaracteresEspeciais(txtObservacoes.Text.Trim) & ". "
            Else
                row("THistorico") = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim)
            End If
            row("TDia") = Day(DataBaixa)
            row("TAno") = Year(DataBaixa)
            row("TMes") = Month(DataBaixa)
            ''* Rotina do Mes inicio

            mes = Month(DataBaixa)
            If mes = 1 Then
                Mesx = "JANEIRO"
            End If

            If mes = 2 Then
                Mesx = "FEVEREIRO"
            End If

            If mes = 3 Then
                Mesx = "MARCO"
            End If

            If mes = 4 Then
                Mesx = "ABRIL"
            End If

            If mes = 5 Then
                Mesx = "MAIO"
            End If

            If mes = 6 Then
                Mesx = "JUNHO"
            End If

            If mes = 7 Then
                Mesx = "JULHO"
            End If

            If mes = 8 Then
                Mesx = "AGOSTO"
            End If

            If mes = 9 Then
                Mesx = "SETEMBRO"
            End If

            If mes = 10 Then
                Mesx = "OUTUBRO"
            End If

            If mes = 11 Then
                Mesx = "NOVEMBRO"
            End If

            If mes = 12 Then
                Mesx = "DEZEMBRO"
            End If

            row("TMes") = Mesx
            ''* rotina do mes fim 
            row("TFormaPagto") = DdlTiposDePagamentos.SelectedItem
            '#contabancaria
            '#03062016
            If lblBanco.Text = "Banco" Then
                row("TBanco") = ""
                row("TAgencia") = ""
                row("TConta") = ""
            Else
                row("TBanco") = lblBanco.Text
                row("TAgencia") = lblAgencia.Text
                row("TConta") = lblContaCorrente.Text
            End If

            row("TVencimento") = txtMovimento.Text
            row("TBaixa") = DataBaixa

            '' calculo de registro . 
            RegistroI = txtRegistro.Text()
            RegistroS = " "
            RegistroS = Trim(RegistroS)
            If Len(RegistroS) < 6 Then
                For k = 1 To (6 - Len(Trim(RegistroI)))
                    RegistroS &= "0"
                Next
            End If
            row("Trecibo") = Trim(RegistroS) & (Trim(RegistroI))
            row("TRegistro") = Trim(RegistroS) & (Trim(RegistroI))
            row("TValorDoDocumento") = txtValorDoDocumento.Text
            row("TDescontos") = txtDescontos.Text
            row("TDeducoes") = txtDeducoes.Text
            row("TJuros") = txtJuros.Text
            row("TAcrescimos") = txtAcrescimos.Text
            row("TDigito") = ""
            row("TDigitoAgencia") = ""

            ''tvalordodocumento
            dtRecibo.Rows.Add(row)

            Dim param As New Dictionary(Of String, Object)
            param.Add("XNome", ENome)

            Funcoes.BindReport(Me.Page, dsRecibo, IIf(HTabela.Value = "ContasAPagar", "Cr_ReciboPagar", "Cr_ReciboReceber"), eExportType.PDF, param)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub cmdAdiantamento_Click(sender As Object, e As EventArgs) Handles cmdAdiantamento.Click
        If DdlEmpresaCliente.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada!")
        ElseIf txtCodigoFornecedor.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Fornecedor não foi selecionado!")
        ElseIf txtCodigoFornecedor.Value.ToString.Length > 0 Then
            SessaoRecuperaTitulo()
            Dim Parametros As New Hashtable
            Parametros.Add("Titulo", IIf(Not IsNumeric(txtRegistro.Text), 0, txtRegistro.Text))

            Cliente = DdlEmpresaCliente.SelectedValue
            campo = Cliente.Split("-")
            Parametros.Add("Empresa", campo(0))
            Parametros.Add("EndEmpresa", campo(1))
            campo = txtCodigoFornecedor.Value.Split("-")
            Parametros.Add("Cliente", campo(0))
            Parametros.Add("EndCliente", campo(1))

            Parametros.Add("Pedido", txtPedido.Text)
            Parametros.Add("PedidoCliente", txtPedido.Text + " - " + txtFornecedor.Text)

            Dim cart As CarteiraFinanceira
            If ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento Then
                Parametros.Add("ContaContabil", ObjTitulo.ObjContaContabilPagadora.Conta)
                Parametros.Add("ContaContabilDescricao", ObjTitulo.ObjContaContabilPagadora.Conta + " - " + ObjTitulo.ObjContaContabilPagadora.Titulo)
            Else
                cart = New CarteiraFinanceira(ddlCarteiras.SelectedValue)
                Parametros.Add("ContaContabil", cart.CodigoContaCliente)
                Parametros.Add("ContaContabilDescricao", cart.CodigoContaCliente + " - " + cart.Descricao)
            End If

            Parametros.Add("Moeda", ddlMoeda.SelectedValue)
            Parametros.Add("DescMoeda", ddlMoeda.SelectedItem.Text)
            Parametros.Add("Formulario", "Financeiro")

            Session("Parametros" & HID.Value) = Parametros
            ucConsultaAdiantamentos.BindGridView()
            Popup.ConsultaDeAdiantamentos(Me.Page, "Financeiro" & HID.Value)
        End If
    End Sub

    Protected Sub imgBloqueio_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgBloqueio.Click
        Cliente = DdlEmpresaCliente.SelectedValue
        campo = Cliente.Split("-")

        If Not Funcoes.VerificaPermissao(HTabela.Value, "LIBERAR") Then
            MsgBox(Me.Page, "Usuário sem permissão para liberar Registro")
        ElseIf Not Funcoes.VerificaAcesso(campo(0), campo(1), txtMovimento.Text, "FINANCEIRO") Then
            MsgBox(Me.Page, "Movimento já Fechado para esta data " & txtMovimento.Text & ", para empresa " & campo(0))
        Else
            SessaoRecuperaTitulo()
            Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)

            If txtNumeroAdto.Text.Length > 0 Then
                If objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento Then
                    Dim objAdiantamento As New [Lib].Negocio.Adiantamento()
                    objAdiantamento.CodigoEmpresa = campo(0)
                    objAdiantamento.EndEmpresa = campo(1)
                    Cliente = txtCodigoFornecedor.Value
                    campo = Cliente.Split("-")
                    objAdiantamento.CodigoCliente = campo(0)
                    objAdiantamento.EndCliente = campo(1)
                    objAdiantamento.Codigo = CInt(txtNumeroAdto.Text)

                    Dim objAdtoXBaixa As New [Lib].Negocio.ListAdiantamentoBaixa(objAdiantamento)

                    If objAdtoXBaixa.Count > 0 Then
                        lnkNovo.Parent.Visible = True
                        DdlProvisoes.Enabled = True
                        MsgBox(Me.Page, "Registro com Baixa de Adiantamento no Título " & objAdtoXBaixa(0).CodigoTitulo & ", volte o mesmo para PREVISÃO. Com a alteração desse Registro " & txtRegistro.Text & " será gerado um novo número de Adiantamento, no Registro " & objAdtoXBaixa(0).CodigoTitulo & " vincule novamente o novo número do Adiantamento gerado.")
                        Exit Sub
                    End If
                End If
            End If

            If objCarteira.BaixaAdiantamento Or (ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento) Then
                MsgBox(Me.Page, "Registro de Baixa de Adiantamento nao podem ser alterados, exclua o titulo e refaça a Baixa. Baixas que tem o pedido na hora da exclusao voltam para a previsao e a baixa é desfeita porem o titulo ainda permanece no pedido.")
                lnkNovo.Parent.Visible = True
                DdlProvisoes.Enabled = True
                lnkExcluir.Parent.Visible = True
                Exit Sub
            End If

            Dim Mestre() As String = txtMestre.Text.Split(":")

            If CInt(Trim(Mestre(1))) > 0 Then
                MsgBox(Me.Page, "Registro de Agrupamento não pode ser desbloqueado, para alteração deve desfazer o Agrupamento")
                Exit Sub
            End If

            txtMovimento.Enabled = True
            txtProrrogacao.Enabled = True
            txtLiberarTitulo.Value = "S"
            DdlProvisoes.Enabled = True
            'txtDataBaixa.Enabled = True
            ImgValores.Enabled = True
            lnkNovo.Parent.Visible = True
            ddlIndexador.Enabled = True
        End If
    End Sub

    Protected Sub imgLimparPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLimparPedido.Click
        If Funcoes.VerificaPermissao(HTabela.Value, "LIBERAR") Then
            'txtLiberarPedido.Value = "S"
            'txtPedido.Text = "0"
            'cmdPedido.Enabled = True
            'ddlSafraAdto.Enabled = True
            'ddlSafraAdto.SelectedIndex = 0
            MsgBox(Me.Page, "Desabilitado, entre em contato com o Suporte.")
        Else
            MsgBox(Me.Page, "Usuário sem permissão para remover o Pedido.")
        End If
    End Sub

    Protected Sub imgLimparCheque_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLimparCheque.Click
        If Funcoes.VerificaPermissao(HTabela.Value, "LIBERAR") Then
            txtLiberarCheque.Value = "S"
            txtEmiteCheque.Value = "N"
            txtNumeroCheque.Text = "0"
            txtNumeroCheque.Enabled = True
        Else
            MsgBox(Me.Page, "Usuário sem permissão para liberar Registro")
        End If
    End Sub

    Protected Sub imgLimparAdto_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLimparAdto.Click
        If Funcoes.VerificaPermissao(HTabela.Value, "LIBERAR") Then
            Dim cart As New CarteiraFinanceira(ddlCarteiras.SelectedValue)

            If Not (cart.isAdiantamento And Not cart.BaixaAdiantamento) Then
                txtNumeroAdto.Text = "0"
                cmdAdiantamento.Enabled = True
            Else
                MsgBox(Me.Page, "Este titulo deu origem a um adiantamento. para apaga-lo, caso nao tenha baixas atreladas a ele mude para previsao e salve.")
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para remover o Adiantamento.")
        End If
    End Sub

    Protected Sub imgExtratoPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If txtRegistro.Text.Length = 0 Then
            MsgBox(Me.Page, "Consulte o Registro para visualização do Extrato.")
        ElseIf DdlEmpresaCliente.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa do Registro não encontrada.")
        ElseIf txtCodigoFornecedor.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Cliente do Registro não encontrado.")
        ElseIf txtPedido.Text.Length = 0 OrElse txtPedido.Text = "0" Then
            MsgBox(Me.Page, "Registro sem Pedido não pode ser visualizado.")
        Else
            Extrato.Emitir(Me.Page, FinanceiroNovo, DdlEmpresaCliente.SelectedValue.Split("-")(0), DdlEmpresaCliente.SelectedValue.Split("-")(1), "T",
                           txtPedido.Text.Trim, "", 0, "", Nothing, "", False, False, False, False, True)
            'Dim strQueryString As String = "?fim=" & DateTime.Now.ToString("dd/MM/yyyy")
            'strQueryString &= "&empresa=" & DdlEmpresaCliente.SelectedValue
            'strQueryString &= "&cliente=" & Replace(txtCodigoFornecedor.Value, ";", "-")
            'strQueryString &= "&pedido=" & txtPedido.Text
            'strQueryString &= "&es=ES"

            'campo = DdlEmpresaCliente.SelectedValue.ToString.Split("-")
            'Dim objEmpresa As New [Lib].Negocio.Cliente(campo(0), campo(1))

            'For Each row As [Lib].Negocio.ClientexTipo In objEmpresa.Tipos
            '    If row.CodigoTipo = eTipoCliente.Revenda Then
            '        strQueryString &= "&desprd=S"
            '    End If
            'Next
            'ScriptManager.RegisterClientScriptBlock(Me, Me.GetType, "CarregarHTML", "window.open(""ExtratoPedidoHTML.aspx" & strQueryString & """, ""relatorio"", ""status=no, toolbar=yes, location=no, menubar=yes, resizable=yes, height=600, width=800, scrollbars=yes"");", True)
        End If
    End Sub

    Protected Sub chkReprogramarAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If GridConsultaTitulos.Rows.Count > 0 Then
                Dim chkReprogramarAll As CheckBox = CType(sender, CheckBox)

                For Each rowgrid As GridViewRow In GridConsultaTitulos.Rows
                    Dim chkReprogramar As CheckBox = CType(rowgrid.FindControl("chkReprogramar"), CheckBox)
                    chkReprogramar.Checked = chkReprogramarAll.Checked
                Next
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkAllTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If GridConsultaTitulos.Rows.Count > 0 Then
                Dim chkAllTitulos As CheckBox = CType(sender, CheckBox)

                If chkProvisao.Checked Then
                    chkAllTitulos.Checked = False
                    MsgBox(Me.Page, "Função não pode ser utilizada com o parametro Provisao marcado na consulta. Desmarque e refaça a consulta novamente.")
                    Limpar_ConsultaTitulos(True)
                    Exit Sub
                End If

                Dim passed As Boolean = False
                For Each rowgrid As GridViewRow In GridConsultaTitulos.Rows
                    Dim chkTitulo As CheckBox = CType(rowgrid.FindControl("ChkGridTitulos"), CheckBox)

                    If Not chkTitulo.Enabled Then
                        chkAllTitulos.Checked = False
                        Exit Sub
                    End If

                    If chkAllTitulos.Checked Then
                        Dim strMoeda As String = rowgrid.Cells(10).Text.ToString.Split("-")(1)

                        If Not passed Then
                            txtRealDolar.Value = strMoeda
                            HiddenIndexador.Value = rowgrid.Cells(11).Text
                            passed = True
                        End If
                        chkTitulo.Checked = IIf(strMoeda <> txtRealDolar.Value OrElse rowgrid.Cells(11).Text <> HiddenIndexador.Value, False, True)
                    Else
                        chkTitulo.Checked = False
                        HiddenIndexador.Value = String.Empty
                        txtRealDolar.Value = String.Empty
                    End If
                Next
                TotalizadorTitulosAgrupados()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ChkGridTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkTitulo As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(chkTitulo.NamingContainer, GridViewRow)

        If chkTitulo.Checked Then

            Dim Titulo As New Titulo(CInt(GridConsultaTitulos.Rows(row.RowIndex).Cells(3).Text()))

            If Titulo.CodigoProvisao = eProvisao.Provisao Then
                chkTitulo.Checked = False
                MsgBox(Me.Page, "Titulo em Provisao não pode ser usado no Agrupamento.")
                Exit Sub
            End If

            Dim strMoeda() As String = GridConsultaTitulos.Rows(row.RowIndex).Cells(10).Text.ToString.Split("-")
            If txtRealDolar.Value.ToString.Length = 0 Then
                txtRealDolar.Value = strMoeda(1)
            End If

            If HiddenIndexador.Value.ToString.Length = 0 Then
                HiddenIndexador.Value = GridConsultaTitulos.Rows(row.RowIndex).Cells(11).Text
            End If

            If Not strMoeda(1) = txtRealDolar.Value Then
                chkTitulo.Checked = False
                MsgBox(Me.Page, "Agrupamento permitido apenas para Títulos com a mesma moeda.")
                Exit Sub
            End If

            If Not GridConsultaTitulos.Rows(row.RowIndex).Cells(11).Text = HiddenIndexador.Value Then
                chkTitulo.Checked = False
                MsgBox(Me.Page, "Agrupamento permitido apenas para Títulos com o mesmo indexador.")
                Exit Sub
            End If
        End If
        TotalizadorTitulosAgrupados()
    End Sub

    Public Sub TotalizadorTitulosAgrupados()
        Dim Quantidade As Integer = 0
        Dim Valor As Decimal = 0
        For Each row As GridViewRow In GridConsultaTitulos.Rows
            Dim chkTitulo As CheckBox = CType(row.FindControl("ChkGridTitulos"), CheckBox)
            If (chkTitulo.Checked) Then
                Quantidade = Quantidade + 1
                If (row.Cells(10).Text.Equals("R$-1")) Then
                    Valor = Valor + CDec(row.Cells(8).Text)
                Else
                    Valor = Valor + CDec(row.Cells(7).Text)
                End If
            End If
        Next
        If Quantidade > 1 Then
            lnkAgruparPagamento.Parent.Visible = True
        Else
            lnkAgruparPagamento.Parent.Visible = False
        End If
        If Quantidade = 0 Then
            lblTotalRegistroAgrupado.Text = String.Empty
            lblTotalRegistroAgrupado.Parent.Visible = False
            Dim ckAll As CheckBox = CType(GridConsultaTitulos.HeaderRow.FindControl("chkAllTitulos"), CheckBox)
            ckAll.Checked = False
        Else
            lblTotalRegistroAgrupado.Text = Quantidade & " Título(s) a selecionado(s) no valor total de: " & String.Format("{0:N2}", Valor)
            lblTotalRegistroAgrupado.Parent.Visible = True
        End If
    End Sub


    Protected Sub DdlProvisoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'If DdlProvisoes.SelectedIndex = 0 Then Exit Sub

        If DdlProvisoes.SelectedValue = 3 Then
            DdlProvisoes.SelectedValue = 2
            Exit Sub
        End If

        If Not DdlProvisoes.SelectedValue = 1 AndAlso txtProrrogacao.Enabled = False Then
            txtProrrogacao.Enabled = True
        End If

        'If lblAgrupar.Text = "AP" Then chkConciliado.Parent.Visible = True
        chkConciliado.Parent.Visible = (DdlProvisoes.SelectedValue = 1)
        chkConciliado.Enabled = (DdlProvisoes.SelectedValue = 1)

        '?
        'txtDataBaixa.Parent.Visible = (DdlProvisoes.SelectedValue = 1)
        'txtDataBaixa.Enabled = False

        'If txtRegistro.Text = "" Then
        '    txtDataBaixa.Enabled = True
        'End If

        If Session("objDestinoContabil" & HID.Value) Is Nothing AndAlso DdlProvisoes.SelectedValue = 1 Then
            If (Not String.IsNullOrWhiteSpace(txtCodigoFornecedor.Value)) Then
                Dim strCliente() As String = txtCodigoFornecedor.Value.Split("-")
                If (strCliente IsNot Nothing AndAlso strCliente.Length > 0) Then
                    Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
                    If objCliente.DesdobrarFornecedor = True Then
                        ucDestinoContabil.Limpar()
                        Dim parameters As New Dictionary(Of String, Object)
                        If Request.QueryString("idTabela").Equals("ContasAPagar") Then
                            parameters.Add("tipo", "P")
                        Else
                            parameters.Add("tipo", "R")
                        End If
                        Popup.ConsultaDeDestinoContabil(Me.Page, "objDestinoContabil" & HID.Value)
                        ucDestinoContabil.Carregar(parameters)
                    End If
                End If
            Else
                DdlProvisoes.SelectedIndex = 0
                MsgBox(Me.Page, "É necessário selecionar o campo fornecedor!")
            End If
        End If
    End Sub

    Protected Sub btnCancelarNonoVencimento_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Limpar_ConsultaTitulos(True)
    End Sub

    Protected Sub btnNovoVencimento_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If txtNovoVencimento.Text.Length = 0 Then
            MsgBox(Me.Page, "Informe o novo vencimento para os Títulos")
        ElseIf Not IsDate(txtNovoVencimento.Text) Then
            MsgBox(Me.Page, "Data para novo vencimento não é válida")
        Else
            SqlArray.Clear()

            Dim i As Integer = 0
            While i < GridConsultaTitulos.Rows.Count
                If CType(GridConsultaTitulos.Rows(i).FindControl("chkReprogramar"), CheckBox).Checked = True Then

                    Sql = "UPDATE " & HTabela.Value & vbCrLf &
                          "   Set Prorrogacao = '" & CDate(txtNovoVencimento.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf &
                          "     , UsuarioAlteracao = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf &
                          "     , UsuarioAlteracaoData = GETDATE()" & vbCrLf &
                          " Where Registro_Id = " & GridConsultaTitulos.Rows(i).Cells(3).Text()

                    SqlArray.Add(Sql)
                End If

                i += 1
            End While

            If SqlArray.Count > 0 Then
                If Banco.GravaBanco(SqlArray) Then
                    Limpar_ConsultaTitulos(True)
                    MsgBox(Me.Page, "Processo concluído.")
                Else
                    MsgBox(Me.Page, "Erro ao Salvar: " & Funcoes.EliminarCaracteresEspeciais(RTrim(Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            End If
        End If
    End Sub

    Protected Sub btnBoleto_Click(ByVal sender As Object, ByVal e As CommandEventArgs)
        Try
            Dim Titulo As New Titulo(CInt(txtRegistro.Text))

            Dim verBoleto As New BoletoBancario()

            Dim nomePDF As String = verBoleto.RetornarBoletoBancario(Titulo, DdlBancoPagador.SelectedValue)

            Response.Clear()
            Response.ContentType = "application/pdf"
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & nomePDF & ".pdf")
            Response.TransmitFile(Server.MapPath("~/Boletos/" & nomePDF & ".pdf"))
            Response.End()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        Finally
            'Kill(Server.MapPath("~/Boletos/" & e.CommandName & ".pdf"))
        End Try
    End Sub


#Region "PROCESSAR RETORNO BANCARIO"

    Dim IdentificacaoDoRegistro As Integer
    Dim IdentificacaoEmpresaNoBanco As Integer
    Dim TipoCnpjCpfPagador As Integer
    Dim CnpjCpfPagador As Integer
    Dim NomePagador As String
    Dim TipoServico As Integer
    Dim CodOrigemArq As Integer
    Dim NumRemessa As Integer
    Dim NumRetorno As Integer
    Dim DataGravacaoArq As Date
    Dim HoraGravacaoArq As String
    Dim TipoProcessamento As Integer
    Dim ReservadoEmpresa As String
    Dim ReservadoBanco As String
    Dim ReservadoExpansao As String
    Dim NumSequencialRegistro As Integer

    Public Function ValidaDados(ByVal infoarquivo As Object) As Boolean
        If infoarquivo.Name.ToString.Length = 0 Then
            Mensagem = "Arquivo não foi informado."
            Return False
        Else
            Return True
        End If
    End Function

    Private Function ValidarArquivo(ByVal infoarquivo As Object) As Boolean
        Dim arquivo As String = Server.MapPath("~/Files/" & infoarquivo.Name)
        Dim objArquivo As New StreamReader(arquivo)
        Dim strLinha As String
        'Dim campos() As String
        Dim intLinha As Integer
        Dim dsArquivo As New DataSet
        Dim SituacaoBancaria() As String

        Mensagem = ""

        dsArquivo = New DataSet
        Dim rowNew As DataRow

        With dsArquivo
            .Tables.Add()
            With .Tables(0).Columns
                .Add("CnpjCpfPagador", Type.GetType("System.String"))
                .Add("NomePagador", Type.GetType("System.String"))
                .Add("Registro", Type.GetType("System.Int32"))
                .Add("Cliente", Type.GetType("System.String"))
                .Add("Valor", Type.GetType("System.Decimal"))
                .Add("Vencimento", Type.GetType("System.DateTime"))
                .Add("DataPgto", Type.GetType("System.DateTime"))
                .Add("CodRetorno", Type.GetType("System.String"))
                .Add("DescRetorno", Type.GetType("System.String"))
                .Add("CK", Type.GetType("System.Boolean"))
            End With
        End With

        Session("ssRetornoDs" & HID.Value) = dsArquivo

        Dim ListTitulo As New [Lib].Negocio.ListTitulo
        Try
            Do While objArquivo.Peek >= 0
                intLinha += 1
                rowNew = dsArquivo.Tables(0).NewRow
                strLinha = objArquivo.ReadLine()
                If intLinha <> 1 Then
                    If Mid(strLinha, 252, 10).ToString.Trim <> "" AndAlso CInt(Mid(strLinha, 252, 10).ToString.Trim) > 0 Then
                        rowNew("CnpjCpfPagador") = Mid(strLinha, 11, 15)
                        rowNew("NomePagador") = Mid(strLinha, 26, 40)
                        rowNew("Registro") = Mid(strLinha, 252, 10)
                        rowNew("Cliente") = Mid(strLinha, 18, 30)
                        rowNew("Valor") = CDbl(Mid(strLinha, 205, 15)) / 100
                        rowNew("Vencimento") = Left(Mid(strLinha, 166, 8), 4) & "/" & Mid(Mid(strLinha, 166, 8), 5, 2) & "/" & Mid(Mid(strLinha, 166, 8), 7, 2)
                        rowNew("DataPgto") = Left(Mid(strLinha, 266, 8), 4) & "/" & Mid(Mid(strLinha, 266, 8), 5, 2) & "/" & Mid(Mid(strLinha, 266, 8), 7, 2)
                        rowNew("CodRetorno") = Mid(strLinha, 279, 2)

                        SituacaoBancaria = VerificaSituacoesBancarias(Mid(strLinha, 279, 2)).ToString.Split("-")

                        rowNew("DescRetorno") = SituacaoBancaria(0)
                        rowNew("CK") = IIf(SituacaoBancaria(1) = 1, True, False)
                        Dim Titulo As New [Lib].Negocio.Titulo(Mid(strLinha, 252, 10))
                        ListTitulo.Add(Titulo)

                        dsArquivo.Tables(0).Rows.Add(rowNew)
                        dsArquivo.AcceptChanges()
                    End If
                End If
            Loop

            GridRetornoTitulos.DataSource = dsArquivo
            GridRetornoTitulos.DataBind()

            Session("ssRetornoDs" & HID.Value) = dsArquivo
            Session("ssRetorno" & HID.Value) = ListTitulo

            If Mensagem.Length = 0 And Not ListTitulo Is Nothing Then
                Return True
            ElseIf Mensagem.Length = 0 And dsArquivo.Tables(0).Rows.Count = 0 Then
                Mensagem = "Arquivo selecionado para importação vázio ou com problema"
                Return False
            Else
                Return False
            End If
        Catch ex As Exception
            Mensagem = Funcoes.EliminarCaracteresEspeciais(ex.Message)
            Return False
        End Try

        If Mensagem.Length > 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function Upload(ByVal infoarquivo As Object) As Boolean
        Mensagem = ""
        Try
            'Verificamos se tem alguma coisa postada 
            If Not IsNothing(fup.PostedFile) Then
                'Pegamos as informacoes do arquivo postado 
                'Definimos onde ele será salvo 
                Dim strCaminho As String = Server.MapPath("Files/") & infoarquivo.Name
                'Salvamos o mesmo 
                fup.PostedFile.SaveAs(strCaminho)
                'Mensagem de confirmacao 
                Return True
            Else
                Mensagem = "Selecione um arquivo!"
                Return False
            End If
        Catch ex As Exception
            'Se der algum erro, exibimos a mensagem 
            Mensagem = "Há erros!. " & ex.Message
            Return False
        End Try
    End Function

    Private Sub ImportarRetorno()
        If fup.HasFile Then
            Dim infoArquivo As New IO.FileInfo(fup.PostedFile.FileName)
            If Upload(infoArquivo) Then
                If ValidaDados(infoArquivo) Then
                    Dim alSQL As New ArrayList
                    If ValidarArquivo(infoArquivo) Then
                        'btBaixar.Enabled = True
                        lnkBaixar.Enabled = True
                        TabContainer1.ActiveTabIndex = 5
                    Else
                        MsgBox(Me.Page, Mensagem.ToString)
                    End If
                Else
                    MsgBox(Me.Page, Mensagem.ToString)
                End If
            Else
                MsgBox(Me.Page, Mensagem.ToString)
            End If
        Else
            MsgBox(Me.Page, "Arquivo não encontrado!")
        End If
    End Sub

    Private Function VerificaSituacoesBancarias(ByVal CodRetorno As String)
        Dim Sql As String
        Dim ds As New DataSet
        Dim Resultado As String

        Sql = " SELECT     Situacao_Id, Descricao, Provisao " & vbCrLf &
              "   FROM SituacoesBancariasRetorno " & vbCrLf &
              "   WHERE Situacao_Id = '" & CodRetorno & "'"

        ds = Banco.ConsultaDataSet(Sql, "SituacoesBancariasRetorno")

        If ds.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                Resultado = Dr("Descricao") & "-" & Dr("Provisao")
                Return Resultado
            Next
        End If
        Return "Não Encontrada"
    End Function

#End Region

    Protected Sub BtValidarCodBarras_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If ckPreImpresso.Checked = False Then
            If DdlProvisoes.Text <> "" And DdlTiposDePagamentos.Text <> "" Then
                If DdlTiposDePagamentos.SelectedValue = 4 Then
                    If Trim(txtCodigoDeBarras.Text) <> "" Then
                        If CkbCodigoDeBarras.Checked Then txtCodigoDeBarras.Text = Funcoes.FormataLinhaDigitavelOriginal(txtCodigoDeBarras.Text)
                        'Deve passar a Empresa do Título para procurar em Datas não programáveis, não o Fornecedor - Furlan - 18/11/2013
                        'Dim strFornecedor As String() = txtCodigoFornecedor.Value.Split("-")
                        Dim strEmpresa As String() = DdlEmpresaCliente.SelectedValue.Split("-")
                        If Funcoes.ValidaCodigoBarras(txtCodigoDeBarras.Text, CkbCodigoDeBarras.Checked, txtProrrogacao.Text, txtValorCobrado.Text, strEmpresa(0), strEmpresa(1), Banco) Then
                            MsgBox(Me.Page, "Codigo de Barras Valido!")
                        Else
                            MsgBox(Me.Page, "Código de barras Inválido!" & Session("ssMessage").ToString())
                        End If
                    End If
                Else
                    MsgBox(Me.Page, "Preenchimento do Codigo de Barras Somente Aceito para Boletos Bancarios")
                End If
            Else
                MsgBox(Me.Page, "Tipo de Pagto e Previsao são Obrigatórios Para Validação Do Codigo De Barras...")
            End If
        Else
            MsgBox(Me.Page, "Sistema não Valida Codigo De Barras de Boletos Pré Impressos...")
        End If
    End Sub

    Protected Sub txtProrrogacao_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Try
            SessaoRecuperaTitulo()

            MostrarCotacao()

            If Not String.IsNullOrWhiteSpace(txtProrrogacao.Text) AndAlso IsDate(txtProrrogacao.Text) Then

                ObjTitulo.Prorrogacao = CDate(txtProrrogacao.Text)

                If Not ValidaData(txtProrrogacao.Text, "Vencimento", String.Empty, 0) Then
                    MsgBox(Me.Page, Mensagem)
                    txtProrrogacao.Text = String.Empty
                    txtProrrogacao.Focus()

                ElseIf (Not String.IsNullOrWhiteSpace(txtRegistro.Text)) AndAlso (Not String.IsNullOrWhiteSpace(ddlMoeda.SelectedValue)) Then

                    SessaoSalvaTitulo()


                    If ddlIndexador.SelectedValue = 99 Then
                        ValidaValores(True)
                    Else
                        Dim vlrMoeda As Decimal = Decimal.Zero
                        Dim sql As String = "SELECT * FROM Cotacoes WHERE Data_Id = '" & CDate(txtProrrogacao.Text).ToString("yyyy-MM-dd") & "' AND Indexador_Id = " & ddlIndexador.SelectedValue & " "
                        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Cotacoes")
                        If (ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0) Then
                            Dim row As DataRow = ds.Tables(0).Rows(0)
                            vlrMoeda = CDec(row("Indice"))
                        End If

                        If ddlMoeda.SelectedValue = 3 Then
                            txtValorCobrado.Text = String.Format("{0:N2}", (CDec(txtValorEmMoeda.Text) * vlrMoeda))
                            Dim vlrDiff As Decimal = CDec(txtValorCobrado.Text) - CDec(txtValorDoDocumento.Text)
                            If (vlrDiff > 0) Then
                                txtAcrescimos.Text = String.Format("{0:N2}", vlrDiff)
                            Else
                                txtDeducoes.Text = String.Format("{0:N2}", (vlrDiff * -1))
                            End If
                        Else
                            txtValorCobradoMoeda.Text = String.Format("{0:N2}", (CDec(txtValorDoDocumento.Text) / vlrMoeda))
                            txtValorEmMoeda.Text = txtValorCobradoMoeda.Text

                            'Dim vlrDiff As Decimal = CDec(txtValorCobradoMoeda.Text) - CDec(txtValorEmMoeda.Text)
                            'If (vlrDiff > 0) Then
                            '    txtAcrescimosMoeda.Text = String.Format("{0:N2}", vlrDiff)
                            'Else
                            '    txtDeducoesMoeda.Text = String.Format("{0:N2}", (vlrDiff * -1))
                            'End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtMovimento_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Try
            SessaoRecuperaTitulo()

            ObjTitulo.Baixa = CDate(txtMovimento.Text)

            If Not ObjTitulo.TituloOriginal Is Nothing Then
                ObjTitulo.TituloOriginal.Baixa = CDate(txtMovimento.Text)
            End If

            hdnMovimentoOriginal.Value = txtMovimento.Text

            SessaoSalvaTitulo()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao(HTabela.Value, "GRAVAR") Then
            Try
                GravaTitulo()
            Catch ex As Exception
                BloquearValores(ImgValores.ImageUrl.Contains("Bloquear"))
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        Else
            MsgBox(Me.Page, "Usuário sem permissão para incluir registro!")
        End If
    End Sub

    Protected Sub btnAdicionar_Click(sender As Object, e As EventArgs) Handles btnAdicionar.Click

        If String.IsNullOrWhiteSpace(txtDescricaoDocumento.Text) Then
            MsgBox(Me.Page, "Descrição do Documento não foi informada", eTitulo.Info)
            txtDescricaoDocumento.Focus()
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

            SessaoRecuperaTitulo()

            Dim fXd = New FinanceiroXDocumento(ObjTitulo)

            fXd.IUD = "I"
            fXd.Codigo = [Lib].Negocio.Numerador.PegarNumero(HttpContext.Current.Session("ssNomeServidor"), [Lib].Negocio.eTiposNumerador.Documentos)
            fXd.Descricao = txtDescricaoDocumento.Text
            fXd.NomeDoArquivo = NomeDoArquivo

            fXd.Arquivo = fupArquivo.FileBytes
            txtNomeDoArquivo.Text = NomeDoArquivo

            ObjTitulo.FinanceiroXDocumentos.Add(fXd)

            If ObjTitulo.Codigo > 0 Then

                If fXd.Salvar Then

                    ObjTitulo.FinanceiroXDocumentos = New ListFinanceiroXDocumento(ObjTitulo)

                    For Each fXd In ObjTitulo.FinanceiroXDocumentos
                        fXd.IUD = "I"
                    Next
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                    Exit Sub
                End If
            End If

            SessaoSalvaTitulo()

            CarregarDocumentos()

            txtDescricaoDocumento.Text = String.Empty
            txtNomeDoArquivo.Text = String.Empty
        Else
            MsgBox(Me.Page, "Selecione um arquivo.")
        End If
    End Sub

    Protected Sub imgDownload_Click(sender As Object, e As ImageClickEventArgs)
        Try
            SessaoRecuperaTitulo()

            Dim imgArquivo As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgArquivo.NamingContainer, GridViewRow)

            Dim CaminhoNomeArquivo As String = Server.MapPath("~/Files/" & ObjTitulo.FinanceiroXDocumentos(row.RowIndex).NomeDoArquivo)
            If Dir(CaminhoNomeArquivo).Length > 0 Then Kill(CaminhoNomeArquivo)

            Dim Fs As FileStream = New FileStream(CaminhoNomeArquivo, FileMode.Create)
            Fs.Write(ObjTitulo.FinanceiroXDocumentos(row.RowIndex).Arquivo, 0, ObjTitulo.FinanceiroXDocumentos(row.RowIndex).Arquivo.Length)
            Fs.Flush()
            Fs.Close()

            Funcoes.AbrirArquivo(Me.Page, "Files/" & ObjTitulo.FinanceiroXDocumentos(row.RowIndex).NomeDoArquivo)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarDocumentos()

        gridDocumentos.DataSource = ObjTitulo.FinanceiroXDocumentos
        gridDocumentos.DataBind()

        Dim x As Integer = 0
        While x < gridDocumentos.Rows.Count
            Dim pExtensao() As String = gridDocumentos.Rows(x).Cells(2).Text.Split(".")

            If pExtensao(1).ToUpper.Equals("XML") Then
                CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/xml16x16.png"
            ElseIf pExtensao(1).ToUpper.Equals("PDF") Then
                CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icopdf16X16.jpg"
            ElseIf pExtensao(1).ToUpper.Equals("XLS") Then
                CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
            ElseIf pExtensao(1).ToUpper.Equals("XLSX") Then
                CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
            ElseIf pExtensao(1).ToUpper.Equals("DOC") Then
                CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
            ElseIf pExtensao(1).ToUpper.Equals("DOCX") Then
                CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
            Else
                CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icoJpg.gif"
            End If

            CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ToolTip = String.Format("Download - {0}.{1}", pExtensao(0), pExtensao(1))
            CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"

            x += 1
        End While
    End Sub

    Protected Sub imgExcluirDocumento_Click(sender As Object, e As ImageClickEventArgs)

        Try
            If Funcoes.VerificaPermissao(HTabela.Value, "EXCLUIR") Then
                SessaoRecuperaTitulo()

                Dim img As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

                Dim Pos = ObjTitulo.FinanceiroXDocumentos.FindIndex(Function(s) s.Codigo = row.Cells(0).Text())

                ObjTitulo.FinanceiroXDocumentos(Pos).IUD = "D"

                If ObjTitulo.FinanceiroXDocumentos(Pos).Salvar Then

                    ObjTitulo.FinanceiroXDocumentos = New ListFinanceiroXDocumento(ObjTitulo)

                    For Each fXd In ObjTitulo.FinanceiroXDocumentos
                        fXd.IUD = "I"
                    Next

                    SessaoSalvaTitulo()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                    Exit Sub
                End If

                CarregarDocumentos()

            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        If Funcoes.VerificaPermissao(HTabela.Value, "LEITURA") Then
            ucConsultaTitulo.SetarHID(HID.Value)
            ucConsultaTitulo.Limpar()
            Popup.ConsultaDeTitulo(Me.Page, "objCarregarTitulo" & HID.Value.ToString, "txtRegistro")
        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        'chkManterLancamento.Checked = False
        Limpar(chkManterLancamento.Checked)
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        EmitirRecibo()
    End Sub

    Protected Sub lnkRelatorioAgrupamento_Click(sender As Object, e As EventArgs) Handles lnkRelatorioAgrupamento.Click
        Try
            If String.IsNullOrWhiteSpace(txtRegistro.Text) Then
                MsgBox(Me.Page, "É necessário informar o campo número do registro!")
            Else
                viewAgrupamento(txtRegistro.Text.Trim())
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, HTabela.Value)
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkConsultarTitulo_Click(sender As Object, e As EventArgs) Handles lnkConsultarTitulo.Click
        Dim Und As String = DdlUnidadeConsultaTitulos.SelectedValue
        Dim Emp As String = DdlEmpresaConsultaTitulos.SelectedValue
        If Und.Length > 0 Then
            DdlUnidadeConsultaTitulos.SelectedValue = Und
            ddl.Carregar(DdlEmpresaConsultaTitulos, CarregarDDL.Tabela.Empresas, DdlUnidadeConsultaTitulos.SelectedValue, True)
        End If
        If Emp.Length > 0 Then DdlEmpresaConsultaTitulos.SelectedValue = Emp
        TitulosConsulta()
    End Sub

    Protected Sub lnkRelatorioCTitulo_Click(sender As Object, e As EventArgs) Handles lnkRelatorioCTitulo.Click
        If Funcoes.VerificaPermissao(HTabela.Value, "RELATORIO") Then

            Dim crpt As New ReportDocument()

            Try
                If RbDiaGeral.Checked = True Then
                    DiaGeral(True)
                    Limpar_ConsultaTitulos(True)
                    Exit Sub
                End If
                If RbFilialDiario.Checked = True Then
                    FilialGeral()
                End If
                If RbCarteiraDia.Checked = True Then
                    CarteiraDia()
                End If
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Finally
                crpt.Close()
                crpt.Dispose()
            End Try
        Else
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(Session("ssMessage").ToString()))
        End If
    End Sub

    Protected Sub lnkExcelCTitulo_Click(sender As Object, e As EventArgs) Handles lnkExcelCTitulo.Click
        DiaGeral(False)
        Limpar_ConsultaTitulos(True)
        Exit Sub
    End Sub

    Protected Sub lnkExcelCTituloDados_Click(sender As Object, e As EventArgs) Handles lnkExcelCTituloDados.Click
        DiaGeralExcel()
    End Sub

    Protected Sub lnkLimparConsultaTitulo_Click(sender As Object, e As EventArgs) Handles lnkLimparConsultaTitulo.Click
        Limpar_ConsultaTitulos(True)
    End Sub

    Protected Sub lnkAgruparPagamento_Click(sender As Object, e As EventArgs) Handles lnkAgruparPagamento.Click
        Try
            Dim titDolar As Boolean = False
            Dim Mensagem As String = "Agrupamento dos Registros"
            lblAgrupar.Text = "AP"
            Dim TestaCliente As String = ""
            Dim ValidaCliente As String = "S"
            Dim FaltaTributo As Boolean = False
            Dim BaixaAdiantamento As Boolean = False
            Dim Registros As New ArrayList
            Dim SqlOrAnd As String = "WHERE "
            ChkLiberado.Checked = False
            txtRegistro.Text = String.Empty

            Sql = "SELECT UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, Indexador," & vbCrLf &
                  "       Sum(isnull(MoedaValorDoDocumento,0)) as MoedaValorDoDocumento," & vbCrLf &
                  "       Sum(ValorDoDocumento) as ValorDoDocumento," & vbCrLf &
                  "       Sum(Descontos) as Descontos," & vbCrLf &
                  "       Sum(Deducoes) as Deducoes, " & vbCrLf &
                  "       Sum(Juros) as Juros," & vbCrLf &
                  "       Sum(Acrescimos) as Acrescimos," & vbCrLf &
                  "       Sum(ValorLiquido) as ValorLiquido, " & vbCrLf &
                  "       Sum(MoedaValorLiquido) as MoedaValorLiquido" & vbCrLf &
                  "  FROM " & HTabela.Value & vbCrLf &
                  " WHERE Registro_Id = 99999999 " & vbCrLf

            Dim i As Integer = 0
            While i < GridConsultaTitulos.Rows.Count
                If CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Checked = True Then
                    Mensagem &= " - " & GridConsultaTitulos.Rows(i).Cells(3).Text()
                    Sql &= " or Registro_Id = " & GridConsultaTitulos.Rows(i).Cells(3).Text()
                    Registros.Add(GridConsultaTitulos.Rows(i).Cells(3).Text())

                    Dim strMoeda() As String = GridConsultaTitulos.Rows(i).Cells(10).Text.ToString.Split("-")
                    If strMoeda(0) = "U$" Then
                        titDolar = True
                    End If
                End If
                i += 1
            End While

            Sql &= " Group By UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, Indexador " & vbCrLf
            Dim ds As New DataSet

            ds = Banco.ConsultaDataSet(Sql, "Titulo")

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Agrupamento não pode ser realizado.")
                Exit Sub
            End If

            Sql = "Select cp.Registro_Id, cp.Carteira, cXp.Descricao, cp.Tributo, isnull(cXp.ContaClientes,'') as ContaClientes, " & vbCrLf &
                  "       case " & vbCrLf &
                  "       	when LEN(isnull(cXp.ContaClientes,'')) = 0 and LEN(isnull(cp.Tributo,'')) = 0 " & vbCrLf &
                  "           then 'S' " & vbCrLf &
                  "       	  else 'N' " & vbCrLf &
                  "       end as FaltaTributo, " & vbCrLf &
                  "       case " & vbCrLf &
                  "       	when isnull(cXp.BaixaAdiantamento,0) = 1 " & vbCrLf &
                  "           then 'S' " & vbCrLf &
                  "       	  else 'N' " & vbCrLf &
                  "       end as BaixaAdiantamento" & vbCrLf &
                  "  from " & HTabela.Value & " cp " & vbCrLf &
                  " inner join ComprasXProdutos cXp " & vbCrLf &
                  "    on cXp.Produto_Id = cp.Carteira " & vbCrLf

            Dim j As Integer = 0
            While j < GridConsultaTitulos.Rows.Count
                If CType(GridConsultaTitulos.Rows(j).FindControl("ChkGridTitulos"), CheckBox).Checked = True Then
                    Sql &= SqlOrAnd & " Registro_Id = " & GridConsultaTitulos.Rows(j).Cells(3).Text() & vbCrLf
                    SqlOrAnd = "   OR "
                End If
                j += 1
            End While

            Dim dscXp As New DataSet
            dscXp = Banco.ConsultaDataSet(Sql, "cpXCarteiras")

            For Each dr As DataRow In dscXp.Tables(0).Rows
                If dr("FaltaTributo") = "S" Then
                    FaltaTributo = True
                    MsgBox(Me.Page, "Titulo " & dr("Registro_Id") & " não foi selecionado o Tributo " & " da Carteira " & dr("Descricao"))
                    Exit For
                End If

                If dr("BaixaAdiantamento") = "S" Then
                    BaixaAdiantamento = True
                    MsgBox(Me.Page, "Titulo " & dr("Registro_Id") & " Baixa de Adiantamento não pode ser Agrupado. " & " Carteira " & dr("Descricao"))
                    Exit For
                End If
            Next

            If FaltaTributo Or BaixaAdiantamento Then
                Limpar(True)
                Exit Sub
            End If

            For Each Dr As DataRow In ds.Tables(0).Rows
                If TestaCliente <> "" Then
                    If TestaCliente <> Dr("Cliente") Then
                        ValidaCliente = "N"
                        Exit For
                    End If
                End If

                DdlUnidadeDeNegocioEmpresaCliente.SelectedIndex = DdlUnidadeDeNegocioEmpresaCliente.Items.IndexOf(DdlUnidadeDeNegocioEmpresaCliente.Items.FindByValue(Dr("UnidadeDeNegocio")))
                CargaEmpresaCliente()
                DdlEmpresaCliente.SelectedIndex = DdlEmpresaCliente.Items.IndexOf(DdlEmpresaCliente.Items.FindByValue(Dr("Empresa") & "-" & CStr(Dr("EndEmpresa"))))
                txtCodigoFornecedor.Value = Dr("Cliente") & "-" & Dr("EndCliente")
                txtFornecedor.Text = ConsultaCLientes(Dr("Cliente"), Dr("EndCliente"))
                DdlProvisoes.SelectedIndex = 1

                txtCodigoFavorecido.Value = txtCodigoFornecedor.Value
                txtFavorecido.Text = txtFornecedor.Text

                txtMovimento.Text = CDate(Today).ToString("dd/MM/yyyy")
                hdnMovimentoOriginal.Value = CDate(Today).ToString("dd/MM/yyyy")
                txtProrrogacao.Text = CDate(Today).ToString("dd/MM/yyyy")
                hdnProrrogacaoOriginal.Value = CDate(Today).ToString("dd/MM/yyyy")

                ddlMoeda.SelectedValue = txtRealDolar.Value
                ddlIndexador.SelectedValue = HiddenIndexador.Value
                ddlMoeda.Enabled = False
                ddlIndexador.Enabled = False

                If ddlMoeda.SelectedValue = 1 Then
                    If ddlIndexador.SelectedValue = 99 Then
                        Dr("MoedaValorLiquido") = 0
                    Else
                        Dr("MoedaValorLiquido") = Funcoes.ConverteParaMoedaExtrangeira(Dr("ValorLiquido"), ddlIndexador.SelectedValue, CDate(txtProrrogacao.Text).ToString("yyyy-MM-dd"))
                    End If
                End If

                If ddlMoeda.SelectedValue = 3 Then
                    If ddlIndexador.SelectedValue = 99 Then
                        Dr("ValorLiquido") = 0
                    ElseIf CDec(Dr("ValorLiquido")) = 0 Then
                        Dr("ValorLiquido") = Funcoes.ConverteParaMoedaOficial(Dr("MoedaValorLiquido"), ddlIndexador.SelectedValue, CDate(txtProrrogacao.Text).ToString("yyyy-MM-dd"))
                    End If
                End If

                txtValorDoDocumento.Text = CDec(Dr("ValorLiquido")).ToString("N2")
                txtValorEmMoeda.Text = CDec(Dr("MoedaValorLiquido")).ToString("N2")
                txtValorCobrado.Text = CDec(Dr("ValorLiquido")).ToString("N2")

                txtValorLiquido.Value = CDec(Dr("ValorLiquido"))

                txtDescontos.Text = "0,00"
                txtDeducoes.Text = "0,00"
                txtJuros.Text = "0,00"
                txtAcrescimos.Text = "0,00"

                If titDolar Then
                    txtValorDoDocumento.Enabled = True
                    txtDescontos.Enabled = True
                    txtDeducoes.Enabled = True
                    txtJuros.Enabled = True
                    txtAcrescimos.Enabled = True
                    txtValorCobrado.Enabled = True
                Else
                    txtValorDoDocumento.Enabled = False
                    txtDescontos.Enabled = False
                    txtDeducoes.Enabled = False
                    txtJuros.Enabled = False
                    txtAcrescimos.Enabled = False
                    txtValorCobrado.Enabled = False
                End If

                txtValorEmMoeda.Enabled = False
                txtDescontosMoeda.Enabled = False
                txtDeducoesMoeda.Enabled = False
                txtJurosMoeda.Enabled = False
                txtAcrescimosMoeda.Enabled = False
                txtValorCobradoMoeda.Enabled = False

                ChkLiberado.Checked = True
                TestaCliente = Dr("Cliente")
            Next

            If ValidaCliente = "S" Then
                Dim objCliente As New [Lib].Negocio.Cliente(txtCodigoFornecedor.Value.ToString.Split("-")(0), txtCodigoFornecedor.Value.ToString.Split("-")(1))
                txtHistorico.Text = "PGTO. " & objCliente.Nome

                Session("ssRegistros" & HID.Value) = Registros
                Session("ssObservacoes" & HID.Value) = Mensagem
                TabPanel4.Visible = False
                TabContainer1.ActiveTabIndex = 0
                Session("ssGrupado" & HID.Value) = "S"
                cmdPedido.Enabled = False
            Else
                Limpar(True)
                MsgBox(Me.Page, "Agrupamento só pode ser realizado se for para o mesmo cliente.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarBaixaLote()
        Try
            If CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).Codigo = 999999999 AndAlso CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).UsuarioLiberacao = "BaixaLoteFinanceiro" Then

                SqlArray.Clear()

                Dim vlrdoc As Decimal

                For Each tit In CType(Session("objRegistrosSelecionados" & HID.Value), ListTitulo)
                    tit.Carregando = False

                    tit.IUD = "U"

                    tit.CodigoProvisao = 1

                    tit.Processo = "BaixaFinanceiroPorLote"

                    tit.CodigoEmpresaPagadora = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).CodigoEmpresaPagadora
                    tit.EndEmpresaPagadora = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).EndEmpresaPagadora

                    tit.CodigoTipoPgto = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).CodigoTipoPgto

                    tit.CodigoBancoPagador = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).CodigoBancoPagador

                    tit.CodigoAgenciaPagadora = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).CodigoAgenciaPagadora
                    tit.DigitoAgenciaPagadora = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).DigitoAgenciaPagadora
                    tit.ContaPagadora = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).ContaPagadora
                    tit.DigitoContaPagadora = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).DigitoContaPagadora
                    tit.TipoContaPagadora = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).TipoContaPagadora
                    tit.ContaContabilPagadora = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).ContaContabilPagadora

                    If String.IsNullOrEmpty(tit.ContaContabilPagadora) OrElse String.IsNullOrWhiteSpace(tit.ContaContabilPagadora) Then
                        MsgBox(Me.Page, "Registro " & Registro & " esta sem a conta contabil pagadora. Verifique.", eTitulo.Info)
                        Exit Sub
                    End If

                    tit.CodigoCarteiraDoTitulo = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).CodigoCarteiraDoTitulo

                    tit.Prorrogacao = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).Prorrogacao
                    tit.DataMoeda = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).Prorrogacao
                    tit.Baixa = CType(Session("objBaixaLoteFinanceiro" & HID.Value), Titulo).Baixa

                    If Not Funcoes.VerificaAcesso(tit.CodigoEmpresa, tit.EnderecoEmpresa, tit.Baixa.ToString("dd-MM-yyyy"), "FINANCEIRO") Then
                        MsgBox(Me.Page, "Movimento já Fechado para esta data " & tit.Baixa.ToString("dd-MM-yyyy") & ", para empresa " & tit.CodigoEmpresa)
                        Exit Sub
                    End If

                    If tit.RegistroMestre = 0 Then tit.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado

                    tit.UsuarioBaixa = Session("ssNomeUsuario")
                    tit.UsuarioBaixaData = Today()

                    vlrdoc = tit.ValorDoDocumento

                    If tit.IndiceTitulo = 0 Then tit.IndiceTitulo = tit.DolarizaBaixa(tit.Baixa.ToString("yyyy/MM/dd"), tit.ValorDoDocumento, IIf(tit.CodigoIndexador = 99, 2, tit.CodigoIndexador))

                    If Not tit.IndiceTitulo > 0 Then
                        MsgBox(Me.Page, "Indice de Dólar não cadastrado ou não realizado para a data " & tit.Baixa.ToString("dd/MM/yyyy") & ", verifique.", eTitulo.Info)
                        Exit Sub
                    End If

                    tit.ValorDoDocumento = vlrdoc

                    tit.HistoricoRemessa = "TESTANDO GRUPADO SEM NADA"

                    tit.SalvarSql(SqlArray, False)

                    ''**************************************************************************************************
                    ''******************************* Razão - Valor do Documento  **************************************
                    ''**************************************************************************************************
                    'Registro = tit.Codigo
                    'Raz_Empresa = tit.CodigoEmpresa                     'EmpresaCliente
                    'Raz_EndEmpresa = tit.EnderecoEmpresa                'Endereco Empresa Cliente

                    'Raz_Conta = tit.Carteira.CodigoContaCliente         'Conta sem tributo
                    'If tit.Carteira.CodigoContaCliente.Length > 0 AndAlso tit.Carteira.ContaCliente.TemCliente Then
                    '    Raz_Cliente = tit.CodigoCliente                 'Cliente
                    '    Raz_EndCliente = tit.EndCliente                 'Endereco do Cliente
                    'Else
                    '    Raz_Cliente = ""                                'Cliente
                    '    Raz_EndCliente = 0                              'Endereco do Cliente
                    'End If

                    'If Not String.IsNullOrEmpty(tit.Tributo) Then
                    '    Dim Encargo2 As New Encargo(tit.Tributo)
                    '    Raz_Conta = Encargo2.ContaDebito                'Conta com tributo
                    '    Dim objPlaContaTributo As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                    '    If objPlaContaTributo.TemCliente Then
                    '        Raz_Cliente = tit.CodigoCliente             'Cliente
                    '        Raz_EndCliente = tit.EndCliente             'Endereco do Cliente
                    '    Else
                    '        Raz_Cliente = ""                            'Cliente
                    '        Raz_EndCliente = 0                          'Endereco do Cliente
                    '    End If
                    'End If

                    'Raz_UnidadeDeNegocio = tit.CodigoUnidadeDeNegocio

                    'Raz_ValorOficial = tit.ValorDoDocumento

                    'Raz_ValorMoeda = tit.MoedaValorDoDocumento

                    'Raz_Historico = Funcoes.EliminarCaracteresEspeciais(tit.Historico.Trim & ". " & tit.Observacoes.Trim)

                    ''Raz_DebitoCredito = "D"
                    'Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "D", "C")

                    'LanctosContabeis(tit.Codigo, tit.Baixa.ToString("yyyy/MM/dd"))


                    ''**************************************************************************************************
                    ''**************************************  Descontos  ***********************************************
                    ''**************************************************************************************************
                    'If tit.Descontos > 0 Then
                    '    Registro = tit.Codigo
                    '    Raz_Empresa = tit.CodigoEmpresa                     'EmpresaCliente
                    '    Raz_EndEmpresa = tit.EnderecoEmpresa                'Endereco Empresa Cliente

                    '    Raz_Conta = tit.Carteira.CodigoContaDesconto        'Conta de Descontos Obtidos
                    '    If tit.Carteira.CodigoContaDesconto.Length > 0 AndAlso tit.Carteira.ContaDesconto.TemCliente Then
                    '        Raz_Cliente = tit.CodigoCliente                 'Cliente
                    '        Raz_EndCliente = tit.EndCliente                 'Endereco do Cliente
                    '    Else
                    '        Raz_Cliente = ""                                'Cliente
                    '        Raz_EndCliente = 0                              'Endereco do Cliente
                    '    End If

                    '    Raz_UnidadeDeNegocio = tit.CodigoUnidadeDeNegocio

                    '    Raz_ValorOficial = tit.Descontos

                    '    Raz_ValorMoeda = tit.MoedaDescontos

                    '    Raz_Historico = Funcoes.EliminarCaracteresEspeciais(tit.Historico.Trim & ". " & tit.Observacoes.Trim)

                    '    'Raz_DebitoCredito = "C"
                    '    Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "C", "D")

                    '    LanctosContabeis(tit.Codigo, tit.Baixa.ToString("yyyy/MM/dd"))
                    'End If


                    ''**************************************************************************************************
                    ''**************************************  Deducoes  ************************************************
                    ''**************************************************************************************************
                    'If tit.Deducoes > 0 Then
                    '    Registro = tit.Codigo
                    '    Raz_Empresa = tit.CodigoEmpresa                     'EmpresaCliente
                    '    Raz_EndEmpresa = tit.EnderecoEmpresa                'Endereco Empresa Cliente

                    '    Raz_Conta = tit.Carteira.CodigoContaDeducao        'Conta de Deduções
                    '    If tit.Carteira.CodigoContaDeducao.Length > 0 AndAlso tit.Carteira.ContaDeducao.TemCliente Then
                    '        Raz_Cliente = tit.CodigoCliente                 'Cliente
                    '        Raz_EndCliente = tit.EndCliente                 'Endereco do Cliente
                    '    Else
                    '        Raz_Cliente = ""                                'Cliente
                    '        Raz_EndCliente = 0                              'Endereco do Cliente
                    '    End If

                    '    Raz_UnidadeDeNegocio = tit.CodigoUnidadeDeNegocio

                    '    Raz_ValorOficial = tit.Deducoes

                    '    Raz_ValorMoeda = tit.MoedaDeducoes

                    '    Raz_Historico = Funcoes.EliminarCaracteresEspeciais(tit.Historico.Trim & ". " & tit.Observacoes.Trim)

                    '    'Raz_DebitoCredito = "C"
                    '    Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "C", "D")

                    '    LanctosContabeis(tit.Codigo, tit.Baixa.ToString("yyyy/MM/dd"))
                    'End If


                    ''**************************************************************************************************
                    ''*************************************  Juros  ****************************************************
                    ''**************************************************************************************************
                    'If tit.Juros > 0 Then
                    '    Registro = tit.Codigo
                    '    Raz_Empresa = tit.CodigoEmpresa                     'EmpresaCliente
                    '    Raz_EndEmpresa = tit.EnderecoEmpresa                'Endereco Empresa Cliente

                    '    Raz_Conta = tit.Carteira.CodigoContaJuro            'Conta de Juros
                    '    If tit.Carteira.CodigoContaJuro.Length > 0 AndAlso tit.Carteira.ContaJuro.TemCliente Then
                    '        Raz_Cliente = tit.CodigoCliente                 'Cliente
                    '        Raz_EndCliente = tit.EndCliente                 'Endereco do Cliente
                    '    Else
                    '        Raz_Cliente = ""                                'Cliente
                    '        Raz_EndCliente = 0                              'Endereco do Cliente
                    '    End If

                    '    Raz_UnidadeDeNegocio = tit.CodigoUnidadeDeNegocio

                    '    Raz_ValorOficial = tit.Juros

                    '    Raz_ValorMoeda = tit.MoedaJuros

                    '    Raz_Historico = Funcoes.EliminarCaracteresEspeciais(tit.Historico.Trim & ". " & tit.Observacoes.Trim)

                    '    'Raz_DebitoCredito = "D"
                    '    Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "D", "C")

                    '    LanctosContabeis(tit.Codigo, tit.Baixa.ToString("yyyy/MM/dd"))
                    'End If


                    ''**************************************************************************************************
                    ''*************************************  Acrescimos  ***********************************************
                    ''**************************************************************************************************
                    'If tit.Acrescimos > 0 Then
                    '    Registro = tit.Codigo
                    '    Raz_Empresa = tit.CodigoEmpresa                     'EmpresaCliente
                    '    Raz_EndEmpresa = tit.EnderecoEmpresa                'Endereco Empresa Cliente

                    '    Raz_Conta = tit.Carteira.CodigoContaAcrescimo       'Conta de Juros
                    '    If tit.Carteira.CodigoContaAcrescimo.Length > 0 AndAlso tit.Carteira.ContaAcrescimo.TemCliente Then
                    '        Raz_Cliente = tit.CodigoCliente                 'Cliente
                    '        Raz_EndCliente = tit.EndCliente                 'Endereco do Cliente
                    '    Else
                    '        Raz_Cliente = ""                                'Cliente
                    '        Raz_EndCliente = 0                              'Endereco do Cliente
                    '    End If

                    '    Raz_UnidadeDeNegocio = tit.CodigoUnidadeDeNegocio

                    '    Raz_ValorOficial = tit.Acrescimos

                    '    Raz_ValorMoeda = tit.MoedaAcrescimos

                    '    Raz_Historico = Funcoes.EliminarCaracteresEspeciais(tit.Historico.Trim & ". " & tit.Observacoes.Trim)

                    '    'Raz_DebitoCredito = "D"
                    '    Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "D", "C")

                    '    LanctosContabeis(tit.Codigo, tit.Baixa.ToString("yyyy/MM/dd"))
                    'End If


                    ''************************************************************************************************************************
                    ''****************************  LIQUIDO DO DOCUMENTO *********************************************************************
                    ''************************************************************************************************************************
                    'Registro = tit.Codigo
                    'Raz_Empresa = tit.CodigoEmpresaPagadora             'EmpresaPagadora
                    'Raz_EndEmpresa = tit.EndEmpresaPagadora             'Endereco Empresa Pagadora

                    'Raz_Conta = tit.ContaContabilPagadora               'Conta Contábil
                    'Dim objPlaContaLiquido As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                    'If objPlaContaLiquido.TemCliente Then
                    '    Raz_Cliente = tit.CodigoCliente                 'Cliente
                    '    Raz_EndCliente = tit.EndCliente                 'Endereco do Cliente
                    'Else
                    '    Raz_Cliente = ""                                'Cliente
                    '    Raz_EndCliente = 0                              'Endereco do Cliente
                    'End If

                    'Raz_UnidadeDeNegocio = tit.UnidadeDeNegocioEmpresaPagadora.CodigoUnidade

                    'Raz_ValorOficial = tit.ValorLiquido

                    'Raz_ValorMoeda = tit.MoedaValorLiquido

                    'Raz_Historico = Funcoes.EliminarCaracteresEspeciais(tit.Historico.Trim & ". " & tit.Observacoes.Trim)

                    ''Raz_DebitoCredito = "C"
                    'Raz_DebitoCredito = IIf(HTabela.Value = "ContasAPagar", "C", "D")

                    'LanctosContabeis(tit.Codigo, tit.Baixa.ToString("yyyy/MM/dd"))

                    ''-------------------------------------------
                    ''Transferencias Financeiras
                    ''-------------------------------------------
                    'Sql = " SELECT EmpresaDebito, EnderecoDebito, EmpresaCredito, EnderecoCredito, EmpresaContabil, EnderecoContabil, " & vbCrLf &
                    '      "        ContaContabil, ClienteContabil,EndClienteContabil, " & vbCrLf &
                    '      IIf(HTabela.Value = "ContasAPagar", "DebitoCredito", "case when DebitoCredito='D'then 'C' else 'D' end DebitoCredito ") & vbCrLf &
                    '      "   FROM TransferenciasFinanceiras " & vbCrLf &
                    '      "  WHERE EmpresaDebito   ='" & tit.CodigoEmpresa & "'" & vbCrLf &
                    '      "    and EnderecoDebito  = " & tit.EnderecoEmpresa & vbCrLf &
                    '      "    and EmpresaCredito  ='" & tit.CodigoEmpresaPagadora & "'" & vbCrLf &
                    '      "    and EnderecoCredito = " & tit.EndEmpresaPagadora

                    'For Each DrT As DataRow In Banco.ConsultaDataSet(Sql, "TransferenciasFinanceiras").Tables(0).Rows
                    '    Raz_Empresa = DrT("EmpresaContabil")                'EmpresaCliente
                    '    Raz_EndEmpresa = DrT("EnderecoContabil")            'Endereco Empresa Cliente
                    '    Raz_Conta = DrT("ContaContabil")                    'Grupo de Contas
                    '    Raz_Cliente = DrT("ClienteContabil")                'Cliente
                    '    Raz_EndCliente = DrT("EndClienteContabil")          'Endereco do Cliente

                    '    Raz_UnidadeDeNegocio = tit.EmpresaPagadora.CodigoUnidadeDeNegocio

                    '    Raz_ValorOficial = tit.ValorLiquido

                    '    Raz_ValorMoeda = tit.MoedaValorLiquido

                    '    Raz_Historico = Funcoes.EliminarCaracteresEspeciais(tit.Historico.Trim & ". " & tit.Observacoes.Trim)

                    '    Raz_DebitoCredito = DrT("DebitoCredito")            'Debito/Credito

                    '    LanctosContabeis(tit.Codigo, tit.Baixa.ToString("yyyy/MM/dd"))
                    'Next

                Next

                If SqlArray.Count > 0 Then
                    If SqlArray.Count >= (CType(Session("objRegistrosSelecionados" & HID.Value), ListTitulo).Count * 3) Then
                        If Banco.GravaBanco(SqlArray) Then
                            Limpar_ConsultaTitulos(True)

                            MsgBox(Me.Page, "Registros baixados com sucesso.", eTitulo.Sucess)
                        Else
                            MsgBox(Me.Page, Session("ssMessage"), eTitulo.Erro)
                        End If
                    Else
                        MsgBox(Me.Page, "A contabilização da Baixa por Lote está em desacordo, refaça a operação. Caso o erro persista entre em contato com o Suporte.", eTitulo.Erro)
                    End If
                Else
                    MsgBox(Me.Page, "Houve um erro no processo da Baixa por Lote, refaça a operação. Caso o erro persista entre em contato com o Suporte.", eTitulo.Erro)
                End If
            Else
                MsgBox(Me.Page, "Informações para Baixa por Lote não foram encontradas, refaça a operação. Caso o erro persista entre em contato com o Suporte.", eTitulo.Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message), eTitulo.Erro)
        Finally
            Session.Remove("objRegistrosSelecionados" & HID.Value)
            Session.Remove("objBaixaLoteFinanceiro" & HID.Value)
        End Try
    End Sub

    Protected Sub lnkBaixarSelecionados_Click(sender As Object, e As EventArgs) Handles lnkBaixarSelecionados.Click
        If Funcoes.VerificaPermissao(HTabela.Value, "GRAVAR") Then
            Try
                Dim carregarBaixaLote As Boolean = True
                Dim titulosSelecionados As Integer = 0
                Dim tituloComTributo As Integer = 0
                Dim bltMensagem As String = String.Empty

                Session.Remove("objRegistrosSelecionados" & HID.Value)
                Session.Remove("objBaixaLoteFinanceiro" & HID.Value)

                Dim objTitulosSelecionados As New ListTitulo()

                For Each rowgrid As GridViewRow In GridConsultaTitulos.Rows
                    If CType(rowgrid.FindControl("ChkGridTitulos"), CheckBox).Checked Then
                        titulosSelecionados += 1

                        Registro = CInt(rowgrid.Cells(3).Text())

                        Dim bltTitulo As New Titulo(Registro, IIf(HTabela.Value = "ContasAPagar", "P", "R"))

                        If bltTitulo.ValorDoDocumento <= 0 OrElse bltTitulo.ValorLiquido <= 0 Then
                            carregarBaixaLote = False
                            bltMensagem = "Registro " & Registro & " com valor 0(zero) não pode ser utilizado no processo."
                        End If

                        If CType(rowgrid.FindControl("ChkLiberado"), CheckBox).Checked = False AndAlso Not Funcoes.VerificaPermissao("AutorizacoesDePagamentos", "GRAVAR") Then
                            If ddlCarteiras.SelectedValue = "001001057" OrElse
                               ddlCarteiras.SelectedValue = "001001069" OrElse
                               ddlCarteiras.SelectedValue = "001001067" OrElse
                               ddlCarteiras.SelectedValue = "001001005" OrElse
                               ddlCarteiras.SelectedValue = "001001052" OrElse
                               ddlCarteiras.SelectedValue = "001001074" Then
                            ElseIf CDbl(txtValorCobrado.Text) < 151 Then
                            Else
                                carregarBaixaLote = False
                                bltMensagem = "Registro " & Registro & " não esta autorizado para pagamento."
                            End If
                        End If

                        If String.IsNullOrEmpty(bltTitulo.CodigoCarteira) Then
                            carregarBaixaLote = False
                            bltMensagem = "Registro " & Registro & " não possui Finalidade Financeira selecionada."
                        End If

                        If bltTitulo.Carteira.isAdiantamento Then
                            carregarBaixaLote = False
                            bltMensagem = "Registro " & Registro & " de adiantamento não pode ser utilizado no processo."
                        ElseIf bltTitulo.Carteira.BaixaAdiantamento Then
                            carregarBaixaLote = False
                            bltMensagem = "Registro " & Registro & " de baixa de adiantamento não pode ser utilizado no processo."
                        End If

                        If Not bltTitulo.CodigoMoeda = 1 Then
                            carregarBaixaLote = False
                            bltMensagem = "Registro " & Registro & " deve ser baixado separadamente por se tratar de um título em dólar."
                        End If

                        objTitulosSelecionados.Add(bltTitulo)

                        Sql = "SELECT Tributo_Id as Codigo, (Encargos.Descricao + ' - ' + Tributo_Id) as Descricao " & vbCrLf &
                              "  FROM CarteirasXTributos " & vbCrLf &
                              " INNER JOIN Encargos " & vbCrLf &
                              "         ON CarteirasXTributos.Tributo_ID = Encargos.Encargo_id " & vbCrLf &
                              " Where CarteirasXTributos.Carteira_id = '" & bltTitulo.CodigoCarteira & "'"

                        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "CarteirasXTributos")

                        If Not ds Is Nothing AndAlso Not ds.Tables("CarteirasXTributos") Is Nothing AndAlso ds.Tables("CarteirasXTributos").Rows.Count > 0 Then
                            tituloComTributo += 1
                        End If
                    End If
                Next

                If Not carregarBaixaLote Then
                    MsgBox(Me.Page, bltMensagem, eTitulo.Info)
                    Exit Sub
                End If

                If tituloComTributo > 0 AndAlso titulosSelecionados <> tituloComTributo Then
                    MsgBox(Me.Page, "Registro de Finalidade Financeira com e sem Tributos, verifique.", eTitulo.Info)
                    Exit Sub
                End If

                Session("objRegistrosSelecionados" & HID.Value) = objTitulosSelecionados

                If objTitulosSelecionados.Count = 0 OrElse objTitulosSelecionados.Count < 2 Then
                    MsgBox(Me.Page, "Baixa em Lote só é permitida para 2 ou mais Registros selecionados.", eTitulo.Info)
                    Exit Sub
                End If

                Session("ssCampo" & HID.Value) = "LivreClasse"

                ucBaixaLoteFinanceiro.SetarHID(HID.Value)

                ucBaixaLoteFinanceiro.Limpar()

                Popup.BaixaLoteFinanceiro(Me.Page, "objBaixaLoteFinanceiro" & HID.Value)
            Catch ex As Exception
                MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message), eTitulo.Erro)
            End Try
        Else
            MsgBox(Me.Page, "Usuário sem permissão para baixar selecionados!", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkReprogramar_Click(sender As Object, e As EventArgs) Handles lnkReprogramar.Click
        pnlReprogramaVencimentos.Visible = True
        txtNovoVencimento.Focus()
    End Sub

    Protected Sub lnkRecibo_Click(sender As Object, e As EventArgs) Handles lnkRecibo.Click
        '' rotina de imprimir recibo por registros checados. .'.
        Dim xextenso As String = ""
        Dim yextenso As String = ""
        Dim dsEmitir As New DataSet
        'Dim dtEmitir As DataTable
        Dim row As DataRow
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim mes As Integer
        Dim MESX As String = ""
        'Dim NumeroDoCheque As Integer
        'Dim NumeroDoChequeX As String
        Dim Historico As String = ""
        Dim RegistroI As String = ""
        Dim RegistroS As String = ""
        Dim SqlArray As New ArrayList
        'Dim Sqlupdate As String
        'Dim Campo() As String
        'Dim Mensagem As String
        ''*Definicao de campos da empresa
        Dim Empresa As String = ""
        Dim EndEmpresa As String = ""
        Dim ENome As String = ""
        Dim EEndereco As String = ""
        Dim ECep As String = ""
        Dim ECidade As String = ""
        Dim EEstado As String = ""
        Dim ECnpj As String = ""
        Dim EInscricao As String = ""
        Dim Efone As String = ""
        Dim EBairro As String = ""
        Dim EComplemento As String = ""
        Dim ENumero As Integer
        ''*Definicao de campos da empresa Fim 
        ''* DAdos do cliente inicio
        Dim Cliente As String = ""
        Dim EndCliente As String = ""
        Dim CNome As String = ""
        Dim CEndereco As String = ""
        Dim CCep As String = ""
        Dim CCidade As String = ""
        Dim CEstado As String = ""
        Dim CCnpj As String = ""
        Dim CInscricao As String = ""
        Dim Cfone As String = ""
        Dim CBairro As String = ""
        Dim CComplemento As String = ""
        Dim CNumero As Integer

        Dim CBancoCliente As String = ""
        Dim CAgenciaCliente As String = ""
        Dim CDigitoAgenciaCliente As String = ""
        Dim CCcontaCliente As String = ""
        Dim CDigitoContaCliente As String = ""

        ''* Dados do Cliente fim 
        ''* Campos do DAta Set + titulos inicio
        'Dim myDataRow As DataRow
        '' Cria data Set que vai ser utilizado no relatorio 
        Dim dsRecibo As New DataSet
        Dim dtRecibo As DataTable
        Dim TnumeroDoCheque As Integer
        ''Dim row As DataRow
        ''Dim RegistroI As String
        ''Dim RegistroS As String
        dtRecibo = New DataTable("ReciboAPagar")
        '' campos da empresa 
        dtRecibo.Columns.Add("ENome", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EEndereco", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ECep", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ECidade", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EEstado", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ECnpj", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EInscricao", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EFone", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EBairro", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EComplemento", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ENumero", Type.GetType("System.Int32")).DefaultValue = 0
        ' ''campos do cliente
        dtRecibo.Columns.Add("CNome", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CEndereco", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CCep", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CCidade", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CEstado", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CCnpj", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CInscricao", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CFone", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CBairro", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CComplemento", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CNumero", Type.GetType("System.Int32")).DefaultValue = 0
        '' campos to titulo 
        dtRecibo.Columns.Add("TNumtit", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TValor", Type.GetType("System.Decimal"))
        dtRecibo.Columns.Add("TExtenso", Type.GetType("System.String"))
        dtRecibo.Columns.Add("THistorico", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TDia", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TMes", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TAno", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TFormaPagto", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TBanco", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TAgencia", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TConta", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TVencimento", Type.GetType("System.DateTime"))
        dtRecibo.Columns.Add("TBaixa", Type.GetType("System.DateTime"))
        dtRecibo.Columns.Add("TRecibo", Type.GetType("System.Int32")).DefaultValue = 0
        dtRecibo.Columns.Add("TRegistro", Type.GetType("System.Int32")).DefaultValue = 0
        dtRecibo.Columns.Add("TNumeroDoCheque", Type.GetType("System.Int32")).DefaultValue = 0
        '' Campos novos do cheque 05/05/2010
        dtRecibo.Columns.Add("TValorDoDocumento", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TDescontos", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TDeducoes", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TJuros", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TAcrescimos", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TDigito", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TDigitoAgencia", Type.GetType("System.String"))

        '' Campos novos do cheque 05/05/2010
        dsRecibo.Tables.Add(dtRecibo)
        Dim ValorCobrado As Decimal
        Dim ValorDoDocumento As Decimal
        Dim Juros As Decimal
        Dim Acrescimos As Decimal
        Dim Descontos As Decimal
        Dim deducoes As Decimal
        Dim TipoPagto As Integer
        Dim FormaDePagamento As String = ""
        '' Campos novos do cheque 05/05/2010
        Dim TvalorDoDocumento As Decimal
        Dim Tdescontos As Decimal
        Dim Tdeducoes As Decimal
        Dim TJuros As Decimal
        Dim TAcrescimos As Decimal
        Dim Tdigito As String = ""
        '' Campos novos do cheque 05/05/2010
        ''* Campos do Data Set + titulos fim.
        Dim emitirRecibo As Boolean = False

        Dim dsCP As DataSet
        While i < GridConsultaTitulos.Rows.Count
            If CType(GridConsultaTitulos.Rows(i).FindControl("chkRecibo"), CheckBox).Checked = True Then
                emitirRecibo = True
                'row = dtEmitir.NewRow()

                '' Rotina para leitura de registro buscando a data da baixa. 
                Dim DataBaixa As String = ""
                Dim dataVencimento As String = ""
                ''Dim Dr2 As New DataRow
                Registro = GridConsultaTitulos.Rows(i).Cells(3).Text()
                ''txtRegistro.Text = Registro
                Sql = "SELECT Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf &
                      "       UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente,Isnull(TipoContaCliente,'C') AS TipoContaCliente, " & vbCrLf &
                      "       ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf &
                      "       ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, isnull(EmpresaPedido,'') AS EmpresaPedido, isnull(EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(Pedido,0) AS Pedido, isnull(PedidoFixacao,0) AS PedidoFixacao, isnull(Procuracao,0) AS Procuracao, ValorDoDocumento, " & vbCrLf &
                      "       Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, ISNULL(MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, ISNULL(MoedaDescontos, 0) " & vbCrLf &
                      "       AS MoedaDescontos, ISNULL(MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(MoedaJuros, 0) AS MoedaJuros, ISNULL(MoedaAcrescimos, 0) AS MoedaAcrescimos, " & vbCrLf &
                      "       ISNULL(MoedaValorLiquido, 0) AS MoedaValorLiquido, (case when len(convert(nvarchar(4000),Observacoes)) > 0 then Historico + ' ' + convert(nvarchar(4000),Observacoes ) else Historico end) as Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, " & vbCrLf &
                      "       solicitacao, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, " & vbCrLf &
                      "       UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, isnull(RegistroMestre, 0) as RegistroMestre, Observacoes, SituacaoBancaria,T.Descricao, ISNULL(NumeroDoCheque, 0) AS NumeroDoCheque  " & vbCrLf &
                      "  FROM " & HTabela.Value & vbCrLf &
                      " INNER JOIN TIPOSDEPAGAMENTOS AS T " & vbCrLf &
                      "    ON T.TipoDePagamento_ID = TipoPagto " & vbCrLf &
                      " WHERE Registro_Id = " & Registro

                dsCP = Banco.ConsultaDataSet(Sql, "Titulo")
                If dsCP.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Tipo de pagamento do Título não foi definido.")
                    Exit Sub
                End If


                For Each Dr1 As DataRow In dsCP.Tables(0).Rows
                    DataBaixa = CDate(Dr1("Baixa")).ToString("dd/MM/yyyy")
                    dataVencimento = CDate(Dr1("Vencimento")).ToString("dd/MM/yyyy")
                    Empresa = Dr1("Empresa")
                    EndEmpresa = Dr1("EndEmpresa")
                    Cliente = Dr1("Cliente")
                    EndCliente = Dr1("EndCliente")
                    ValorCobrado = 0
                    ValorDoDocumento = Dr1("ValorDoDocumento")
                    Juros = Dr1("Juros")
                    Acrescimos = Dr1("Acrescimos")
                    Descontos = Dr1("Descontos")
                    deducoes = Dr1("Deducoes")
                    If Dr1("Observacoes").ToString.Length > 0 Then
                        Historico = Funcoes.EliminarCaracteresEspeciais(Dr1("Historico")) & ". " & Funcoes.EliminarCaracteresEspeciais(Dr1("Observacoes"))
                    Else
                        Historico = Funcoes.EliminarCaracteresEspeciais(Dr1("Historico"))
                    End If
                    CBancoCliente = Dr1("BancoCliente")
                    CAgenciaCliente = Dr1("AgenciaCliente")
                    CDigitoAgenciaCliente = Dr1("DigitoAgenciaCliente")
                    CCcontaCliente = Dr1("ContaCliente")
                    CDigitoContaCliente = Dr1("DigitoContaCliente")
                    ValorCobrado = ValorDoDocumento + Juros + Acrescimos - Descontos - deducoes
                    TipoPagto = Dr1("TipoPagto")
                    FormaDePagamento = Dr1("Descricao")
                    TnumeroDoCheque = Dr1("NumeroDoCheque")
                    Tdescontos = Dr1("Descontos")
                    Tdeducoes = Dr1("Deducoes")
                    TJuros = Dr1("Juros")
                    TAcrescimos = Dr1("Acrescimos")
                    TvalorDoDocumento = Dr1("ValorDodocumento")
                    Tdigito = Dr1("DigitoContaCliente")
                Next

                '' Rotina para leitura de registro buscando a data da baixa (fim).
                '' Dados da empresa
                'Dim GrupoEmpresa As String
                Dim dr As DataRow
                '' Dados da empresa - fim 
                '' Consultado empresa 
                Sql = " SELECT Emp.Cliente_Id ," & vbCrLf &
                      "        Emp.Nome, Emp.Cidade," & vbCrLf &
                      "        Emp.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf &
                      "        Emp.Endereco , Emp.Cep," & vbCrLf &
                      "        Emp.Inscricao, Emp.Telefone," & vbCrLf &
                      "        Emp.Bairro, Emp.Complemento," & vbCrLf &
                      "        Emp.Numero " & vbCrLf &
                      "   FROM Clientes Emp" & vbCrLf &
                      "  WHERE Emp.Cliente_Id = '" & Empresa & "'" & vbCrLf
                DS = Banco.ConsultaDataSet(Sql, "Clientes")
                If DS.Tables(0).Rows.Count > 0 Then
                    For Each dr In DS.Tables(0).Rows
                        ENome = dr("Nome")
                        EEndereco = dr("Endereco")
                        ECep = dr("Cep")
                        ECidade = dr("Cidade")
                        EEstado = dr("Estado")
                        ECnpj = dr("Cliente_id")
                        EInscricao = dr("Inscricao")
                        Efone = dr("Telefone")
                        EBairro = dr("Bairro")
                        EComplemento = dr("Complemento")
                        ENumero = dr("Numero")
                        Exit For
                    Next
                End If
                '' Consultando Empresa - fim 
                ''**************************************************************
                '' Consultado Cliente 
                Sql = " SELECT Cli.Cliente_Id ," & vbCrLf &
                      "        Cli.Nome, Cli.Cidade," & vbCrLf &
                      "        Cli.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf &
                      "        Cli.Endereco , Cli.Cep," & vbCrLf &
                      "        Cli.Inscricao, Cli.Telefone," & vbCrLf &
                      "        Cli.Bairro, Cli.Complemento," & vbCrLf &
                      "        Cli.Numero " & vbCrLf &
                      "   FROM Clientes Cli " & vbCrLf &
                      "  WHERE Cli.Cliente_Id = '" & Cliente & "'" & vbCrLf &
                      "    AND Cli.Endereco_id = " & EndCliente & vbCrLf
                DS = Banco.ConsultaDataSet(Sql, "Clientes")
                If DS.Tables(0).Rows.Count > 0 Then
                    For Each dr In DS.Tables(0).Rows
                        CNome = dr("Nome")
                        CEndereco = dr("Endereco")
                        CCep = dr("Cep")
                        CCidade = dr("Cidade")
                        CEstado = dr("Estado")
                        CCnpj = dr("Cliente_id")
                        CInscricao = dr("Inscricao")
                        Cfone = dr("Telefone")
                        CBairro = dr("Bairro")
                        CComplemento = dr("Complemento")
                        CNumero = dr("Numero")
                        Exit For
                    Next
                End If
                ' Cria Data Sete que vai ser utilizado no relatorio
                ' Move campos para o Data Set.
                ' Move campos da Empresa
                row = dtRecibo.NewRow()
                row("ENome") = ENome
                row("EEndereco") = EEndereco
                row("ECep") = ECep
                row("ECidade") = ECidade
                row("EEstado") = EEstado
                row("ECnpj") = Funcoes.FormatarCpfCnpj(ECnpj)
                row("EInscricao") = EInscricao
                row("EFone") = Efone
                row("EBairro") = EBairro
                row("EComplemento") = EComplemento
                row("ENumero") = ENumero
                '' Move campos do Cliente / fornecedor
                row("CNome") = CNome
                row("CEndereco") = CEndereco
                row("CCep") = CCep
                row("CCidade") = CCidade
                row("CEstado") = CEstado
                row("CCnpj") = Funcoes.FormatarCpfCnpj(CCnpj)
                row("CInscricao") = CInscricao
                row("CFone") = Cfone
                row("CBairro") = CBairro
                row("CComplemento") = CComplemento
                row("CNumero") = CNumero
                '' Move campos do Titulo
                row("Tnumtit") = Registro
                '' feito
                row("TValor") = ValorCobrado
                row("TNumeroDoCheque") = TnumeroDoCheque
                '' campos novos 05/05/10
                row("TValorDoDocumento") = TvalorDoDocumento
                row("TDescontos") = Tdescontos
                row("TDeducoes") = Tdeducoes
                row("TJuros") = TJuros
                row("TAcrescimos") = TAcrescimos
                row("TDigito") = Tdigito
                row("TDigitoAgencia") = CDigitoAgenciaCliente

                '' campos novos 05/05/10



                Dim valcobradostr As String
                valcobradostr = CStr(ValorCobrado)
                txtValorCobrado.Text = ValorCobrado

                Valor = Replace(txtValorCobrado.Text, ".", "")

                ''* Rotina de extenso inicio
                yextenso = "("
                yextenso &= UCase(Funcoes.Extenso(txtValorCobrado.Text(), "Real", "Reais"))
                yextenso &= " *"
                xextenso = yextenso
                For j = 1 To (120 - Len(xextenso))
                    xextenso &= " *"
                Next
                xextenso &= ")"
                row("TExtenso") = xextenso
                ''* Rotina de extenso fim 
                row("THistorico") = Historico
                row("TDia") = Day(DataBaixa)
                row("TAno") = Year(DataBaixa)

                row("TMes") = Month(DataBaixa)
                ''* Rotina do Mes inicio
                mes = Month(DataBaixa)
                If mes = 1 Then
                    MESX = "JANEIRO"
                End If

                If mes = 2 Then
                    MESX = "FEVEREIRO"
                End If

                If mes = 3 Then
                    MESX = "MARCO"
                End If

                If mes = 4 Then
                    MESX = "ABRIL"
                End If

                If mes = 5 Then
                    MESX = "MAIO"
                End If

                If mes = 6 Then
                    MESX = "JUNHO"
                End If

                If mes = 7 Then
                    MESX = "JULHO"
                End If

                If mes = 8 Then
                    MESX = "AGOSTO"
                End If

                If mes = 9 Then
                    MESX = "SETEMBRO"
                End If

                If mes = 10 Then
                    MESX = "OUTUBRO"
                End If

                If mes = 11 Then
                    MESX = "NOVEMBRO"
                End If

                If mes = 12 Then
                    MESX = "DEZEMBRO"
                End If

                row("TMes") = MESX
                ''* rotina do mes fim 


                row("TFormaPagto") = FormaDePagamento
                row("TBanco") = CBancoCliente
                row("TAgencia") = CAgenciaCliente
                row("TConta") = CCcontaCliente
                row("TVencimento") = dataVencimento
                row("TBaixa") = DataBaixa

                '' calculo de registro . 
                RegistroI = Registro
                RegistroS = " "
                RegistroS = Trim(RegistroS)
                If Len(RegistroS) < 6 Then
                    For k = 1 To (6 - Len(Trim(RegistroI)))
                        RegistroS &= "0"
                    Next
                End If
                row("Trecibo") = Trim(RegistroS) & (Trim(RegistroI))
                row("TRegistro") = Trim(RegistroS) & (Trim(RegistroI))
                dtRecibo.Rows.Add(row)

                ''dsRecibo.Tables.Add(dtRecibo)
                '' Move campos para o Data Fim
                ''
                ''Feito ate aqui falta incluir relatorio e acertar forma de pagamento . 
                '' e corrigir msg de erros
            End If
            i = i + 1
        End While

        If emitirRecibo Then
            'Imagem
            Dim dtImagem As New DataTable("Images")
            dtImagem.Columns.Add("path", GetType(String))
            dtImagem.Columns.Add("image", GetType(System.Byte()))

            Dim drImagem As DataRow = dtImagem.NewRow()
            Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & Session("ssImagemEmpresa"))

            drImagem("path") = strCaminhoImagem
            drImagem("image") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
            dtImagem.Rows.Add(drImagem)

            dsRecibo.Tables.Add(dtImagem)

            '***************************************************************
            Dim crpt As New ReportDocument()

            Try
                crpt.FileName = Server.MapPath("~/Reports/Cr_ReciboPagar.rpt")
                crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                'Dim NomeArquivo2 As String = "ajax/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo2 As String = "files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                Dim arquivo As String = NomeArquivo

                crpt.SetDataSource(dsRecibo)

                Dim crparametervalues As ParameterValues
                Dim crparameterdiscretevalue As ParameterDiscreteValue
                Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

                crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                crparameterfielddefinition = crparameterfielddefinitions.Item("XNome")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = ENome
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                crpt.ExportToDisk(ExportFormatType.PortableDocFormat, NomeArquivo)

                If System.IO.File.Exists(NomeArquivo) Then Funcoes.AbrirArquivo(Page, NomeArquivo2)

            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Finally
                crpt.Close()
                crpt.Dispose()
            End Try
            '***************************************************************
        Else
            MsgBox(Me.Page, "Titulo não foi selecionado para emissão.")
        End If
    End Sub

    Protected Sub lnkSlip_Click(sender As Object, e As EventArgs) Handles lnkSlip.Click
        '' rotina de imprimir slip por registros checados. .'.
        Dim xextenso As String = ""
        Dim yextenso As String = ""
        Dim dsEmitir As New DataSet
        'Dim dtEmitir As DataTable
        Dim row As DataRow
        'Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim mes As Integer
        Dim MESX As String = ""
        'Dim NumeroDoCheque As Integer
        Dim NumeroDoChequeX As String = ""
        Dim Historico As String = ""
        Dim RegistroI As String = ""
        Dim RegistroS As String = ""
        Dim SqlArray As New ArrayList
        Dim Sqlupdate As String = ""
        Dim Campo() As String = Nothing
        Dim Mensagem As String = ""
        ''*Definicao de campos da empresa
        Dim Empresa As String = ""
        Dim EndEmpresa As String = ""
        Dim ENome As String = ""
        Dim EEndereco As String = ""
        Dim ECep As String = ""
        Dim ECidade As String = ""
        Dim EEstado As String = ""
        Dim ECnpj As String = ""
        Dim EInscricao As String = ""
        Dim Efone As String = ""
        Dim EBairro As String = ""
        Dim EComplemento As String = ""
        Dim ENumero As Integer
        ''*Definicao de campos da empresa Fim 
        ''* DAdos do cliente inicio
        Dim Cliente As String = ""
        Dim EndCliente As String = ""
        Dim CNome As String = ""
        Dim CEndereco As String = ""
        Dim CCep As String = ""
        Dim CCidade As String = ""
        Dim CEstado As String = ""
        Dim CCnpj As String = ""
        Dim CInscricao As String = ""
        Dim Cfone As String = ""
        Dim CBairro As String = ""
        Dim CComplemento As String = ""
        Dim CNumero As Integer

        Dim CBancoCliente As String = ""
        Dim CAgenciaCliente As String = ""
        Dim CDigitoAgenciaCliente As String = ""
        Dim CCcontaCliente As String = ""
        Dim CDigitoContaCliente As String = ""

        ''* Dados do Cliente fim 
        ''* Campos do DAta Set + titulos inicio
        'Dim myDataRow As DataRow
        '' Cria data Set que vai ser utilizado no relatorio 
        Dim dsRecibo As New DataSet
        Dim dtRecibo As DataTable
        Dim TnumeroDoCheque As Integer

        SqlArray.Clear()

        dtRecibo = New DataTable("ReciboAPagar")
        '' campos da empresa 
        dtRecibo.Columns.Add("ENome", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EEndereco", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ECep", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ECidade", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EEstado", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ECnpj", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EInscricao", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EFone", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EBairro", Type.GetType("System.String"))
        dtRecibo.Columns.Add("EComplemento", Type.GetType("System.String"))
        dtRecibo.Columns.Add("ENumero", Type.GetType("System.Int32")).DefaultValue = 0
        ' ''campos do cliente
        dtRecibo.Columns.Add("CNome", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CEndereco", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CCep", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CCidade", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CEstado", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CCnpj", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CInscricao", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CFone", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CBairro", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CComplemento", Type.GetType("System.String"))
        dtRecibo.Columns.Add("CNumero", Type.GetType("System.Int32")).DefaultValue = 0
        '' campos to titulo 
        dtRecibo.Columns.Add("TNumtit", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TValor", Type.GetType("System.Decimal"))
        dtRecibo.Columns.Add("TExtenso", Type.GetType("System.String"))
        dtRecibo.Columns.Add("THistorico", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TDia", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TMes", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TAno", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TFormaPagto", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TBanco", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TAgencia", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TConta", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TVencimento", Type.GetType("System.DateTime"))
        dtRecibo.Columns.Add("TBaixa", Type.GetType("System.DateTime"))
        dtRecibo.Columns.Add("TRecibo", Type.GetType("System.Int32")).DefaultValue = 0
        dtRecibo.Columns.Add("TRegistro", Type.GetType("System.Int32")).DefaultValue = 0
        dtRecibo.Columns.Add("TNumeroDoCheque", Type.GetType("System.Int32")).DefaultValue = 0
        '' Campos novos do cheque 05/05/2010
        dtRecibo.Columns.Add("TValorDoDocumento", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TDescontos", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TDeducoes", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TJuros", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TAcrescimos", Type.GetType("System.Decimal")).DefaultValue = 0
        dtRecibo.Columns.Add("TDigito", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TDigitoAgencia", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TContaContabilPagadora", Type.GetType("System.String"))
        dtRecibo.Columns.Add("TMovimento", Type.GetType("System.String"))

        '' Campos novos do cheque 05/05/2010
        dsRecibo.Tables.Add(dtRecibo)
        Dim ValorCobrado As Decimal
        Dim ValorDoDocumento As Decimal
        Dim Juros As Decimal
        Dim Acrescimos As Decimal
        Dim Descontos As Decimal
        Dim deducoes As Decimal
        Dim TipoPagto As Integer
        Dim FormaDePagamento As String = ""
        '' Campos novos do cheque 05/05/2010
        Dim TvalorDoDocumento As Decimal
        Dim Tdescontos As Decimal
        Dim Tdeducoes As Decimal
        Dim TJuros As Decimal
        Dim TAcrescimos As Decimal
        Dim Tdigito As String = ""
        '' Campos novos do cheque 05/05/2010
        ''* Campos do Data Set + titulos fim.
        Dim emitirSlip As Boolean

        Sql = "SELECT Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf &
              "       UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, Isnull(TipoContaCliente,'C') AS TipoContaCliente, " & vbCrLf &
              "       ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf &
              "       ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, isnull(EmpresaPedido,'') AS EmpresaPedido, isnull(EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(Pedido,0) AS Pedido, isnull(PedidoFixacao,0) AS PedidoFixacao, isnull(Procuracao,0) AS Procuracao, ValorDoDocumento, " & vbCrLf &
              "       Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, ISNULL(MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, ISNULL(MoedaDescontos, 0) " & vbCrLf &
              "       AS MoedaDescontos, ISNULL(MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(MoedaJuros, 0) AS MoedaJuros, ISNULL(MoedaAcrescimos, 0) AS MoedaAcrescimos, " & vbCrLf &
              "       ISNULL(MoedaValorLiquido, 0) AS MoedaValorLiquido, (case when len(convert(nvarchar(4000),Observacoes)) > 0 then Historico + ' ' + convert(nvarchar(4000),Observacoes ) else Historico end) as Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, " & vbCrLf &
              "       solicitacao, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, " & vbCrLf &
              "       UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, isnull(RegistroMestre, 0) as RegistroMestre, Observacoes, SituacaoBancaria,isnull(T.Descricao, 'TIPO PGTO NAO DEFINIDO') AS Descricao, ISNULL(NumeroDoCheque, 0) AS NumeroDoCheque  " & vbCrLf &
              "  FROM " & HTabela.Value & vbCrLf &
              "  LEFT JOIN TIPOSDEPAGAMENTOS AS T ON T.TipoDePagamento_ID = TipoPagto " & vbCrLf &
              " WHERE Slips = 'N'" & vbCrLf

        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            Sql &= "   and UnidadeDeNegocio = '" & Cliente & "' "
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            Sql &= "   AND Empresa = '" & Campo(0) & "'" & vbCrLf &
                   "   AND EndEmpresa = " & Campo(1) & vbCrLf
        End If

        If Left(Session("ssEmpresa").ToString, 8) = "05272759" Then
            Sql &= "   AND CASE " & vbCrLf &
                   "   	       WHEN len(ISNULL(usuariobaixa,'')) = 0" & vbCrLf &
                   "               THEN" & vbCrLf &
                   "   		           CASE " & vbCrLf &
                   "   			           WHEN LEn(ISNULL(UsuarioAlteracao,'')) = 0" & vbCrLf &
                   "   			               THEN UsuarioInclusao" & vbCrLf &
                   "   			               ELSE UsuarioAlteracao" & vbCrLf &
                   "   		           END" & vbCrLf &
                   "   	           ELSE UsuarioBaixa" & vbCrLf &
                   "        END = '" & Session("ssNomeUsuario") & "' " & vbCrLf
        End If

        Sql &= " ORDER BY Registro_Id " & vbCrLf

        Dim dsRegistro As New DataSet
        dsRegistro = Banco.ConsultaDataSet(Sql, "RegistroTitulo")

        If dsRegistro Is Nothing OrElse dsRegistro.Tables(0).Rows.Count = 0 Then
            emitirSlip = False
        Else
            emitirSlip = True

            For Each drRegistro As DataRow In dsRegistro.Tables(0).Rows
                Dim DataBaixa As String = ""
                Dim dataVencimento As String
                Dim Tmovimento As String
                Dim TContaContabilPagadora As String

                If Not drRegistro("Baixa") Is DBNull.Value Then
                    DataBaixa = CDate(drRegistro("Baixa")).ToString("dd/MM/yyyy")
                End If

                dataVencimento = CDate(drRegistro("Prorrogacao")).ToString("dd/MM/yyyy")
                Empresa = drRegistro("Empresa")
                EndEmpresa = drRegistro("EndEmpresa")
                Cliente = drRegistro("Cliente")
                EndCliente = drRegistro("EndCliente")
                ValorCobrado = 0
                ValorDoDocumento = drRegistro("ValorDoDocumento")
                Juros = drRegistro("Juros")
                Acrescimos = drRegistro("Acrescimos")
                Descontos = drRegistro("Descontos")
                deducoes = drRegistro("Deducoes")
                Historico = drRegistro("Historico")
                CBancoCliente = drRegistro("BancoCliente")
                CAgenciaCliente = drRegistro("AgenciaCliente")
                CDigitoAgenciaCliente = drRegistro("DigitoAgenciaCliente")
                CCcontaCliente = drRegistro("ContaCliente")
                CDigitoContaCliente = drRegistro("DigitoContaCliente")
                ValorCobrado = ValorDoDocumento + Juros + Acrescimos - Descontos - deducoes
                TipoPagto = drRegistro("TipoPagto")
                FormaDePagamento = drRegistro("Descricao")
                TnumeroDoCheque = drRegistro("NumeroDoCheque")
                Tdescontos = drRegistro("Descontos")
                Tdeducoes = drRegistro("Deducoes")
                TJuros = drRegistro("Juros")
                TAcrescimos = drRegistro("Acrescimos")
                TvalorDoDocumento = drRegistro("ValorDodocumento")
                Tdigito = drRegistro("DigitoContaCliente")
                Tmovimento = CDate(drRegistro("Movimento")).ToString("dd/MM/yyyy")
                TContaContabilPagadora = drRegistro("ContaContabilPagadora")

                Dim dr As DataRow
                '' Dados da empresa - fim 
                '' Consultado empresa 
                Sql = "  SELECT Emp.Cliente_Id ," & vbCrLf &
                      "         Emp.Nome, Emp.Cidade," & vbCrLf &
                      "         Emp.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf &
                      "         Emp.Endereco , Emp.Cep," & vbCrLf &
                      "         Emp.Inscricao, Emp.Telefone," & vbCrLf &
                      "         Emp.Bairro, Emp.Complemento," & vbCrLf &
                      "         Emp.Numero " & vbCrLf &
                      "    FROM Clientes Emp " & vbCrLf &
                      "   WHERE Emp.Cliente_Id = '" & Empresa & "'" & vbCrLf
                DS = Banco.ConsultaDataSet(Sql, "Clientes")
                If DS.Tables(0).Rows.Count > 0 Then
                    For Each dr In DS.Tables(0).Rows
                        ENome = dr("Nome")
                        EEndereco = dr("Endereco")
                        ECep = dr("Cep")
                        ECidade = dr("Cidade")
                        EEstado = dr("Estado")
                        ECnpj = dr("Cliente_id")
                        EInscricao = dr("Inscricao")
                        Efone = dr("Telefone")
                        EBairro = dr("Bairro")
                        EComplemento = dr("Complemento")
                        ENumero = dr("Numero")
                        Exit For
                    Next
                End If

                Sql = " SELECT Cli.Cliente_Id ," & vbCrLf &
                      "        Cli.Nome, Cli.Cidade," & vbCrLf &
                      "        Cli.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf &
                      "        Cli.Endereco , Cli.Cep," & vbCrLf &
                      "        Cli.Inscricao, Cli.Telefone," & vbCrLf &
                      "        Cli.Bairro, Cli.Complemento," & vbCrLf &
                      "        Cli.Numero " & vbCrLf &
                      "   FROM Clientes Cli " & vbCrLf &
                      "  WHERE Cli.Cliente_Id = '" & Cliente & "'" & vbCrLf &
                      "    AND Cli.Endereco_id = " & EndCliente & vbCrLf
                DS = Banco.ConsultaDataSet(Sql, "Clientes")
                If DS.Tables(0).Rows.Count > 0 Then
                    For Each dr In DS.Tables(0).Rows
                        CNome = dr("Nome")
                        CEndereco = dr("Endereco")
                        CCep = dr("Cep")
                        CCidade = dr("Cidade")
                        CEstado = dr("Estado")
                        CCnpj = dr("Cliente_id")
                        CInscricao = dr("Inscricao")
                        Cfone = dr("Telefone")
                        CBairro = dr("Bairro")
                        CComplemento = dr("Complemento")
                        CNumero = dr("Numero")
                        Exit For
                    Next
                End If

                row = dtRecibo.NewRow()
                row("ENome") = ENome
                row("EEndereco") = EEndereco
                row("ECep") = ECep
                row("ECidade") = ECidade
                row("EEstado") = EEstado
                row("ECnpj") = Funcoes.FormatarCpfCnpj(ECnpj)
                row("EInscricao") = EInscricao
                row("EFone") = Efone
                row("EBairro") = EBairro
                row("EComplemento") = EComplemento
                row("ENumero") = ENumero
                '' Move campos do Cliente / fornecedor
                row("CNome") = CNome
                row("CEndereco") = CEndereco
                row("CCep") = CCep
                row("CCidade") = CCidade
                row("CEstado") = CEstado
                row("CCnpj") = Funcoes.FormatarCpfCnpj(CCnpj)
                row("CInscricao") = CInscricao
                row("CFone") = Cfone
                row("CBairro") = CBairro
                row("CComplemento") = CComplemento
                row("CNumero") = CNumero
                '' Move campos do Titulo
                row("Tnumtit") = drRegistro("Registro_Id")
                '' feito
                row("TValor") = ValorCobrado
                row("TNumeroDoCheque") = TnumeroDoCheque
                '' campos novos 05/05/10
                row("TValorDoDocumento") = TvalorDoDocumento
                row("TDescontos") = Tdescontos
                row("TDeducoes") = Tdeducoes
                row("TJuros") = TJuros
                row("TAcrescimos") = TAcrescimos
                row("TDigito") = Tdigito
                row("TDigitoAgencia") = CDigitoAgenciaCliente

                '' campos novos 05/05/10

                Dim valcobradostr As String
                valcobradostr = CStr(ValorCobrado)
                txtValorCobrado.Text = ValorCobrado

                Valor = Replace(txtValorCobrado.Text, ".", "")

                ''* Rotina de extenso inicio
                yextenso = "("
                yextenso &= UCase(Funcoes.Extenso(txtValorCobrado.Text(), "Real", "Reais"))
                yextenso &= " *"
                xextenso = yextenso
                For j = 1 To (120 - Len(xextenso))
                    xextenso &= " *"
                Next
                xextenso &= ")"
                row("TExtenso") = xextenso
                ''* Rotina de extenso fim 
                row("THistorico") = Historico
                row("TDia") = Day(DataBaixa)
                row("TAno") = Year(DataBaixa)

                row("TMes") = Month(DataBaixa)
                ''* Rotina do Mes inicio
                mes = Month(DataBaixa)
                If mes = 1 Then
                    MESX = "JANEIRO"
                End If

                If mes = 2 Then
                    MESX = "FEVEREIRO"
                End If

                If mes = 3 Then
                    MESX = "MARCO"
                End If

                If mes = 4 Then
                    MESX = "ABRIL"
                End If

                If mes = 5 Then
                    MESX = "MAIO"
                End If

                If mes = 6 Then
                    MESX = "JUNHO"
                End If

                If mes = 7 Then
                    MESX = "JULHO"
                End If

                If mes = 8 Then
                    MESX = "AGOSTO"
                End If

                If mes = 9 Then
                    MESX = "SETEMBRO"
                End If

                If mes = 10 Then
                    MESX = "OUTUBRO"
                End If

                If mes = 11 Then
                    MESX = "NOVEMBRO"
                End If

                If mes = 12 Then
                    MESX = "DEZEMBRO"
                End If

                row("TMes") = MESX
                ''* rotina do mes fim 

                row("TFormaPagto") = FormaDePagamento
                row("TBanco") = CBancoCliente
                row("TAgencia") = CAgenciaCliente
                row("TConta") = CCcontaCliente
                row("TVencimento") = dataVencimento
                row("TBaixa") = DataBaixa

                '' calculo de registro . 
                RegistroI = drRegistro("Registro_Id")
                RegistroS = " "
                RegistroS = Trim(RegistroS)
                If Len(RegistroS) < 6 Then
                    For k = 1 To (6 - Len(Trim(RegistroI)))
                        RegistroS &= "0"
                    Next
                End If
                row("Trecibo") = Trim(RegistroS) & (Trim(RegistroI))
                row("TRegistro") = Trim(RegistroS) & (Trim(RegistroI))

                row("TContaContabilPagadora") = TContaContabilPagadora
                row("TMovimento") = Tmovimento

                dtRecibo.Rows.Add(row)

                Sql = " UPDATE " & HTabela.Value & vbCrLf &
                      "    SET Slips = 'S' " & vbCrLf &
                      "  WHERE Registro_Id = " & drRegistro("Registro_Id") & vbCrLf

                SqlArray.Add(Sql)
            Next
        End If

        If emitirSlip Then
            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, "Erro ao Salvar SLIP: " & Funcoes.EliminarCaracteresEspeciais(RTrim(Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
            Else
                Dim dtImagem As New DataTable("Images")
                dtImagem.Columns.Add("path", GetType(String))
                dtImagem.Columns.Add("image", GetType(System.Byte()))

                Dim drImagem As DataRow = dtImagem.NewRow()
                Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & Session("ssImagemEmpresa"))

                drImagem("path") = strCaminhoImagem
                drImagem("image") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
                dtImagem.Rows.Add(drImagem)

                dsRecibo.Tables.Add(dtImagem)

                Dim crpt As New ReportDocument()

                Try
                    crpt.FileName = Server.MapPath("~/Reports/Cr_slip.rpt")
                    crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                    Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                    Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                    Dim arquivo As String = NomeArquivo

                    crpt.SetDataSource(dsRecibo)

                    Dim crparametervalues As ParameterValues
                    Dim crparameterdiscretevalue As ParameterDiscreteValue
                    Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                    Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

                    crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                    crparameterfielddefinition = crparameterfielddefinitions.Item("XNome")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = ENome
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                    crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
                    ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo2 & "');", True)
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                Finally
                    crpt.Close()
                    crpt.Dispose()
                End Try
            End If
        Else
            MsgBox(Me.Page, "Sem Título(s) para emissão.")
        End If
    End Sub

    Protected Sub lnkBaixar_Click(sender As Object, e As EventArgs) Handles lnkBaixar.Click
        Limpar(True)
        Dim i As Integer = 0

        'For Each obj As [Lib].Negocio.Titulo In CType(Session("ssRetorno"), [Lib].Negocio.ListTitulo)

        While i < GridRetornoTitulos.Rows.Count
            If CType(GridRetornoTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Checked = True Then
                ' Limpar()
                txtRegistro.Text = GridRetornoTitulos.Rows(i).Cells(2).Text()
                'Consulta titulo
                ConsultaTitulo()
                DdlProvisoes.SelectedValue = 1 'marca para baixar
                lnkNovo_Click(Nothing, Nothing) 'Baixa

                'Titulos Liquidados
                Sql = " UPDATE " & HTabela.Value & vbCrLf &
                      "    SET SituacaoRemessaBancaria = 204" & vbCrLf &
                      "  WHERE Registro_id in (" & GridRetornoTitulos.Rows(i).Cells(2).Text() & ")" & vbCrLf
                SqlArray.Add(Sql)
                Banco.GravaBanco(SqlArray)
            Else

                'Titulos Rejeitados
                Sql = "  UPDATE " & HTabela.Value & vbCrLf &
                      "     SET SituacaoRemessaBancaria = 203" & vbCrLf &
                      "   WHERE Registro_id in (" & GridRetornoTitulos.Rows(i).Cells(2).Text() & ")" & vbCrLf
                SqlArray.Add(Sql)
                Banco.GravaBanco(SqlArray)
            End If
            i += 1
        End While
        'Next
    End Sub

    Protected Sub lnkImportar_Click(sender As Object, e As EventArgs) Handles lnkImportar.Click
        Try
            ImportarRetorno()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Function getDataSetAgrupamento(ByVal registro As String) As DataSet
        Dim ds As New DataSet
        Dim sql As String = "SELECT CP2.Registro_Id as Mestre_Id, " & vbCrLf &
                            "       CP2.Provisao as Mestre_Provisao, " & vbCrLf &
                            "       CP2.Vencimento as Mestre_Vencimento, " & vbCrLf &
                            "       CP2.Prorrogacao as Mestre_Prorrogacao, " & vbCrLf &
                            "       CP2.ValorDoDocumento as Mestre_ValorDoDocumento, " & vbCrLf &
                            "       CP2.Deducoes as Mestre_Deducoes, " & vbCrLf &
                            "       CP2.Acrescimos as Mestre_Acrescimos, " & vbCrLf &
                            "       CP2.ValorLiquido as Mestre_ValorLiquido, " & vbCrLf &
                            "       CP2.Cliente as Mestre_Cliente," & vbCrLf &
                            "       CLI2.Nome as Mestre_NomeCliente," & vbCrLf &
                            "       CP1.Registro_Id, " & vbCrLf &
                            "       CP1.Provisao, " & vbCrLf &
                            "       CP1.Grupado, " & vbCrLf &
                            "       CP1.Vencimento, " & vbCrLf &
                            "       CP1.Prorrogacao, " & vbCrLf &
                            "       CP1.ValorDoDocumento, " & vbCrLf &
                            "       CP1.Deducoes, " & vbCrLf &
                            "       CP1.Acrescimos, " & vbCrLf &
                            "       CP1.ValorLiquido, " & vbCrLf &
                            "       CP1.Cliente as Cliente," & vbCrLf &
                            "       CLI1.Nome as NomeCliente," & vbCrLf &
                            "       CAST(FFxI.Nota_Id as VARCHAR) + '-' + CAST(FFxI.Serie_Id as varchar) as DACTE " & vbCrLf &
                            "  FROM " & HTabela.Value & " cp1 " & vbCrLf &
                            " INNER JOIN Clientes	cli1" & vbCrLf &
                            "    ON CLI1.Cliente_Id = CP1.Cliente" & vbCrLf &
                            "   AND CLI1.Endereco_Id = CP1.EndCliente" & vbCrLf &
                            "  LEFT JOIN " & HTabela.Value & " CP2  " & vbCrLf &
                            "    ON (CP1.RegistroMestre = CP2.Registro_Id) " & vbCrLf &
                            "  LEFT JOIN Clientes cli2 " & vbCrLf &
                            "    ON CLI2.Cliente_Id = cp2.Cliente " & vbCrLf &
                            "   AND CLI2.Endereco_Id = cp2.EndCliente " & vbCrLf &
                            "  LEFT JOIN FaturaDeFreteXTitulo FFxT" & vbCrLf &
                            "    ON CP1.Registro_Id = FFxT.Titulo_Id" & vbCrLf &
                            "  LEFT JOIN FaturasDeFretesXItens FFxI" & vbCrLf &
                            "    ON FFxT.Empresa_Id       = FFxI.EmpresaPagadora_Id " & vbCrLf &
                            "   AND FFxT.EndEmpresa_Id    = FFxI.EndEmpresaPagadora_Id" & vbCrLf &
                            "   AND FFxT.Conveniado_Id    = FFxI.Conveniado_Id" & vbCrLf &
                            "   AND FFxT.EndConveniado_Id = FFxI.EndConveniado_Id" & vbCrLf &
                            "   AND FFxT.Fatura_Id        = FFxI.Fatura_Id" & vbCrLf &
                            " WHERE 1=1 " & vbCrLf &
                            "   AND CP1.Situacao = 1 " & vbCrLf &
                            "   AND CP1.Grupado <> 'M' " & vbCrLf &
                            "   AND CP1.Provisao = 1 " & vbCrLf &
                            "   AND CP2.Registro_Id IS NOT NULL " & vbCrLf

        If Not String.IsNullOrWhiteSpace(registro) Then
            sql &= "AND cp1.Registro_Id in (SELECT Registro_Id FROM " & HTabela.Value & " WHERE 1=1 AND RegistroMestre  = '" & registro & "')"
        End If

        ds = Banco.ConsultaDataSet(sql, "ContasAPagar")
        Return ds
    End Function

    Private Function getCabecalho() As String
        Dim cab As String = String.Empty

        If Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
            Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0), DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(1))
            cab = objEmpresa.Nome & "(" & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo) & ")" & vbCrLf &
            objEmpresa.Endereco & ", " & objEmpresa.Numero & vbCrLf &
            objEmpresa.CEP & " - " & objEmpresa.Cidade & "-" & objEmpresa.CodigoEstado
        End If
        Return cab
    End Function

    Private Sub viewAgrupamento(ByVal registroMestre As String)
        Try
            If Funcoes.VerificaPermissao(HTabela.Value, "RELATORIO") Then
                Dim parameters As New Dictionary(Of String, Object)

                parameters.Add("Titulo", "Relatório Agrupamento de Contas À " & IIf(HTabela.Value = "ContasAPagar", "Pagar", "Receber") & ".")

                'Funcoes.BindReport(Me.Page, getDataSetAgrupamento(registroMestre), "Cr_" & HTabela.Value, eExportType.PDF, parameters)
                Funcoes.BindReport(Me.Page, getDataSetAgrupamento(registroMestre), "Cr_ContasAPagar", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para visualizar o relatório!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ConfigurarContas(CodigoEmpresa As String, CodigoEndEmpresa As Integer, CodigoMoeda As Integer, CodigoCarteira As String)
        Dim Empresa As New Cliente(CodigoEmpresa, CodigoEndEmpresa)
        Dim Moeda As New Moeda(CodigoMoeda)

        lblDeducoes.Parent.Visible = False
        lblDescontos.Parent.Visible = False
        lblJuros.Parent.Visible = False
        lblAcrescimos.Parent.Visible = False

        If Not String.IsNullOrWhiteSpace(CodigoCarteira) Then
            Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(CodigoCarteira)

            If Not String.IsNullOrWhiteSpace(objCarteira.CodigoContaDeducao) Then
                If Empresa.Empresa.CodigoContaVariacaoMonetariaAtiva = objCarteira.CodigoContaDeducao Or Empresa.Empresa.CodigoContaVariacaoMonetariaPassiva = objCarteira.CodigoContaDeducao Then
                    If Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                        lblDeducoes.Parent.Visible = True
                        txtDeducoesMoeda.Text = "0"
                        txtDeducoesMoeda.Visible = False
                        lblDeducoes.Text = objCarteira.ContaDeducao.Titulo
                    End If
                Else
                    lblDeducoes.Parent.Visible = True
                    txtDeducoesMoeda.Visible = True
                    lblDeducoes.Text = objCarteira.ContaDeducao.Titulo
                End If
            End If

            If Not String.IsNullOrWhiteSpace(objCarteira.CodigoContaDesconto) Then
                If Empresa.Empresa.CodigoContaVariacaoMonetariaAtiva = objCarteira.CodigoContaDesconto Or Empresa.Empresa.CodigoContaVariacaoMonetariaPassiva = objCarteira.CodigoContaDesconto Then
                    If Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                        lblDescontos.Parent.Visible = True
                        txtDescontosMoeda.Text = 0
                        txtDescontosMoeda.Visible = False
                        lblDescontos.Text = objCarteira.ContaDesconto.Titulo
                    End If
                Else
                    lblDescontos.Parent.Visible = True
                    txtDescontosMoeda.Visible = True
                    lblDescontos.Text = objCarteira.ContaDesconto.Titulo
                End If
            End If

            If Not String.IsNullOrWhiteSpace(objCarteira.CodigoContaJuro) Then
                If Empresa.Empresa.CodigoContaVariacaoMonetariaAtiva = objCarteira.CodigoContaJuro Or Empresa.Empresa.CodigoContaVariacaoMonetariaPassiva = objCarteira.CodigoContaJuro Then
                    If Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                        lblJuros.Parent.Visible = True
                        txtJurosMoeda.Text = 0
                        txtJurosMoeda.Visible = False
                        lblJuros.Text = objCarteira.ContaJuro.Titulo
                    End If
                Else
                    lblJuros.Parent.Visible = True
                    txtJurosMoeda.Visible = True
                    lblJuros.Text = objCarteira.ContaJuro.Titulo
                End If
            End If

            If Not String.IsNullOrWhiteSpace(objCarteira.CodigoContaAcrescimo) Then
                If Empresa.Empresa.CodigoContaVariacaoMonetariaAtiva = objCarteira.CodigoContaAcrescimo Or Empresa.Empresa.CodigoContaVariacaoMonetariaPassiva = objCarteira.CodigoContaAcrescimo Then
                    If Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                        lblAcrescimos.Parent.Visible = True
                        txtAcrescimosMoeda.Text = "0"
                        txtAcrescimosMoeda.Visible = False
                        lblAcrescimos.Text = objCarteira.ContaAcrescimo.Titulo
                    End If
                Else
                    lblAcrescimos.Parent.Visible = True
                    txtAcrescimosMoeda.Visible = True
                    lblAcrescimos.Text = objCarteira.ContaAcrescimo.Titulo
                End If
            End If
        End If
    End Sub

    Protected Sub ddlMoeda_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlMoeda.SelectedIndexChanged
        ConfigurarContas(DdlEmpresaCliente.SelectedValue.Split("-")(0), DdlEmpresaCliente.SelectedValue.Split("-")(1), ddlMoeda.SelectedValue, ddlCarteiras.SelectedValue)
    End Sub

    Private Function getDataSet(ByRef Parametros As String) As DataSet
        Dim ds As New DataSet

        If Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.Text) Then
            Parametros &= "Empresa: " & DdlEmpresaConsultaTitulos.SelectedItem.Text & vbCrLf
        End If

        If txtCodigoClienteConsultaEndosso.Value.Length > 0 Then
            Parametros &= "Cliente " & IIf(chkClientes.Checked, " - Consolidado:", ":") & txtClienteConsultaEndosso.Text & vbCrLf
        ElseIf txtCodigoClienteConsulta.Value.Length > 0 Then
            Parametros &= "Cliente " & IIf(chkClientes.Checked, " - Consolidado:", ":") & txtClienteConsulta.Text & vbCrLf
        End If

        Sql = "SELECT Empresas.Reduzido AS ReduzidoEmpresa, Titulos.Empresa, Titulos.EndEmpresa, Empresas.Nome AS NomeEmpresa" & vbCrLf &
              "       ,Empresas.Nome AS NomeEmpresa, Empresas.Cidade AS CidadeEmpresa, Empresas.Estado AS EstadoEmpresa, " & vbCrLf &
              "       convert(nvarchar,Titulos.Registro_Id) + ' ' +" & vbCrLf &
              "       Case" & vbCrLf &
              "         when isnull(Titulos.Moeda,1) = 1 " & vbCrLf &
              "           then 'R$'" & vbCrLf &
              "           Else 'U$'" & vbCrLf &
              "       end Registro," & vbCrLf &
              "       '' as Faturamento," & vbCrLf &
              "       '' as Lote, " & vbCrLf &
              "       0 as LoteTotal," & vbCrLf &
              "       0 as LoteEntregue," & vbCrLf &
              "       Titulos.Pedido," & vbCrLf &
              "       Titulos.Cliente, " & vbCrLf &
              "         CASE " & vbCrLf &
              "             when ISNULL(ed.Codigo_Id,0) > 0" & vbCrLf &
              "                then CliED.Nome" & vbCrLf &
              "                else Clientes.Nome" & vbCrLf &
              "             end AS NomeCliente, " & vbCrLf &
              "       Titulos.Movimento, " & vbCrLf &
              "       Titulos.Vencimento AS VencimentoOriginal, " & vbCrLf &
              "         CASE " & vbCrLf &
              "             when LEN(ISNULL(ed.ClienteEndosso_Id,'')) > 0" & vbCrLf &
              "                then ed.Vencimento" & vbCrLf &
              "                else Titulos.Prorrogacao" & vbCrLf &
              "             end AS Vencimento, " & vbCrLf &
              "       Titulos.Baixa, " & vbCrLf &
              "       Titulos.Carteira, " & vbCrLf &
              "       Titulos.Provisao, " & vbCrLf &
              "       Carteira.Descricao AS NomeCarteira, " & vbCrLf &
              "       Titulos.Historico" & IIf(chkObservacao.Checked, "+ ' / OBS: ' + cast(Titulos.Observacoes as varchar) as Historico,", ",") & vbCrLf &
              "       Titulos.solicitacao, " & vbCrLf &
              "       Titulos.ValorLiquido, " & vbCrLf &
              "       Titulos.MoedaValorLiquido, " & vbCrLf &
              "       (SELECT ISNULL(P.PedidoEfetivo,0) FROM Pedidos P WHERE P.Pedido_id = Titulos.Pedido AND P.Empresa_Id = Titulos.Empresa AND P.EndEmpresa_Id = Titulos.EndEmpresa) PedidoEfetivo, " & vbCrLf &
              "       '" & IIf(HTabela.Value = "ContasAReceber", "R", "P") & "' AS Tipo" & vbCrLf &
              "  FROM " & HTabela.Value & " AS Titulos" & vbCrLf &
              "  LEFT Join cotacoes" & vbCrLf &
              "    ON Cotacoes.Data_id      = Titulos.Prorrogacao " & vbCrLf &
              "   AND Cotacoes.Indexador_Id = Titulos.Indexador " & vbCrLf &
              " INNER JOIN Clientes AS Empresas " & vbCrLf &
              "    ON Titulos.Empresa    = Empresas.Cliente_Id " & vbCrLf &
              "   AND Titulos.EndEmpresa = Empresas.Endereco_Id " & vbCrLf &
              " INNER JOIN Clientes " & vbCrLf &
              "    ON Titulos.Cliente    = Clientes.Cliente_Id " & vbCrLf &
              "   AND Titulos.EndCliente = Clientes.Endereco_Id " & vbCrLf &
              "  LEFT OUTER JOIN ComprasXProdutos AS Carteiras " & vbCrLf &
              "    ON Titulos.Carteira = Carteiras.Produto_Id" & vbCrLf &
              "  LEFT OUTER JOIN Carteira " & vbCrLf &
              "    ON Titulos.CarteiraDoTitulo = Carteira.Carteira_Id " & vbCrLf &
              "   AND (Titulos.Situacao = 1) " & vbCrLf &
              "   AND (Titulos.Grupado <> 'M') " & vbCrLf &
              "   AND (Titulos.ContaContabilCliente NOT LIKE '1010101%') " & vbCrLf &
              "   AND (Titulos.ContaContabilCliente NOT LIKE '1010102%') " & vbCrLf &
              "  LEFT JOIN NotaFiscalXTitulo NFXT" & vbCrLf &
              "    ON Titulos.Empresa      = NFXT.Empresa_Id" & vbCrLf &
              "   AND Titulos.EndEmpresa  = NFXT.EndEmpresa_Id" & vbCrLf &
              "   AND Titulos.Registro_Id = NFXT.Titulo_Id" & vbCrLf &
              "  LEFT JOIN FaturaDeFreteXTitulo FFxT" & vbCrLf &
              "    ON Titulos.Registro_Id = FFxT.Titulo_Id" & vbCrLf &
              "  LEFT JOIN FaturasDeFretesXItens FFxI" & vbCrLf &
              "    ON FFxT.Empresa_Id       = FFxI.EmpresaPagadora_Id " & vbCrLf &
              "   AND FFxT.EndEmpresa_Id    = FFxI.EndEmpresaPagadora_Id" & vbCrLf &
              "   AND FFxT.Conveniado_Id    = FFxI.Conveniado_Id" & vbCrLf &
              "   AND FFxT.EndConveniado_Id = FFxI.EndConveniado_Id" & vbCrLf &
              "   AND FFxT.Fatura_Id        = FFxI.Fatura_Id" & vbCrLf &
              "  LEFT JOIN EndossoXTitulo eTxT" & vbCrLf &
              "    ON eTxT.Empresa_Id    = Titulos.Empresa" & vbCrLf &
              "   AND eTxT.EndEmpresa_Id = Titulos.EndEmpresa" & vbCrLf &
              "   AND eTxT.Titulo_Id     = Titulos.Registro_Id" & vbCrLf &
              "  LEFT JOIN Endosso ed" & vbCrLf &
              "    ON ed.Empresa_Id    = eTxT.Empresa_Id" & vbCrLf &
              "   AND ed.EndEmpresa_Id = eTxT.EndEmpresa_Id" & vbCrLf &
              "   AND ed.Codigo_Id     = eTxT.Codigo_Id" & vbCrLf &
              "  LEFT JOIN Clientes CliED" & vbCrLf &
              "    ON CliED.Cliente_Id  = ed.ClienteEndosso_Id" & vbCrLf &
              "   AND CliED.Endereco_Id = ed.EndClienteEndosso_Id" & vbCrLf
        Sql &= " WHERE 1 = 1" & vbCrLf

        If RbAtivo.Checked Then
            If chkPrevisao.Visible AndAlso chkProvisao.Visible AndAlso chkBaixado.Visible Then
                If chkPrevisao.Checked AndAlso chkProvisao.Checked AndAlso chkBaixado.Checked Then
                    Sql &= "   AND (Titulos.Provisao = 1 OR Titulos.Provisao = 2 OR Titulos.Provisao = 3) " & vbCrLf
                    Parametros &= "Titulos Baixas, Previsionados, Provisionados" & vbCrLf
                ElseIf chkPrevisao.Checked And chkProvisao.Checked Then
                    Sql &= "   AND (Titulos.Provisao = 2 OR Titulos.Provisao = 3) " & vbCrLf
                    Parametros &= "Titulos Previsionados, Provisionados" & vbCrLf
                ElseIf chkPrevisao.Checked AndAlso chkBaixado.Checked Then
                    Sql &= "   AND (Titulos.Provisao = 2 OR Titulos.Provisao = 1) " & vbCrLf
                    Parametros &= "Titulos Previsionados, Baixados" & vbCrLf
                ElseIf chkProvisao.Checked AndAlso chkBaixado.Checked Then
                    Sql &= "   AND (Titulos.Provisao = 3 OR Titulos.Provisao = 1) " & vbCrLf
                    Parametros &= "Titulos Provisionados, Baixados" & vbCrLf
                ElseIf chkPrevisao.Checked Then
                    Sql &= "   AND ((Titulos.Provisao = 2) or (Titulos.Provisao = 3 and isnull(Titulos.PedidoFixacao,0) > 0 and isnull(Titulos.Pedido,0) <> isnull(Titulos.PedidoFixacao,0))) " & vbCrLf
                    Parametros &= "Titulos Previsionados" & vbCrLf
                ElseIf chkProvisao.Checked Then
                    Sql &= "   AND Titulos.Provisao = 3 " & vbCrLf
                    Parametros &= "Titulos Provisionados" & vbCrLf
                ElseIf chkBaixado.Checked Then
                    Sql &= "   AND Titulos.Provisao = 1 " & vbCrLf
                    Parametros &= "Titulos Baixados" & vbCrLf
                Else
                    Sql &= "   AND ((Titulos.Provisao = 2) or (Titulos.Provisao = 3 and isnull(Titulos.PedidoFixacao,0) > 0 and isnull(Titulos.Pedido,0) <> isnull(Titulos.PedidoFixacao,0))) " & vbCrLf
                    Parametros &= "Titulos Previsionados" & vbCrLf
                End If
            Else
                Sql &= "   AND ((Titulos.Provisao = 2) or (Titulos.Provisao = 3 and isnull(Titulos.PedidoFixacao,0) > 0 and isnull(Titulos.Pedido,0) <> isnull(Titulos.PedidoFixacao,0))) " & vbCrLf
                Parametros &= "Titulos Previsionados" & vbCrLf
            End If
            Parametros &= "Titulos Abertos" & vbCrLf
        End If

        If RbCancelado.Checked Then
            Sql &= "   AND Titulos.Situacao = 3 " & vbCrLf
            Parametros &= "Titulos Cancelados" & vbCrLf
        ElseIf RbAtivo.Checked Then
            If Funcoes.VerificaPermissao("TituloDeFuncionario", "LEITURA") Then
                Sql &= "   AND Titulos.Situacao in(1,101,102,105) " & vbCrLf
            Else
                Sql &= "   AND Titulos.Situacao in(1,101,102) " & vbCrLf
            End If

            Parametros &= "Titulos Ativos" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlUnidadeConsultaTitulos.SelectedValue) Then
            Sql &= "   AND Titulos.UnidadeDeNegocio = '" & DdlUnidadeConsultaTitulos.SelectedValue & "' " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
            Sql &= "   AND Titulos.Empresa = '" & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                   "   AND Titulos.EndEmpresa = " & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoClienteConsultaEndosso.Value) Then
            Sql &= "   AND ed.ClienteEndosso_Id    = '" & txtCodigoClienteConsultaEndosso.Value.Split("-")(0) & "'" & vbCrLf &
                   "   AND ed.EndClienteEndosso_Id = " & txtCodigoClienteConsultaEndosso.Value.Split("-")(1) & vbCrLf
        ElseIf Not String.IsNullOrWhiteSpace(txtCodigoClienteConsulta.Value) Then
            If chkClientes.Checked Then
                Sql &= "   AND left(Titulos.Cliente, 8) = '" & Left(txtCodigoClienteConsulta.Value.Split("-")(0), 8) & "'" & vbCrLf
            Else
                Sql &= "   AND Titulos.Cliente = '" & txtCodigoClienteConsulta.Value.Split("-")(0) & "'" & vbCrLf &
                       "   AND Titulos.EndCliente = " & txtCodigoClienteConsulta.Value.Split("-")(1) & vbCrLf
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtNumNota.Text) Then
            Sql &= "   AND (ISNULL(NFXT.Nota_Id,0) in(" & txtNumNota.Text & ") OR ISNULL(FFxI.Fatura_Id,0) in(" & txtNumNota.Text & ") OR ISNULL(FFxI.Nota_Id,0) in(" & txtNumNota.Text & "))" & vbCrLf
        ElseIf Not String.IsNullOrWhiteSpace(txtCodigoClienteConsultaEndosso.Value) Then
            Sql &= "   AND ed.Vencimento BETWEEN '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf
            Parametros &= "Vecimentos entre " & txtPeriodoInicialConsultaTitulos.Text & " a " & txtPeriodoFinalConsultaTitulos.Text & vbCrLf
        ElseIf chkClienteEndosso.Checked AndAlso String.IsNullOrWhiteSpace(txtCodigoClienteConsultaEndosso.Value) Then
            Sql &= "   AND ((ed.Vencimento BETWEEN '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "') OR " & vbCrLf
            Sql &= "        (Titulos.Prorrogacao BETWEEN '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'))" & vbCrLf
            Parametros &= "Vecimentos entre " & txtPeriodoInicialConsultaTitulos.Text & " a " & txtPeriodoFinalConsultaTitulos.Text & vbCrLf
        ElseIf IsDate(txtPeriodoInicialConsultaTitulos.Text) AndAlso IsDate(txtPeriodoFinalConsultaTitulos.Text) Then
            Sql &= "   AND Titulos.Prorrogacao BETWEEN '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf
            Parametros &= "Vecimentos entre " & txtPeriodoInicialConsultaTitulos.Text & " a " & txtPeriodoFinalConsultaTitulos.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtPedidoConsultaTitulos.Text) Then
            Sql &= "   AND Titulos.Pedido = '" & txtPedidoConsultaTitulos.Text & "'" & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(Sql, "Titulos")

        Return ds
    End Function

    Protected Sub lnkSituacaoBancaria_Click(sender As Object, e As EventArgs) Handles lnkSituacaoBancaria.Click
        Try
            If String.IsNullOrWhiteSpace(txtRegistro.Text) OrElse txtRegistro.Text = 0 Then
                MsgBox(Me.Page, "Informe o número do título a ser excluído.")
            ElseIf Not String.IsNullOrWhiteSpace(txtRegistro.Text) AndAlso String.IsNullOrWhiteSpace(txtValorCobrado.Text) Then
                MsgBox(Me.Page, "Faça a consulta do título antes de executar a exclusão!")
            Else
                Dim tituloOrigem As Titulo = New Titulo(txtRegistro.Text)

                If tituloOrigem.Codigo = 0 Then
                    MsgBox(Me.Page, "Informe o número do título.")
                ElseIf Not tituloOrigem.CodigoProvisao = 2 Then
                    MsgBox(Me.Page, "Título não está em Provisão.")
                ElseIf Not tituloOrigem.CodigoSituacao = 101 Then
                    MsgBox(Me.Page, "Título não é de Cobrança Bancária.")
                Else
                    tituloOrigem.CodigoSituacao = 1
                    tituloOrigem.BoletoBancario = False
                    tituloOrigem.HistoricoRemessa = tituloOrigem.HistoricoRemessa & ". Título removido da Cobrança pelo Usuário " & Session("ssNomeUsuario") & " em " & CDate(Today).ToString("dd/MM/yyyy")
                    tituloOrigem.UsuarioBoletoBancario = Session("ssNomeUsuario")
                    tituloOrigem.UsuarioBoletoBancarioDate = CDate(Today).ToString("yyyy/MM/dd")
                    tituloOrigem.IUD = "U"

                    If tituloOrigem.Salvar Then
                        MsgBox(Me.Page, "Registro < " & txtRegistro.Text & " > alterado com sucesso.", eTitulo.Sucess)
                        Limpar(True)
                    Else
                        MsgBox(Me.Page, "Erro ao Salvar Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI", eTitulo.Erro)
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Private Function validaExclusao() As Boolean
        If String.IsNullOrWhiteSpace(txtRegistro.Text) OrElse txtRegistro.Text = 0 Then
            MsgBox(Me.Page, "Informe o número do título a ser excluído.")
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(txtRegistro.Text) AndAlso String.IsNullOrWhiteSpace(txtValorCobrado.Text) Then
            MsgBox(Me.Page, "Faça a consulta do título antes de executar a exclusão!")
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(txtRegistro.Text) AndAlso txtNumeroCheque.Text <> 0 Then
            MsgBox(Me.Page, "Título com emissão de cheque não pode ser excluído!")
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao(HTabela.Value, "EXCLUIR") Then
                If validaExclusao() Then
                    SessaoRecuperaTitulo()

                    Cliente = DdlEmpresaCliente.SelectedValue
                    campo = Cliente.Split("-")

                    Dim Cart As New CarteiraFinanceira(ddlCarteiras.SelectedValue)
                    SqlArray.Clear()

                    If txtRegistro.Text = "" Then
                        MsgBox(Me.Page, "Informe o número do Registro para Excluir...", eTitulo.Info)
                    ElseIf Cart.isAdiantamento And Not Cart.BaixaAdiantamento And ObjTitulo.BaixasAdiantamento.Count > 0 Then
                        MsgBox(Me.Page, "O titulo é origem de um Adiantamento o qual já possui Baixas", eTitulo.Info)
                    ElseIf txtNumeroCheque.Text > 0 Then
                        MsgBox(Me.Page, "Título com emissão de cheque não pode ser excluído...", eTitulo.Info)
                    Else
                        If (Cart.BaixaAdiantamento Or (ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento)) AndAlso (IsNumeric(txtPedido.Text) AndAlso CDec(txtPedido.Text) > 0) Then
                            Sql = " UPDATE " & HTabela.Value & " SET " & vbCrLf &
                                  "   Provisao =2" & vbCrLf &
                                  "  ,UsuarioCancelamento     ='" & Session("ssNomeUsuario") & "'" & vbCrLf &
                                  "  ,UsuarioCancelamentoData ='" & CDate(Today).ToString("yyyy/MM/dd") & "'" & vbCrLf
                        Else
                            If TemNota(txtRegistro.Text) Then
                                MsgBox(Me.Page, "Titulo vinculado com Nota Fiscal não pode ser Cancelado.", eTitulo.Info)
                                Exit Sub
                            End If

                            Sql = " UPDATE " & HTabela.Value & " SET " & vbCrLf &
                                  "   Situacao                =3" & vbCrLf &
                                  "  ,UsuarioCancelamento     ='" & Session("ssNomeUsuario") & "'" & vbCrLf &
                                  "  ,UsuarioCancelamentoData ='" & CDate(Today).ToString("yyyy/MM/dd") & "'" & vbCrLf
                        End If

                        'Usuario que Liberou Titulo 'Data da Liberação do Titulo
                        If txtLiberarTitulo.Value = "S" Then Sql &= ", UsuarioLiberacaoBloqueio = '" & Session("ssNomeUsuario") & "', UsuarioLiberacaoBloqueioDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"
                        'Usuario que Liberou Pedido 'Data da Liberação do Pedido
                        If txtLiberarPedido.Value = "S" Then Sql &= ", UsuarioLiberacaoPedido = '" & Session("ssNomeUsuario") & "', UsuarioLiberacaoPedidoDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"
                        'Usuario que Liberou Cheque 'Data da Liberação do Cheque
                        If txtLiberarCheque.Value = "S" Then Sql &= ", UsuarioLiberacaoCheque = '" & Session("ssNomeUsuario") & "', UsuarioLiberacaoChequeDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"

                        Sql &= " WHERE Registro_ID = " & txtRegistro.Text & vbCrLf
                        SqlArray.Add(Sql)


                        Sql = "DELETE Razao " & vbCrLf &
                              " WHERE Titulo = " & txtRegistro.Text & vbCrLf
                        SqlArray.Add(Sql)

                        Sql = "DELETE AdiantamentosXBaixas" & vbCrLf &
                              " WHERE Titulo = " & txtRegistro.Text & vbCrLf
                        SqlArray.Add(Sql)

                        Sql = "DELETE Adiantamentos " & vbCrLf &
                             " WHERE Titulo = " & txtRegistro.Text & vbCrLf
                        SqlArray.Add(Sql)

                        Sql = "DELETE TitulosXDesdobrarFornecedor " & vbCrLf &
                              " WHERE Registro_Id = " & txtRegistro.Text & vbCrLf

                        SqlArray.Add(Sql)
                        Sql = "SELECT Registro_id " & vbCrLf &
                              "  FROM " & HTabela.Value & vbCrLf &
                              " WHERE RegistroMestre = " & txtRegistro.Text

                        Dim temMestre As Boolean = False

                        Dim dsMestre As New DataSet
                        dsMestre = Banco.ConsultaDataSet(Sql, "Registros")

                        If Not dsMestre Is Nothing AndAlso dsMestre.Tables(0).Rows.Count > 0 Then

                            temMestre = True

                            For Each drFilho As DataRow In dsMestre.Tables(0).Rows
                                Sql = " UPDATE " & HTabela.Value & vbCrLf &
                                      "    SET Situacao = 1" & vbCrLf &
                                      "       ,Provisao = 2" & vbCrLf &
                                      "       ,Grupado = 'N'" & vbCrLf &
                                      "       ,RegistroMestre = 0" & vbCrLf &
                                      "       ,UsuarioCancelamento = '" & Session("ssNomeUsuario") & "'" & vbCrLf &
                                      "       ,UsuarioCancelamentoData = '" & CDate(Today).ToString("yyyy/MM/dd") & "'" & vbCrLf &
                                      "       ,ObservacoesControleInterno = 'PROCESSO DE EXCLUSAO DO AGRUPAMENTO EM " & CDate(Today).ToString & " POR " & Session("ssNomeUsuario") & "'" & vbCrLf

                                If txtLiberarTitulo.Value = "S" Then
                                    Sql &= "       , UsuarioLiberacaoBloqueio = '" & Session("ssNomeUsuario") & "'" 'Usuario que Liberou Titulo
                                    Sql &= "       , UsuarioLiberacaoBloqueioDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"       'Data da Liberação do Titulo
                                End If

                                If txtLiberarPedido.Value = "S" Then
                                    Sql &= "       , UsuarioLiberacaoPedido = '" & Session("ssNomeUsuario") & "'"   'Usuario que Liberou Pedido
                                    Sql &= "       , UsuarioLiberacaoPedidoDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"         'Data da Liberação do Pedido
                                End If

                                If txtLiberarCheque.Value = "S" Then
                                    Sql &= "       , UsuarioLiberacaoCheque = '" & Session("ssNomeUsuario") & "'"   'Usuario que Liberou Cheque
                                    Sql &= "       , UsuarioLiberacaoChequeDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"         'Data da Liberação do Cheque
                                End If

                                Sql &= "  WHERE Registro_Id = " & drFilho("Registro_id")

                                SqlArray.Add(Sql)

                                Sql = "DELETE Razao " & vbCrLf &
                                      " WHERE Titulo = " & drFilho("Registro_id") & vbCrLf

                                SqlArray.Add(Sql)

                                Sql = "DELETE AdiantamentosXBaixas" & vbCrLf &
                                      " WHERE Titulo = " & drFilho("Registro_id") & vbCrLf

                                SqlArray.Add(Sql)

                                Sql = "DELETE Adiantamentos" & vbCrLf &
                                      " WHERE Titulo = " & drFilho("Registro_id") & vbCrLf

                                SqlArray.Add(Sql)
                            Next
                        End If

                        If Funcoes.VerificaAcesso(campo(0), campo(1), txtMovimento.Text, "Financeiro") = False Then
                            If temMestre Then
                                If DdlProvisoes.SelectedValue = 1 Then
                                    MsgBox(Me.Page, "Movimento já Fechado para esta data " & txtMovimento.Text & ", para empresa " & campo(0), eTitulo.Info)
                                    Exit Sub
                                End If
                            Else
                                MsgBox(Me.Page, "Movimento já Fechado para esta data " & txtMovimento.Text & ", para empresa " & campo(0), eTitulo.Info)
                                Exit Sub
                            End If
                        End If

                        If Banco.GravaBanco(SqlArray) = False Then
                            MsgBox(Me.Page, "Erro ao Salvar Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI", eTitulo.Erro)
                        Else
                            MsgBox(Me.Page, "Registro < " & txtRegistro.Text & " > Excluido com sucesso.", eTitulo.Sucess)
                            Limpar(True)
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridAdiantamentosDisponiveis_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridAdiantamentosDisponiveis.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                SessaoRecuperaTitulo()

                Dim txtVlrABaixarOficial As TextBox = e.Row.FindControl("txtVlrBaixaOficial")
                txtVlrABaixarOficial.Visible = ObjTitulo.Adiantamentos(e.Row.RowIndex).Elegivel
                Dim txtVlrABaixarMoeda As TextBox = e.Row.FindControl("txtVlrBaixaMoeda")
                txtVlrABaixarMoeda.Visible = ObjTitulo.Adiantamentos(e.Row.RowIndex).Elegivel
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub FifoAdiantamentos()
        SessaoRecuperaTitulo()
        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)

        If Not objCarteira.BaixaAdiantamento AndAlso (ObjTitulo.ContaContabilPagadora.Length = 0 OrElse (ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso Not ObjTitulo.ObjContaContabilPagadora.Adiantamento)) Then
            Exit Sub
        End If

        Dim moeda As New Moeda(ddlMoeda.SelectedValue)
        Dim Valor As Decimal

        If objCarteira IsNot Nothing And objCarteira.BaixaAdiantamento Then
            If moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                Valor = txtValorEmMoeda.Text
                ObjTitulo.IndiceTitulo = CDec(txtValorDoDocumento.Text) / CDec(txtValorEmMoeda.Text)
            Else
                Valor = txtValorDoDocumento.Text
            End If
        ElseIf ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento Then
            If moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                Valor = txtValorCobradoMoeda.Text
                ObjTitulo.IndiceTitulo = CDec(txtValorCobrado.Text) / CDec(txtValorCobradoMoeda.Text)
            Else
                Valor = txtValorCobrado.Text
            End If
        Else
            If IsNumeric(txtValorEmMoeda.Text) AndAlso CDec(txtValorEmMoeda.Text) > 0 Then
                ObjTitulo.IndiceTitulo = CDec(txtValorDoDocumento.Text) / CDec(txtValorEmMoeda.Text)
            End If
        End If

        Dim preenche As Boolean = False

        If ObjTitulo.Carteira.BaixaAdiantamento AndAlso Left(ObjTitulo.Carteira.CodigoContaCliente, 1) = IIf(HTabela.Value = "ContasAPagar", "2", "1") Then
            preenche = True
        ElseIf ObjTitulo.Carteira.BaixaAdiantamento AndAlso Left(ObjTitulo.Carteira.CodigoContaCliente, 1) = IIf(HTabela.Value = "ContasAPagar", "2", "1") Then
            preenche = True
        ElseIf ObjTitulo.ObjContaContabilPagadora.Adiantamento AndAlso ObjTitulo.Adiantamentos.Count > 0 Then
            preenche = True
        End If

        'If Left(ObjTitulo.ContaContabilPagadora, 1) = IIf(HTabela.Value = "ContasAPagar", "2", "1") Then
        '    'Continua sem fazer nada, está dando adiantamento ao invés de baixar - Furlan - 09/05/2020
        'Else
        If preenche Then

            If ObjTitulo.Adiantamentos.Select(Function(s) s.VlrSaldo).Sum() < Valor Then
                MsgBox(Me.Page, "O Saldo dos Adiantamentos é insuficiente para a liquidaçao do Titulo.")
                lnkNovo.Parent.Visible = False
                BloquearValores(ImgValores.ImageUrl.Contains("Bloquear"))
                Exit Sub
            End If

            For Each row In ObjTitulo.Adiantamentos
                If row.Elegivel Then
                    If Valor > row.VlrSaldo Then
                        If moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                            row.VlrABaixarMoeda = row.VlrSaldo
                            row.VlrABaixarOficial = Math.Round(row.VlrABaixarMoeda * ObjTitulo.IndiceTitulo, 2)
                            Valor -= row.VlrSaldo
                        Else
                            row.VlrABaixarOficial = row.VlrSaldo
                            Valor -= row.VlrSaldo
                        End If
                    Else
                        If moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                            row.VlrABaixarMoeda = Valor
                            row.VlrABaixarOficial = Math.Round(row.VlrABaixarMoeda * ObjTitulo.IndiceTitulo, 2)
                        Else
                            row.VlrABaixarOficial = Valor
                        End If

                        SessaoSalvaTitulo()
                        gridAdiantamentosDisponiveis.DataSource = ObjTitulo.Adiantamentos
                        gridAdiantamentosDisponiveis.DataBind()
                        Exit Sub
                    End If
                End If
            Next
        End If
        'End If

        SessaoSalvaTitulo()
    End Sub

    Protected Sub gridAdiantamentosDisponiveis_DataBinding(sender As Object, e As EventArgs) Handles gridAdiantamentosDisponiveis.DataBinding
        Dim moeda As New Moeda(ddlMoeda.SelectedValue)
        gridAdiantamentosDisponiveis.Columns(12).Visible = moeda.Classificacao = eTiposMoeda.MoedaEstrangeira
        gridAdiantamentosDisponiveis.Columns(13).Visible = moeda.Classificacao = eTiposMoeda.MoedaEstrangeira
        gridAdiantamentosDisponiveis.Columns(14).Visible = moeda.Classificacao = eTiposMoeda.MoedaEstrangeira
    End Sub

    Protected Sub txtVlrBaixaOficial_TextChanged(sender As Object, e As EventArgs)
        SessaoRecuperaTitulo()
        Dim txtOficial As TextBox = CType(sender, TextBox)
        Dim txtMoeda As TextBox = CType(txtOficial.NamingContainer, GridViewRow).FindControl("txtVlrBaixaMoeda")

        Dim index As Integer = CType(txtOficial.NamingContainer, GridViewRow).RowIndex

        If Not IsNumeric(txtOficial.Text) OrElse CDec(txtOficial.Text) < 0 Then txtOficial.Text = "0,00"

        ObjTitulo.Adiantamentos(index).VlrABaixarOficial = txtOficial.Text

        If ObjTitulo.Adiantamentos(index).VlrABaixarOficial > ObjTitulo.Adiantamentos(index).VlrSaldoOficial Then
            MsgBox(Me.Page, "Valor informando não pode ser maior que o valor do saldo do Adiantamento.")
            ObjTitulo.Adiantamentos(index).VlrABaixarOficial = 0.0
            txtOficial.Text = 0
        End If

        Dim moeda As New Moeda(ddlMoeda.SelectedValue)
        If moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then

            Dim vlrmoeda As Decimal = CDec(txtMoeda.Text)

            If ObjTitulo.IndiceTitulo > 0 Then Math.Round(ObjTitulo.Adiantamentos(index).VlrABaixarOficial / ObjTitulo.IndiceTitulo, 2)

            If Math.Abs(vlrmoeda - ObjTitulo.Adiantamentos(index).VlrABaixarMoeda) > 1 Then
                ObjTitulo.Adiantamentos(index).VlrABaixarMoeda = vlrmoeda
                'txtMoeda.Text = vlrmoeda.ToString("N2")
            End If
            gridAdiantamentosDisponiveis.DataSource = ObjTitulo.Adiantamentos
            gridAdiantamentosDisponiveis.DataBind()
        End If

        SessaoSalvaTitulo()
    End Sub

    Protected Sub txtVlrBaixaMoeda_TextChanged(sender As Object, e As EventArgs)
        SessaoRecuperaTitulo()
        Dim txtMoeda As TextBox = CType(sender, TextBox)
        Dim txtOficial As TextBox = CType(txtMoeda.NamingContainer, GridViewRow).FindControl("txtVlrBaixaOficial")

        Dim index As Integer = CType(txtMoeda.NamingContainer, GridViewRow).RowIndex

        If Not IsNumeric(txtMoeda.Text) OrElse CDec(txtMoeda.Text) < 0 Then txtMoeda.Text = "0,00"

        ObjTitulo.Adiantamentos(index).VlrABaixarMoeda = txtMoeda.Text
        Dim moeda As New Moeda(ddlMoeda.SelectedValue)
        If moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then

            Dim vlroficial As Decimal = CDec(txtOficial.Text)

            'If ObjTitulo.IndiceTitulo > 0 Then Math.Round(ObjTitulo.Adiantamentos(index).VlrABaixarMoeda * ObjTitulo.IndiceTitulo, 2)

            If Math.Abs(vlroficial - ObjTitulo.Adiantamentos(index).VlrABaixarOficial) > 1 Then
                ObjTitulo.Adiantamentos(index).VlrABaixarOficial = vlroficial
                'txtOficial.Text = vlroficial.ToString("N2")
            End If
            gridAdiantamentosDisponiveis.DataSource = ObjTitulo.Adiantamentos
            gridAdiantamentosDisponiveis.DataBind()
        End If

        SessaoSalvaTitulo()
    End Sub

    Protected Sub ImgValores_Click(sender As Object, e As ImageClickEventArgs) Handles ImgValores.Click
        Try
            If lblAgrupar.Text = "AP" AndAlso IsNumeric(txtRegistro.Text) AndAlso CInt(txtRegistro.Text) > 0 Then
                MsgBox(Me.Page, "Os Valores do Titulo Agrupador 'Mestre' não pode ser Alterado.")
                Exit Sub
            End If

            ValidaValores(True)
            BloquearValores(ImgValores.ImageUrl.Contains("Bloquear"))
            FifoAdiantamentos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtPedido_TextChanged(sender As Object, e As EventArgs) Handles txtPedido.TextChanged
        ValidaAparecimentoDaSafra()
    End Sub

    Public Sub ValidaAparecimentoDaSafra()
        Dim objCarteira As New CarteiraFinanceira(ddlCarteiras.SelectedValue)
        ddlSafraAdto.Parent.Visible = objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento
        ddlSafraAdto.Enabled = IsNumeric(txtPedido.Text) AndAlso CInt(txtPedido.Text) > 0
    End Sub

    Protected Sub btnDadosBancarios_Click(sender As Object, e As EventArgs) Handles btnDadosBancarios.Click
        Try
            SessaoRecuperaTitulo()

            If DdlProvisoes.SelectedValue = eProvisao.Baixa AndAlso IsNumeric(txtRegistro.Text) AndAlso CInt(txtRegistro.Text) > 0 Then
                MsgBox(Me.Page, "Somente títulos não baixados podem ter sua conta bancária alterada.")
                Exit Sub
            End If

            If txtCodigoFornecedor.Value.Length < 5 Then
                MsgBox(Me.Page, "Selecione um cliente para continuar.")
                Exit Sub
            End If

            If txtCodigoFornecedor.Value.Length < 5 Then
                MsgBox(Me.Page, "Selecione um cliente para continuar.")
                Exit Sub
            End If

            Dim cli() As String = txtCodigoFavorecido.Value.Split("-")
            ucConsultaDadosBancarios.SetarHID(HID.Value)
            ucConsultaDadosBancarios.CarregaGrid(cli(0), cli(1))
            Popup.ConsultaDeDadosBancarios(Me.Page, "objBancoFRMTIT" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub AtualizaDadosBancariosNoForm(ByRef pObjTitulo As TituloV)
        Try
            If pObjTitulo.CodigoBancoCliente > 0 And Not pObjTitulo.AgenciaCliente Is Nothing Then
                lblBanco.Text = pObjTitulo.CodigoBancoCliente.ToString & " - " & pObjTitulo.BancoCliente.Descricao
                lblAgencia.Text = pObjTitulo.CodigoAgenciaCliente & IIf(ObjTitulo.DigitoAgenciaCliente.Length > 0, "-", "") & pObjTitulo.DigitoAgenciaCliente & " | " & pObjTitulo.AgenciaCliente.Praca
                lblContaCorrente.Text = pObjTitulo.ContaCliente & IIf(ObjTitulo.DigitoContaCliente.Length > 0, "-", "") & pObjTitulo.DigitoContaCliente
                lblTipoConta.Text = ObjTitulo.TipoDaContaCliente
                imgContaObs.ToolTip = "Obs: " & ObjTitulo.BancoXConta.Observacoes
            Else
                lblBanco.Text = "Banco"
                lblAgencia.Text = "Agência"
                lblContaCorrente.Text = "Conta"
                lblTipoConta.Text = "Tipo Da Conta"
                imgContaObs.ToolTip = "Observações"
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlTiposDePagamentos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlTiposDePagamentos.SelectedIndexChanged
        divCheque.Visible = DdlTiposDePagamentos.SelectedValue = 2
    End Sub

    Protected Sub DdlContaPagadora_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlContaPagadora.SelectedIndexChanged
        SessaoRecuperaTitulo()

        If txtCodigoFornecedor.Value.Length = 0 Then
            MsgBox(Me.Page, "Fornecedor/Cliente não foi informado.")
            DdlContaPagadora.SelectedIndex = 0
            Exit Sub
        End If

        '*********************************************************************************************************************
        '*********************************************************************************************************************
        If DdlContaPagadora.SelectedIndex = 0 Then
            ObjTitulo.CodigoBancoPagador = 0
            ObjTitulo.CodigoAgenciaPagadora = String.Empty
            ObjTitulo.DigitoAgenciaPagadora = String.Empty
            ObjTitulo.ContaPagadora = String.Empty
            ObjTitulo.DigitoContaPagadora = String.Empty
            ObjTitulo.TipoDaContaPagadora = String.Empty
            ObjTitulo.ContaContabilPagadora = String.Empty
            SessaoSalvaTitulo()

            If ddlCarteiras.SelectedIndex > 0 Then
                Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)
                If Not (objCarteira.isAdiantamento And objCarteira.BaixaAdiantamento) Then
                    gridAdiantamentosDisponiveis.DataSource = Nothing
                    gridAdiantamentosDisponiveis.DataBind()
                    divSelecaoAdBX.Visible = False
                End If
            End If
            Exit Sub
        End If

        '*********************************************************************************************************************
        '*********************************************************************************************************************
        Dim dadosBanc As String() = DdlContaPagadora.SelectedValue.Split("-")
        ObjTitulo.CodigoBancoPagador = DdlBancoPagador.SelectedValue
        ObjTitulo.CodigoAgenciaPagadora = dadosBanc(0)
        ObjTitulo.DigitoAgenciaPagadora = dadosBanc(1)
        ObjTitulo.ContaPagadora = dadosBanc(2)
        ObjTitulo.DigitoContaPagadora = dadosBanc(3)
        ObjTitulo.TipoDaContaPagadora = dadosBanc(4)
        ObjTitulo.ContaContabilPagadora = dadosBanc(5)
        SessaoSalvaTitulo()

        If Not ObjTitulo.ContaContabilPagadora Is Nothing AndAlso ObjTitulo.ContaContabilPagadora.Length > 0 AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento AndAlso ddlCarteiras.SelectedValue.Length = 0 Then
            MsgBox(Me.Page, "Finalidade Financeira não foi informada.")
            DdlContaPagadora.SelectedIndex = 0
            Exit Sub
        End If

        If ObjTitulo.ObjContaContabilPagadora IsNot Nothing AndAlso Not ObjTitulo.ObjContaContabilPagadora.Adiantamento Then

            Dim strEmpresa() As String = DdlEmpresaCliente.SelectedValue.Split("-")
            Dim objEmpresa As New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
            Dim strCliente() As String = txtCodigoFornecedor.Value.Split("-")
            Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
            Dim Moeda As New Moeda(ddlMoeda.SelectedValue)

            Dim listadto As New ListAdiantamento(objEmpresa, objCliente, 1, Moeda.Classificacao, IIf(Not IsNumeric(txtPedido.Text), 0, txtPedido.Text), ObjTitulo.ContaContabilPagadora)
            ObjTitulo.Adiantamentos = listadto

            divSelecaoAdBX.Visible = ObjTitulo.Adiantamentos.Count > 0

            gridAdiantamentosDisponiveis.DataSource = listadto.ToArray
            gridAdiantamentosDisponiveis.DataBind()

        Else
            Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)
            If objCarteira.isAdiantamento AndAlso ObjTitulo.ObjContaContabilPagadora.Adiantamento Then
                ddlCarteiras.SelectedIndex = 0
                txtNumeroAdto.Text = ""
                MsgBox(Me.Page, "Conta de Baixa de Adiantamento não pode ser usada com Carteira Principal de Adiantamento.")
                TabContainer1.ActiveTabIndex = 0

                gridAdiantamentosDisponiveis.DataSource = Nothing
                gridAdiantamentosDisponiveis.DataBind()
                divSelecaoAdBX.Visible = False
            ElseIf Left(ObjTitulo.ContaContabilPagadora, 1) = IIf(HTabela.Value = "ContasAPagar", "2", "1") Then
                gridAdiantamentosDisponiveis.DataSource = Nothing
                gridAdiantamentosDisponiveis.DataBind()
                divSelecaoAdBX.Visible = False
            Else
                divSelecaoAdBX.Visible = True
                Dim strEmpresa() As String = DdlEmpresaCliente.SelectedValue.Split("-")
                Dim objEmpresa As New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
                Dim strCliente() As String = txtCodigoFornecedor.Value.Split("-")
                Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
                Dim Moeda As New Moeda(ddlMoeda.SelectedValue)

                Dim listadto As New ListAdiantamento(objEmpresa, objCliente, 1, Moeda.Classificacao, IIf(Not IsNumeric(txtPedido.Text), 0, txtPedido.Text), ObjTitulo.ContaContabilPagadora)
                ObjTitulo.Adiantamentos = listadto

                gridAdiantamentosDisponiveis.DataSource = listadto.ToArray
                gridAdiantamentosDisponiveis.DataBind()
            End If
        End If

        If lblAgrupar.Text = "AP" Then chkConciliado.Parent.Visible = True

        SessaoSalvaTitulo()
    End Sub

    Protected Sub imgLimparContaBancaria_Click(sender As Object, e As ImageClickEventArgs) Handles imgLimparContaBancaria.Click
        lblBanco.Text = "Banco"
        lblAgencia.Text = "Agência"
        lblContaCorrente.Text = "Conta"
        lblTipoConta.Text = "Tipo Da Conta"
        imgContaObs.ToolTip = "Observações"
    End Sub

    Protected Sub chkConciliado_CheckedChanged(sender As Object, e As EventArgs) Handles chkConciliado.CheckedChanged
        If chkConciliado.Checked Then
            txtDataConciliacao.Value = txtMovimento.Text
        Else
            txtDataConciliacao.Value = ""
        End If
    End Sub

    Private Sub lnkConsultar_Unload(sender As Object, e As EventArgs) Handles lnkConsultar.Unload

    End Sub
End Class