Imports System.IO
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class ContasAPagar
    Inherits BasePage

#Region "Atributos / Propriedades"

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

#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            'If Funcoes.VerificaPermissao("ContasAPagar", "ACESSAR") Then
            '    If Funcoes.VerificaPermissao("ContasAPagar", "LIBERAR") Then
            '        chkPrevisao.Visible = True
            '        chkProvisao.Visible = True
            '        chkBaixado.Visible = True
            '    End If

            '    CargaUnidadeDeNegocioEmpresaCliente()
            '    CargaBancos()
            '    TiposDePagamentos()
            '    BuscarMoedas()
            '    BuscarIndexadores()
            '    Provisoes()
            '    CarteiraDoTitulo()
            '    Carteiras()
            '    CarteirasAdto()
            '    Limpar(True)
            '    Limpar_ConsultaTitulos(True)
            '    txtMovimento.Text = CDate(Today).ToString("dd/MM/yyyy")
            '    hdnMovimentoOriginal.Value = CDate(Today).ToString("dd/MM/yyyy")
            '    ddl.Carregar(ddlSelecionarHist, CarregarDDL.Tabela.Historico)

            '    If Not String.IsNullOrWhiteSpace(Request.QueryString("registro")) Then
            '        Dim Registro = Funcoes.Decifrar(Request.QueryString("registro"))

            '        If IsNumeric(Registro) Then
            '            txtRegistro.Text = Registro
            '            lnkConsultar_Click(lnkConsultar, New EventArgs())
            '        End If
            '    End If
            'Else
            MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx", eTitulo.Info)
            Exit Sub
            'End If

            'TabContainer1.ActiveTabIndex = 0
        End If
    End Sub

    Protected Sub ddlSelecionarHist_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlSelecionarHist.SelectedIndexChanged
        Try
            txtHistorico.Text = ddlSelecionarHist.SelectedItem.Text
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub MostrarCotacao()
        If Not String.IsNullOrWhiteSpace(txtProrrogacao.Text) AndAlso IsDate(txtProrrogacao.Text) AndAlso Not String.IsNullOrWhiteSpace(ddlIndexador.SelectedValue) Then
            lblCotacao.Text = Funcoes.PegarValorConversao(ddlIndexador.SelectedValue, txtProrrogacao.Text).ToString("N4")
            lblDescCotacao.Text = ddlIndexador.SelectedItem.Text.Split("-")(1)
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
                ddlIndexador.Items.Add(New ListItem(objIndexador.Codigo.ToString() & "-" & objIndexador.Descricao, _
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

            Popup.ConsultaDePedidos(Me.Page, "objContasAPagar" & HID.Value, "txtNome")
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
            DdlCarteirasAdto.SelectedIndex = 0
            CargaContasCorrentes()
            Session.Remove("objFornecedorCP" & HID.Value)
        ElseIf Not Session("objFavorecidoCP" & HID.Value) Is Nothing Then
            objClient = CType(Session("objFavorecidoCP" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objClient)
            txtFavorecido.Text = itemCliente.Text
            txtCodigoFavorecido.Value = itemCliente.Value
            CargaContasCorrentes()
            Session.Remove("objFavorecidoCP" & HID.Value)
        ElseIf Session("objContasAPagar" & HID.Value) IsNot Nothing Then
            Dim p As [Lib].Negocio.Pedido = CType(Session("objContasAPagar" & HID.Value), [Lib].Negocio.Pedido)
            txtPedidoConsultaTitulos.Text = p.Codigo
            Session.Remove("objContasAPagar" & HID.Value)
        ElseIf Not Session("objClienteCTAXPC" & HID.Value) Is Nothing Then
            Dim objClienteCTAXPC As [Lib].Negocio.Cliente = CType(Session("objClienteCTAXPC" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objClienteCTAXPC)
            txtClienteConsulta.Text = itemCliente.Text
            txtCodigoClienteConsulta.Value = itemCliente.Value
            Session.Remove("objClienteCTAXPC" & HID.Value)
        ElseIf Not Session("objPedidoCTAPAG" & HID.Value) Is Nothing Then
            Cliente = DdlEmpresaCliente.SelectedValue
            campo = Cliente.Split("-")
            Dim strCliente() As String = txtCodigoFornecedor.Value.Split("-")
            objPedido = CType(Session("objPedidoCTAPAG" & HID.Value), [Lib].Negocio.Pedido)
            If Trim(txtValorEmMoeda.Text) = "" Then txtValorEmMoeda.Text = 0
            If objPedido.MomentoFinanceiro = 3 And Funcoes.VerificaPermissao("ContasAPagar", "LIBERARMOMENTOFINANCEIRO") = False And objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                MsgBox(Me.Page, "Processo Não permitido. Pedido Lançado Com Vencimentos Determinados na Emissão da Nota Fiscal.")
            ElseIf objPedido.CodigoUnidadeNegocio <> DdlUnidadeDeNegocioEmpresaCliente.SelectedValue Then
                MsgBox(Me.Page, "Unidade de Negócio da Empresa do Pedido é diferente da Unidade de Negócio da Empresa Fornecedora.")
            ElseIf objPedido.CodigoEmpresa <> campo(0) Or objPedido.EnderecoEmpresa.ToString <> campo(1) Then
                MsgBox(Me.Page, "Empresa do Pedido é diferente da Empresa Fornecedora.")
            ElseIf objPedido.CodigoCliente <> strCliente(0) Or objPedido.EnderecoCliente.ToString <> strCliente(1) Then
                MsgBox(Me.Page, "Fornecedor do Pedido é diferente do Fornecedor informado.")
            Else
                txtPedido.Text = objPedido.Codigo
                ddlMoeda.SelectedValue = objPedido.CodigoMoeda
                ddlIndexador.SelectedValue = objPedido.CodigoIndexador
                ddlMoeda.Enabled = False
                ddlIndexador.Enabled = False
                lnkExcluir.Parent.Visible = False
            End If
            Session.Remove("objPedidoCTAPAG" & HID.Value)
        End If
    End Sub

#Region "Dados Bancários"

    Private Sub CargaBancos()
        Sql = "SELECT Banco_Id as Codigo, convert(varchar,REPLICATE('0', 4 - LEN(CAST(Banco_ID AS varchar))) + CAST(Banco_Id AS varchar)) + '  -  ' + Descricao as Descricao FROM Bancos Order By Banco_Id"
        DdlBancos.DataValueField = "Codigo"
        DdlBancos.DataTextField = "Descricao"
        DdlBancos.DataSource = Banco.ConsultaDataSet(Sql, "Bancos")
        DdlBancos.DataBind()

        DdlBancos.Items.Insert(0, "")
        'DdlBancos.SelectedIndex = 0
    End Sub

    Private Sub LimparDadosBancarios()

        DdlBancos.SelectedIndex = 0
        txtAgencia.Text = ""
        txtDigitoAgencia.Text = ""
        txtContaCorrente.Text = ""
        txtDigitoDaConta.Text = ""
        txtObservacoesDaConta.Text = ""
    End Sub

    Sub CargaContasCorrentes()
        Dim strSQL As String

        Dim Campo = txtCodigoFavorecido.Value.Split("-")
        Cliente = Campo(0)
        Endereco = Campo(1)

        strSQL = " SELECT Banco_Id, Agencia_ID, DigitoAgencia_Id, ContaCorrente_Id, DigitoConta_Id, Isnull(TipoConta,'C') as TipoConta, Observacoes" & vbCrLf & _
                 "   FROM ClientesXContasBancarias" & vbCrLf & _
                 "  WHERE Cliente_Id = '" & Cliente & "' " & vbCrLf & _
                 "    AND Endereco_Id = " & Endereco & vbCrLf & _
                 "  ORDER BY Banco_Id, Agencia_ID " & vbCrLf

        GridContasCorrentes.DataSource = Banco.ConsultaDataSet(strSQL, "Contas")
        GridContasCorrentes.DataBind()
    End Sub

    Protected Sub GridContasCorrentes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridContasCorrentes.SelectedIndexChanged
        DdlBancos.SelectedIndex = DdlBancos.Items.IndexOf(DdlBancos.Items.FindByValue(GridContasCorrentes.SelectedRow.Cells(1).Text()))
        txtAgencia.Text = Server.HtmlDecode(GridContasCorrentes.SelectedRow.Cells(2).Text().Trim())
        txtDigitoAgencia.Text = Server.HtmlDecode(GridContasCorrentes.SelectedRow.Cells(3).Text().Trim())
        txtContaCorrente.Text = Server.HtmlDecode(GridContasCorrentes.SelectedRow.Cells(4).Text().Trim())
        txtDigitoDaConta.Text = Server.HtmlDecode(GridContasCorrentes.SelectedRow.Cells(5).Text().Trim())
        ddlTipoConta.SelectedIndex = ddlTipoConta.Items.IndexOf(ddlTipoConta.Items.FindByValue(GridContasCorrentes.SelectedRow.Cells(6).Text()))
        txtObservacoesDaConta.Text = Server.HtmlDecode(GridContasCorrentes.SelectedRow.Cells(7).Text().Trim())
        TabContainer1.ActiveTabIndex = 1
    End Sub

#End Region

#Region "Consulta Títulos"
    Private Sub TitulosConsulta()
        Dim Cliente As String
        Dim Campo() As String
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue

        Sql = "  SELECT CP.Registro_Id AS Registro, convert(varchar(10),CP.Prorrogacao,103) as Vencimento, " & vbCrLf & _
              "         Cli.Nome AS Cliente, Historico, isnull(CP.MoedaValorDoDocumento, 0) AS Dolar, CP.ValorDoDocumento AS Valor, " & vbCrLf & _
              "         ISNULL(CP.MoedaValorLiquido, 0) AS MoedaLiquido, CP.ValorLiquido AS ValorLiquido, " & vbCrLf & _
              "         UsuarioLiberacao as Liberado, CP.Pedido as Pedido, " & vbCrLf & _
              "         CASE " & vbCrLf & _
              "             WHEN CP.Moeda = 0 THEN 'R$-' + convert(varchar,CP.Moeda) " & vbCrLf & _
              "             ELSE " & vbCrLf & _
              "                 CASE " & vbCrLf & _
              "                     WHEN CP.Moeda = 1  THEN 'R$-' + convert(varchar,CP.Moeda) " & vbCrLf & _
              "                     ELSE 'U$-' + convert(varchar,CP.Moeda) " & vbCrLf & _
              "                 END " & vbCrLf & _
              "         END as Moeda, CP.Indexador, isnull(CP.Grupado,'N') as Grupado, CP.Provisao " & vbCrLf & _
              "    FROM ContasAPagar CP " & vbCrLf & _
              "    LEFT JOIN NotaFiscalXTitulo NFXT" & vbCrLf & _
              "      ON CP.Empresa     = NFXT.Empresa_Id" & vbCrLf & _
              "     AND CP.EndEmpresa  = NFXT.EndEmpresa_Id" & vbCrLf & _
              "     AND CP.Registro_Id = NFXT.Titulo_Id" & vbCrLf & _
              "    LEFT JOIN FaturaDeFreteXTitulo FFxT" & vbCrLf & _
              "      ON CP.Registro_Id = FFxT.Titulo_Id" & vbCrLf & _
              "    LEFT JOIN FaturasDeFretesXItens FFxI" & vbCrLf & _
              "      ON FFxT.Empresa_Id       = FFxI.EmpresaPagadora_Id " & vbCrLf & _
              "     AND FFxT.EndEmpresa_Id    = FFxI.EndEmpresaPagadora_Id" & vbCrLf & _
              "     AND FFxT.Conveniado_Id    = FFxI.Conveniado_Id" & vbCrLf & _
              "     AND FFxT.EndConveniado_Id = FFxI.EndConveniado_Id" & vbCrLf & _
              "     AND FFxT.Fatura_Id        = FFxI.Fatura_Id" & vbCrLf & _
              "   INNER JOIN Clientes Cli" & vbCrLf & _
              "      ON CP.Cliente = Cli.Cliente_Id " & vbCrLf & _
              "     AND CP.EndCliente = Cli.Endereco_Id" & vbCrLf & _
              "   WHERE 1=1 " & vbCrLf

        If RbCancelado.Checked Then
            Sql &= " AND CP.Situacao <> 1 " & vbCrLf
        ElseIf RbAtivo.Checked Then
            Sql &= " AND CP.Situacao = 1 " & vbCrLf
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

        Campo = txtCodigoClienteConsulta.Value.Split("-")
        If Campo(0) <> "" Then
            Sql &= " AND CP.Cliente = '" & Campo(0) & "'" & vbCrLf  'Cliente
            Sql &= " AND CP.EndCliente = " & Campo(1) & vbCrLf    'Cliente da Empresa
        End If

        If Not String.IsNullOrWhiteSpace(txtNumNota.Text) Then
            Sql &= " AND (ISNULL(NFXT.Nota_Id,0) in(" & txtNumNota.Text & ") OR ISNULL(FFxI.Fatura_Id,0) in(" & txtNumNota.Text & ") OR ISNULL(FFxI.Nota_Id,0) in(" & txtNumNota.Text & "))"
        ElseIf txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
            Sql &= " AND Prorrogacao between '" & CDate(txtPeriodoInicialConsultaTitulos.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtPeriodoFinalConsultaTitulos.Text).ToString("yyyy/MM/dd") & "'"
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
                    GridConsultaTitulos.Rows(i).ForeColor = Drawing.Color.Red
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
                    GridConsultaTitulos.Rows(i).Cells(3).ForeColor = Drawing.Color.Red
                    GridConsultaTitulos.Rows(i).Cells(3).ToolTip = "BAIXADO"
                End If

                If RbCancelado.Checked Then
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Enabled = False
                    CType(GridConsultaTitulos.Rows(i).FindControl("ChkLiberado"), CheckBox).Enabled = False
                    CType(GridConsultaTitulos.Rows(i).FindControl("chkReprogramar"), CheckBox).Enabled = False
                    CType(GridConsultaTitulos.Rows(i).FindControl("chkRecibo"), CheckBox).Enabled = False
                End If

                i += 1
            End While

            lnkSlip.Parent.Visible = True

            If Not RbCancelado.Checked Then
                lnkRecibo.Parent.Visible = True
                lnkReprogramar.Parent.Visible = True
            End If

            rowDolar.Visible = True

        End If
    End Sub
#End Region

#Region "Manutenção dos Títulos"

    Private Sub CargaUnidadeDeNegocioEmpresaCliente()
        ddl.Carregar(DdlUnidadeDeNegocioEmpresaCliente, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        ddl.Carregar(DdlUnidadeConsultaTitulos, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        ddl.Carregar(DdlEmpresaPagadora, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

#End Region

    Protected Sub DdlUnidadeDeNegocioEmpresaCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(DdlEmpresaCliente, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocioEmpresaCliente.SelectedValue, True)
    End Sub

    Private Sub CargaEmpresaCliente()
        ddl.Carregar(DdlEmpresaCliente, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocioEmpresaCliente.SelectedValue, True)
    End Sub

    Private Sub TiposDePagamentos()
        Sql = "SELECT TipoDePagamento_Id as Codigo, convert(varchar,REPLICATE('0', 2 - LEN(CAST(TipoDePagamento_Id AS varchar))) + CAST(TipoDePagamento_Id AS varchar)) + '  -  ' + Descricao as Descricao FROM TiposDePagamentos Order By TipoDePagamento_Id"

        DdlTiposDePagamentos.DataValueField = "Codigo"
        DdlTiposDePagamentos.DataTextField = "Descricao"
        DdlTiposDePagamentos.DataSource = Banco.ConsultaDataSet(Sql, "TiposDePagamentos")
        DdlTiposDePagamentos.DataBind()

        DdlTiposDePagamentos.Items.Insert(0, "")
        DdlTiposDePagamentos.SelectedIndex = 0
    End Sub

    Private Sub Provisoes()
        Sql = "SELECT Provisao_Id as Codigo, convert(varchar,REPLICATE('0', 1 - LEN(CAST(Provisao_Id AS varchar))) + CAST(Provisao_Id AS varchar)) + '  -  ' + Descricao as Descricao FROM Provisoes Order By Provisao_Id"

        DdlProvisoes.DataValueField = "Codigo"
        DdlProvisoes.DataTextField = "Descricao"
        DdlProvisoes.DataSource = Banco.ConsultaDataSet(Sql, "Provisoes")
        DdlProvisoes.DataBind()

        DdlProvisoes.Items.Insert(0, "")
        DdlProvisoes.SelectedIndex = 0
    End Sub

    Private Sub CarteiraDoTitulo()
        Dim objCarteiraDoTitulo As New [Lib].Negocio.ListCarteiraDoTitulo()

        ddlCarteiraDoTitulo.DataValueField = "Codigo"
        ddlCarteiraDoTitulo.DataTextField = "Descricao"
        ddlCarteiraDoTitulo.DataSource = objCarteiraDoTitulo.ToArray()
        ddlCarteiraDoTitulo.DataBind()
        ddlCarteiras.SelectedIndex = 0
    End Sub

    Private Sub Carteiras()
        Sql = "SELECT Produto_Id AS Codigo," & vbCrLf & _
              "       Produto_Id + '  -  ' + Descricao AS Descricao" & vbCrLf & _
              "  FROM ComprasXProdutos" & vbCrLf & _
              " Where Classificacao = 'P'" & vbCrLf & _
              " Order By Produto_Id" & vbCrLf

        ddlCarteiras.DataValueField = "Codigo"
        ddlCarteiras.DataTextField = "Descricao"
        ddlCarteiras.DataSource = Banco.ConsultaDataSet(Sql, "Carteiras")
        ddlCarteiras.DataBind()

        ddlCarteiras.Items.Insert(0, "")
        ddlCarteiras.SelectedIndex = 0
    End Sub

    Private Sub CarteirasAdto()
        Sql = "SELECT Produto_Id AS Codigo," & vbCrLf & _
              "       CONVERT(varchar, REPLICATE('0', 9 - LEN(CAST(Produto_Id AS varchar))) + CAST(Produto_Id AS varchar))  + '  -  ' + Descricao AS Descricao" & vbCrLf & _
              "  FROM ComprasXProdutos" & vbCrLf & _
              " Where Classificacao = 'P'" & vbCrLf & _
              "   And Adiantamento in ('S','R')" & vbCrLf & _
              "   And isnull(BaixaAdiantamento,0) = 1" & vbCrLf & _
              "  Order By Produto_Id"

        DdlCarteirasAdto.DataValueField = "Codigo"
        DdlCarteirasAdto.DataTextField = "Descricao"
        DdlCarteirasAdto.DataSource = Banco.ConsultaDataSet(Sql, "Carteiras")
        DdlCarteirasAdto.DataBind()

        DdlCarteirasAdto.Items.Insert(0, "")
        DdlCarteirasAdto.SelectedIndex = 0
    End Sub

    Private Sub CargaBancoPagador()
        DdlBancoPagador.Items.Clear()
        DdlContaPagadora.Items.Clear()

        If DdlEmpresaPagadora.Text <> "" Then

            Cliente = DdlEmpresaPagadora.SelectedValue
            campo = Cliente.Split("-")

            Sql = " SELECT DISTINCT BxC.Banco_Id,  B.Descricao" & vbCrLf & _
                  "   FROM BancosXContas BxC " & vbCrLf & _
                  "  INNER JOIN Bancos B  " & vbCrLf & _
                  "     ON BxC.Banco_Id = B.Banco_Id" & vbCrLf & _
                  "  WHERE BxC.Empresa_Id  = '" & campo(0) & "'" & vbCrLf & _
                  "    AND BxC.EndEmpresa_Id  = " & campo(1)

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Bancos").Tables(0).Rows
                Descricao = Format(Dr("Banco_Id"), "0000") & "- " & Dr("Descricao")
                DdlBancoPagador.Items.Add(New ListItem(Descricao, Dr("Banco_Id")))
            Next

            DdlBancoPagador.Items.Insert(0, "")
            DdlBancoPagador.SelectedIndex = 0
        End If
    End Sub

    Private Function ValidaCampos() As Boolean
        Cliente = DdlEmpresaCliente.SelectedValue
        campo = Cliente.Split("-")

        'Empresa Pagadora para utilizar na validação de datas não programáveis
        Dim Emp() As String = Nothing
        If DdlEmpresaPagadora.SelectedIndex > 0 Then
            Emp = DdlEmpresaPagadora.SelectedValue.Split("-")
        Else
            Emp = DdlEmpresaCliente.SelectedValue.Split("-")
        End If

        If DdlCarteirasAdto.SelectedIndex > 0 And ddlCarteiras.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Primeira Carteira.")
            Return False
        End If

        If DdlProvisoes.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Selecione Previsão / Baixa.")
            DdlProvisoes.Focus()
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
        ElseIf Not ValidaData(txtMovimento.Text, "Movimento", Emp(0), Emp(1)) Then
            MsgBox(Me.Page, Mensagem)
            Return False
        ElseIf DdlProvisoes.SelectedValue = "2" AndAlso CDate(txtProrrogacao.Text) < CDate(txtMovimento.Text) Then
            MsgBox(Me.Page, "Data de Vencimento não pode ser maior que a data de Movimento.")
            Return False
        ElseIf Not ValidaData(txtProrrogacao.Text, "Vencimento", Emp(0), Emp(1)) Then
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
        ElseIf DdlProvisoes.SelectedValue = 1 AndAlso String.IsNullOrWhiteSpace(DdlEmpresaPagadora.SelectedValue) Then
            MsgBox(Me.Page, "Empresa Pagadora não foi Selecionada.")
            Return False
        End If

        If DdlProvisoes.SelectedValue = 1 AndAlso ChkLiberado.Checked = False AndAlso Not Funcoes.VerificaPermissao("AutorizacoesDePagamentos", "GRAVAR") Then
            If ddlCarteiras.SelectedValue = "001001057" OrElse _
               ddlCarteiras.SelectedValue = "001001069" OrElse _
               ddlCarteiras.SelectedValue = "001001067" OrElse _
               ddlCarteiras.SelectedValue = "001001005" OrElse _
               ddlCarteiras.SelectedValue = "001001052" OrElse _
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
                MsgBox(Me.Page, "Empresa Pagadora é obrigatório")
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
                MsgBox(Me.Page, "Valor Pago é obrigatório")
                Return False
            ElseIf Not lblAgrupar.Text = "AP" AndAlso ddlCarteiras.SelectedIndex > 0 AndAlso DdlTributos.Items.Count > 1 AndAlso DdlTributos.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Encargo é obrigatório.")
                Return False
            End If

            Dim SelecaoBancoxConta As Array = DdlContaPagadora.SelectedValue.Split("-") 'Insol - Somente baixar pela filial conta <> caixa atravez de autorização.
            If Not SelecaoBancoxConta(4) = "101010101" And Not Funcoes.VerificaPermissao("BAIXAMATRIZ", "GRAVAR") Then
                MsgBox(Me.Page, "Usuário Sem Autorização para Baixar Registro, Somente Baixas Via Matriz São Autorizadas.")
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
        ElseIf Not String.IsNullOrWhiteSpace(txtPedido.Text) AndAlso CInt(txtPedido.Text) > 0 AndAlso ddlMoeda.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Titulo com Pedido a Moeda deve ser selecionada.")
            Return False
        ElseIf DdlProvisoes.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Previsao é obrigatório.")
            Return False
        ElseIf DdlProvisoes.SelectedValue = 3 And Not Funcoes.VerificaPermissao("ContasAPagar", "LIBERAR") Then
            MsgBox(Me.Page, "Usuário sem autorização para lançamento de Provisão")
            Return False
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

        If DdlTiposDePagamentos.Text <> "" Then
            If DdlTiposDePagamentos.SelectedValue = 3 Then
                If txtContaCorrente.Text = "" Then
                    MsgBox(Me.Page, "Dados Bancários é obrigatório.")
                    Return False
                End If
            End If
        End If

        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)

        Dim Empresa As New [Lib].Negocio.ClienteXEmpresa(campo(0), campo(1))

        If Not String.IsNullOrWhiteSpace(Empresa.CodigoContaFornecedorFrete) _
            AndAlso Not lblAgrupar.Text = "AP" AndAlso String.IsNullOrWhiteSpace(txtRegistro.Text) _
            AndAlso Empresa.CodigoContaFornecedorFrete = objCarteira.CodigoContaCliente Then
            MsgBox(Me.Page, "Titulo não pode ser incluído na Conta de Fornecedor de Frete.")
            Return False
        End If

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

        If txtNumeroCheque.Text.Length > 0 AndAlso txtNumeroCheque.Text > 0 Then
            MsgBox(Me.Page, "Título com cheque emitido não pode ser alterado.")
            Return False
        End If

        If DdlTributos.SelectedIndex > 0 Then 'Consulta Contas Contabeis dos Tributos
            Dim Encargo As New Encargo(DdlTributos.SelectedValue)
            If String.IsNullOrEmpty(Encargo.ContaDebito) And String.IsNullOrEmpty(Encargo.ContaCredito) Then
                MsgBox(Me.Page, "Encargo Sem contas Credito e Debito Cadastradas, Verifique o encargo.")
                Return False
            ElseIf Not String.IsNullOrWhiteSpace(Empresa.CodigoContaFornecedorFrete) AndAlso Not lblAgrupar.Text = "AP" _
                AndAlso String.IsNullOrWhiteSpace(txtRegistro.Text) AndAlso (Empresa.CodigoContaFornecedorFrete = Encargo.ContaDebito OrElse Empresa.CodigoContaFornecedorFrete = Encargo.ContaCredito) Then
                MsgBox(Me.Page, "Titulo não pode ser incluído na Conta de Fornecedor de Frete.")
                Return False
            End If
        End If

        '*****************************************************************************************************************
        '************************** Primeira Carteira Adiantamento/Baixa *************************************************
        '*****************************************************************************************************************
        If objCarteira.isAdiantamento AndAlso objCarteira.BaixaAdiantamento AndAlso (Not IsNumeric(txtNumeroAdto.Text) OrElse CInt(txtNumeroAdto.Text) <= 0) Then
            MsgBox(Me.Page, "Selecione um adiantamento para efetuar a baixa")
            Return False
        End If

        If Not objCarteira.isAdiantamento AndAlso DdlCarteirasAdto.SelectedIndex <= 0 AndAlso IsNumeric(txtNumeroAdto.Text) AndAlso CInt(txtNumeroAdto.Text) > 0 Then
            MsgBox(Me.Page, "Numero de adiantamento nao pode ser informado com carteira que nao sao de adiantamento ou baixa. reinicie o lançamento.")
            Return False
        End If

        If objCarteira.isAdiantamento AndAlso Not objCarteira.BaixaAdiantamento AndAlso String.IsNullOrWhiteSpace(txtVencimentoAdto.Text) Then
            MsgBox(Me.Page, "Vencimento para o Adiantamento não foi informado, Verifique.")
            Return False
        End If

        If objCarteira.isAdiantamento And objCarteira.BaixaAdiantamento Then
            If ddlMoeda.SelectedValue = 1 AndAlso CDec(HDSaldoAdiantamento.Value) < CDec(txtValorDoDocumento.Text) Then
                MsgBox(Me.Page, "Valor da Baixa não pode ser maior do que o saldo do adiantamento, em " & ddlMoeda.SelectedItem.Text & ".")
                Return False
            ElseIf CDec(HDSaldoAdiantamento.Value) < CDec(txtValorEmMoeda.Text) Then
                MsgBox(Me.Page, "Valor da Baixa não pode ser maior do que o saldo do adiantamento, em " & ddlMoeda.SelectedItem.Text & ".")
                Return False
            End If
        End If

        '*****************************************************************************************************************
        '************************** Segunda Carteira Adiantamento/Baixa *************************************************
        '*****************************************************************************************************************
        If DdlCarteirasAdto.SelectedIndex > 0 Then
            Dim objCarteiraadto As New [Lib].Negocio.CarteiraFinanceira(DdlCarteirasAdto.SelectedValue)

            If objCarteiraadto.isAdiantamento AndAlso objCarteiraadto.BaixaAdiantamento AndAlso (Not IsNumeric(txtNumeroAdto.Text) OrElse CInt(txtNumeroAdto.Text) <= 0) Then
                MsgBox(Me.Page, "Selecione um adiantamento para efetuar a baixa")
                Return False
            End If

            If objCarteiraadto.isAdiantamento And objCarteiraadto.BaixaAdiantamento Then
                If ddlMoeda.SelectedValue = 1 AndAlso CDec(HDSaldoAdiantamento.Value) < CDec(txtValorCobrado.Text) Then
                    MsgBox(Me.Page, "Valor da Baixa não pode ser maior do que o saldo do adiantamento, em " & ddlMoeda.SelectedItem.Text & ".")
                    Return False
                ElseIf Not ddlMoeda.SelectedValue = 1 AndAlso CDec(HDSaldoAdiantamento.Value) < CDec(txtValorCobradoMoeda.Text) Then
                    MsgBox(Me.Page, "Valor da Baixa não pode ser maior do que o saldo do adiantamento, em " & ddlMoeda.SelectedItem.Text & ".")
                    Return False
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

        If DdlProvisoes.SelectedValue = 1 And IIf(Trim(DdlTiposDePagamentos.Text <> ""), DdlTiposDePagamentos.SelectedValue, "0") = 4 Then 'baixa e boleto
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
            If Not String.IsNullOrWhiteSpace(DdlCarteirasAdto.SelectedValue) AndAlso DdlEmpresaCliente.SelectedValue <> DdlEmpresaPagadora.SelectedValue Then
                MsgBox(Me.Page, "Empresa pagadora não pode ser diferente da Empresa do titulo, com carteira de adiantamento selecionada.")
                Return False
            ElseIf CDec(txtValorDoDocumento.Text) = 0 AndAlso DdlEmpresaCliente.SelectedValue <> DdlEmpresaPagadora.SelectedValue Then
                MsgBox(Me.Page, "Empresa pagadora não pode ser diferente da Empresa do titulo, com valor do documento igual a zero.")
                Return False
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

        'Quando for colocada a tela nova onde a conta será manipulada pelo user control essa validação poderá ser removida - Cleberson
        'Verificar existência do banco e conta
        If DdlBancos.SelectedIndex > 0 AndAlso Not String.IsNullOrWhiteSpace(txtAgencia.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDigitoAgencia.Text) _
            AndAlso Not String.IsNullOrWhiteSpace(txtContaCorrente.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDigitoDaConta.Text) Then

            Dim Cli = txtCodigoFavorecido.Value.Split("-")
            Dim CliFav As Cliente = New Cliente(Cli(0), Cli(1))

            Dim ObjListaCxc As ListClienteXContaBancaria = New ListClienteXContaBancaria(CliFav)
            If ObjListaCxc.Where(Function(s) s.CodigoBanco = DdlBancos.SelectedValue _
                               AndAlso s.CodigoAgencia.ToString.Trim = txtAgencia.Text.Trim _
                                  AndAlso s.DigitoAgencia.ToString.Trim = txtDigitoAgencia.Text.Trim _
                                  AndAlso s.ContaCorrente.ToString.Trim = txtContaCorrente.Text.Trim _
                                  AndAlso s.DigitoConta.ToString.Trim = txtDigitoDaConta.Text.Trim).Count <= 0 Then
                MsgBox(Me.Page, "A conta bancária do cliente ainda não foi gravada. Grave a conta, selecione-a e então grave o título.")
                TabContainer1.ActiveTabIndex = 1
                Return False
            End If
        End If
        '------------------------------------------------------


        'Verificar se as informações como Unidade, Empresa e Cliente não tenham sido alteradas no "meio tempo" entre a consulta e baixa do título
        If Not String.IsNullOrWhiteSpace(txtPedido.Text) AndAlso CInt(txtPedido.Text) > 0 Then
            Dim sql As String
            sql = "SELECT Empresa_Id, EndEmpresa_Id, UnidadeDeNegocio, Cliente, EndCliente " & vbCrLf & _
                  "  FROM Pedidos " & vbCrLf & _
                  " WHERE pedido_id = " & txtPedido.Text & vbCrLf & _
                  "   AND UnidadeDeNegocio = '" & DdlUnidadeDeNegocioEmpresaCliente.SelectedValue & "'" & vbCrLf & _
                  "   AND Empresa_Id= '" & DdlEmpresaCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                  "   AND EndEmpresa_Id = " & DdlEmpresaCliente.SelectedValue.Split("-")(1) & vbCrLf & _
                  "   AND Cliente =  '" & txtCodigoFornecedor.Value.Split("-")(0) & "'" & vbCrLf & _
                  "   AND EndCliente = " & txtCodigoFornecedor.Value.Split("-")(1) & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Pedido")

            If ds.Tables("Pedido").Rows.Count <= 0 Then
                MsgBox(Me.Page, "Recarregue o título, pois houveram alterações em sua Nota ou Pedido.")
                Return False
            End If
        End If

        Return True
    End Function

    Private Function LanctosContabeis()
        If Len(Raz_Conta) = 0 Or (Raz_ValorMoeda = 0 And Raz_ValorOficial = 0) Then Return False

        Sql = "INSERT INTO Razao " & vbCrLf & _
              "       (Empresa_Id, " & vbCrLf & _
              "       EndEmpresa_Id, " & vbCrLf & _
              "       Conta_Id, " & vbCrLf & _
              "       Cliente_Id, " & vbCrLf & _
              "       EndCliente_Id, " & vbCrLf & _
              "       Movimento_Id, " & vbCrLf & _
              "       Lote_Id, " & vbCrLf & _
              "       Sequencia_Id, " & vbCrLf & _
              "       Titulo, " & vbCrLf & _
              "       UnidadeDeNegocio, " & vbCrLf & _
              "       Indexador, " & vbCrLf & _
              "       DataMoeda, " & vbCrLf & _
              "       DebitoOficial, " & vbCrLf & _
              "       CreditoOficial, " & vbCrLf & _
              "       DebitoMoeda, " & vbCrLf & _
              "       CreditoMoeda, " & vbCrLf & _
              "       Conciliacao, " & vbCrLf & _
              "       DataDaBaixa, " & vbCrLf & _
              "       Historico, " & vbCrLf & _
              "       PrevistoRealizado," & vbCrLf & _
              "       Processo," & vbCrLf & _
              "       UsuarioInclusao," & vbCrLf & _
              "       UsuarioInclusaoData)" & vbCrLf & _
              "VALUES ('" & Raz_Empresa & "'," & vbCrLf & _
              "         " & Raz_EndEmpresa & "," & vbCrLf & _
              "        '" & Raz_Conta & "'" & vbCrLf

        If Len(Raz_Conta) = 7 Then
            Sql &= ", '" & Raz_Cliente & "'"        'Cliente
            Sql &= ", " & Raz_EndCliente            'Endereco do Cliente
        Else
            Sql &= ", ''"                           'Cliente
            Sql &= ", 0"                            'Endereco do Cliente
        End If

        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'"     'Data de Movimento
        Sql &= ", 0070"
        Sql &= ", " & Registro                      'Sequencia no Razao = Registro do Titulo
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

        If Raz_DebitoCredito = "C" AndAlso chkConciliado.Checked Then
            Sql &= ", 'B'"                                                             'Conciliação
            Sql &= ", '" & CDate(txtDataConciliacao.Value).ToString("yyyy/MM/dd") & "'" 'Data Conciliação
        Else
            Sql &= ", NULL "                                                           'Conciliação
            Sql &= ", NULL "                                                           'Data Conciliação
        End If

        Sql &= ", '" & Raz_Historico & "'"          'Histórico
        Sql &= ", 'P'"                              'Previsto/Realizado
        Sql &= ", 'CONTASAPAGAR'"                   'Processo
        Sql &= ", '" & Session("ssNomeUsuario") & "'"  'Usuario que Baixou
        Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')"           'Data da Baixa

        SqlArray.Add(Sql)

        Return True

    End Function

    Private Function IsVariacao(ByVal Empresa As String, ByVal Carteira As String, ByVal campo As String) As Boolean
        Dim sql As String = "   SELECT CASE                                                                                                        " & vbCrLf & _
                            "            WHEN cart.conta" & campo & "     in (empDed.contaVariacaoMonetariaAtiva, empDed.contaVariacaoMonetariaPassiva) " & vbCrLf & _
                            "              THEN 1                                                                                                  " & vbCrLf & _
                            "              ELSE 0                                                                                                  " & vbCrLf & _
                            "          END Variacao                                                                                               " & vbCrLf & _
                            "     FROM ComprasXProdutos cart,                                                                                      " & vbCrLf & _
                            "         clientesxempresas empDed                                                                                     " & vbCrLf & _
                            "    Where PRODUTO_ID           = '" & Carteira & "'                                                                   " & vbCrLf & _
                            "      AND empDed.empresa_id    = '" & Empresa.Split("-")(0) & "'" & vbCrLf & _
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
        Dim sql As String = "SELECT CASE" & vbCrLf & _
                            "        WHEN '" & Conta & "' in (empDed.contaVariacaoMonetariaAtiva, empDed.contaVariacaoMonetariaPassiva) " & vbCrLf & _
                            "          THEN 1" & vbCrLf & _
                            "          ELSE 0" & vbCrLf & _
                            "       END Variacao" & vbCrLf & _
                            "  FROM clientesxempresas empDed" & vbCrLf & _
                            " Where empDed.empresa_id    = '" & Empresa.Split("-")(0) & "'" & vbCrLf & _
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
        Dim Empresa() As String = DdlEmpresaCliente.SelectedValue.Split("-")
        Dim Cliente() As String = txtCodigoFornecedor.Value.Split("-")
        Dim EmpresaPagadora() As String = DdlEmpresaPagadora.SelectedValue.Split("-")

        Dim objPedido As New [Lib].Negocio.Pedido()
        If txtPedido.Text.Length > 0 AndAlso CInt(txtPedido.Text) > 0 Then
            objPedido = New [Lib].Negocio.Pedido(Empresa(0), Empresa(1), txtPedido.Text)
        End If

        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)
        Dim objCarteiraAdto As New [Lib].Negocio.CarteiraFinanceira(DdlCarteirasAdto.SelectedValue)

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
            Next
        Else
            Registro = CInt(txtRegistro.Text)

            Sql = "DELETE TitulosXDesdobrarFornecedor " & vbCrLf & _
                  " WHERE Registro_Id = " & Registro & vbCrLf
            SqlArray.Add(Sql)

            Sql = "DELETE AdiantamentosXBaixas" & vbCrLf & _
                  " WHERE Titulo = " & Registro
            SqlArray.Add(Sql)

            Sql = "DELETE Adiantamentos" & vbCrLf & _
                  " WHERE Titulo = " & Registro
            SqlArray.Add(Sql)

            Sql = "DELETE FROM razao" & vbCrLf & _
                  " WHERE Titulo = " & Registro
            SqlArray.Add(Sql)

            Sql = "DELETE FROM ContasAPagar" & vbCrLf & _
                  " WHERE Registro_Id = " & Registro
            SqlArray.Add(Sql)
        End If


        Sql = "INSERT INTO ContasAPagar " & vbCrLf & _
              "       (Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda" & vbCrLf & _
              "       ,TipoPagto, Situacao, Lote" & vbCrLf & _
              "       ,Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa" & vbCrLf & _
              "       ,UnidadeDeNegocio, Empresa, EndEmpresa" & vbCrLf & _
              "       ,Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, TipoContaCliente, ContaContabilCliente" & vbCrLf & _
              "       ,EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, ContaContabilPagadora" & vbCrLf & _
              "       ,Cheque, Slips, Recibo, Aviso, ReciboDeposito" & vbCrLf & _
              "       ,EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao" & vbCrLf & _
              "       ,ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido" & vbCrLf & _
              "       ,MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos, MoedaValorLiquido" & vbCrLf & _
              "       ,Historico" & vbCrLf & _
              "       ,CodigoDeBarras, CodigoDigitado, CodigoDeBarraPreImpresso" & vbCrLf & _
              "       ,Destinatario, EndDestinatario, solicitacao" & vbCrLf & _
              "       ,UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioBaixa, UsuarioBaixaData, UsuarioLiberacao, UsuarioLiberacaoData" & vbCrLf & _
              "       ,Grupado" & vbCrLf & _
              "       ,Observacoes" & vbCrLf & _
              "       ,SituacaoBancaria, NumeroDoCheque" & vbCrLf & _
              "       ,VencimentoAdto, TaxaAdto" & vbCrLf & _
              "       ,UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, UsuarioLiberacaoCheque, UsuarioLiberacaoChequeDate" & vbCrLf & _
              "       ,CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento, ContratoBancario)" & vbCrLf & _
              "VALUES ( " & Registro & ", 0, " & DdlProvisoes.SelectedValue & ", '" & ddlCarteiras.SelectedValue & "'" & IIf(DdlTributos.Text <> "", ", '" & DdlTributos.SelectedValue & "'", ", ''") & ", " & ddlIndexador.SelectedValue & ", " & ddlMoeda.SelectedValue & vbCrLf & _
              IIf(DdlTiposDePagamentos.Text <> "", ", " & DdlTiposDePagamentos.SelectedValue, ", 0") & ", 1, 70," & vbCrLf & _
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

        '-Dados Bancarios Cliente----------
        If DdlBancos.Text <> "" Then
            Sql &= ", " & DdlBancos.SelectedValue               'Banco Cliente
        Else
            Sql &= ", 0"                                        'Banco Cliente
        End If

        'Agencia do Destinatario 'Digito da Agencia do Destinatário 'Conta Corrente do Destinatário 'Digito da Conta Corrente do Destinatário 'Tipo da Conta do Destinatário
        Sql &= ", '" & txtAgencia.Text.Trim & "', '" & txtDigitoAgencia.Text.Trim & "', '" & txtContaCorrente.Text.Trim & "', '" & txtDigitoDaConta.Text.Trim & "', '" & ddlTipoConta.SelectedValue & "'"


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
            Sql &= ", '" & conta(4) & "'"                           'Conta Contabil
        Else
            Sql &= ", '', '', '', '', ''"
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

        If lblUsuarioIncl.Text.Length = 0 Then
            Sql &= ", '" & Session("ssNomeUsuario") & "'"             'Usuario Que Esta Incluindo
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "'"  'Data da desta Inclusao
            Sql &= ", ''"                                             'UsuarioAlteracao
            Sql &= ", ''"                                             'UsuarioData
        Else
            Dim Usu As Array = lblUsuarioIncl.Text.Trim.Split("-")
            Sql &= ", '" & Usu(0) & "'"                                'Usuario que Incluiu
            Sql &= ", '" & CDate(Usu(1)).ToString("yyyy/MM/dd") & "'"  'Data De Quando Ocorreu a Inclusao
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

        If DdlCarteirasAdto.SelectedIndex > 0 Then                                   'Carteira de Adiantamento 
            Sql &= ", '" & DdlCarteirasAdto.SelectedValue & "'"
        Else
            Sql &= ", ''"
        End If

        Sql &= ", " & ddlCarteiraDoTitulo.SelectedValue & ""                         'Carteira do Titulo 
        Sql &= ", '" & Trim(txtContratoFinanceiro.Text) & "'"                        'Contrato De Financiamento 
        Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(txtContratoBanco.Text) & "')" 'Contrato Bancário

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
                If Not Adiantamento(objPedido, txtRegistro.Text, VlrBrutoOficial, VlrBrutoMoeda) Then
                    MsgBox(Me.Page, Mensagem)
                    Exit Sub
                End If
            End If

            If objCarteira.BaixaAdiantamento Or objCarteiraAdto.BaixaAdiantamento Then
                If Trim(txtNumeroAdto.Text) <> "" And Trim(txtNumeroAdto.Text) <> "0" Then
                    If Not AdiantamentoAmortizacao(objPedido, txtRegistro.Text, txtNumeroAdto.Text, IIf(objCarteira.BaixaAdiantamento, VlrBrutoOficial, VlrLiquidoOficial), IIf(objCarteira.BaixaAdiantamento, VlrBrutoMoeda, VlrLiquidoMoeda)) Then
                        MsgBox(Me.Page, Mensagem)
                        Exit Sub
                    End If
                End If
            End If
        End If

        'AGRUPAMENTO DE TITULOS
        If lblAgrupar.Text = "AP" Then
            For index = 0 To Session("ssRegistros" & HID.Value).Count - 1
                Dim tit As New Titulo(Session("ssRegistros" & HID.Value).Item(index), "P")

                Sql = "SELECT CP.Registro_Id, CP.Sequencia_Id, CP.Provisao, CP.Carteira, CP.Tributo, CP.Indexador, " & vbCrLf & _
                      "       CP.Moeda, CP.TipoPagto, CP.Situacao, CP.Lote, CP.Movimento, CP.Vencimento, CP.Prorrogacao, " & vbCrLf & _
                      "       CP.DataMoeda, isnull(CP.Baixa,CP.Prorrogacao) as Baixa, CP.UnidadeDeNegocio, CP.Empresa, CP.EndEmpresa, CP.Cliente, " & vbCrLf & _
                      "       CP.EndCliente, CP.BancoCliente, CP.AgenciaCliente, CP.DigitoAgenciaCliente, CP.ContaCliente, " & vbCrLf & _
                      "       CP.DigitoContaCliente, Isnull(CP.TipoContaCliente,'C') AS TipoContaCliente, CP.ContaContabilCliente, " & vbCrLf & _
                      "       CP.EmpresaPagadora, CP.EndEmpresaPagadora, CP.BancoPagador, CP.AgenciaPagadora, CP.DigitoAgenciaPagadora, " & vbCrLf & _
                      "       CP.ContaPagadora, CP.DigitoContaPagadora, CP.ContaContabilPagadora, CP.Cheque, isnull(CP.Slips,'N') AS Slips, " & vbCrLf & _
                      "       CP.Recibo, CP.Aviso, CP.ReciboDeposito, isnull(CP.EmpresaPedido,'') AS EmpresaPedido, isnull(CP.EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(CP.Pedido, 0) AS Pedido, " & vbCrLf & _
                      "       isnull(CP.PedidoFixacao,0) AS PedidoFixacao, isnull(CP.Procuracao,0) AS Procuracao, CP.ValorDoDocumento, CP.Descontos, CP.Deducoes, " & vbCrLf & _
                      "       CP.Juros, CP.Acrescimos, CP.ValorLiquido, ISNULL(CP.MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, " & vbCrLf & _
                      "       ISNULL(CP.MoedaDescontos, 0) AS MoedaDescontos, ISNULL(CP.MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(CP.MoedaJuros, 0) AS MoedaJuros, " & vbCrLf & _
                      "       ISNULL(CP.MoedaAcrescimos, 0) AS MoedaAcrescimos, ISNULL(CP.MoedaValorLiquido, 0) AS MoedaValorLiquido, CP.Historico, " & vbCrLf & _
                      "       CP.CodigoDeBarras, CP.CodigoDigitado, CP.Destinatario, CP.EndDestinatario, CP.NomeDoDestinatario, CP.Destinacao, " & vbCrLf & _
                      "       CP.Solicitacao, CP.UsuarioInclusao, CP.UsuarioInclusaoData, CP.UsuarioAlteracao, CP.UsuarioAlteracaoData, " & vbCrLf & _
                      "       CP.UsuarioCancelamento, CP.UsuarioCancelamentoData, isnull(CP.UsuarioLiberacao,'') AS UsuarioLiberacao, CP.UsuarioLiberacaoData, " & vbCrLf & _
                      "       CP.UsuarioBaixa, CP.UsuarioBaixaData, isnull(CP.Grupado,'N') AS Grupado, isnull(CP.RegistroMestre, 0) as RegistroMestre, CP.Observacoes, " & vbCrLf & _
                      "       CP.SituacaoBancaria, ISNULL(CP.NumeroDoCheque,0) AS NumeroDoCheque, isnull(CP.Adiantamento,0) AS Adiantamento, CP.VencimentoAdto, CP.TaxaAdto, " & vbCrLf & _
                      "       isnull(CP.UsuarioLiberacaoBloqueio, '') AS UsuarioLiberacaoBloqueio, CP.UsuarioLiberacaoBloqueioDate, isnull(CP.UsuarioLiberacaoPedido, '') AS UsuarioLiberacaoPedido, " & vbCrLf & _
                      "       CP.UsuarioLiberacaoPedidoDate, isnull(CP.UsuarioLiberacaoCheque, '') AS UsuarioLiberacaoCheque, CP.UsuarioLiberacaoChequeDate, " & vbCrLf & _
                      "       isnull(CP.CarteiraAdto,'') AS CarteiraAdto, isnull(CP.CarteiraDoTitulo,0) AS CarteiraDoTitulo, isnull(CP.ContratoDeFinanciamento,'') AS ContratoDeFinanciamento, " & vbCrLf & _
                      "       ISNULL(NF.TipoDeDocumento,0) AS TipoDeDocumento, ISNULL(FFxT.Fatura_Id, 0) AS FaturaDeFrete" & vbCrLf & _
                      "  FROM ContasAPagar CP " & vbCrLf & _
                      "  LEFT JOIN NotaFiscalXTitulo NFxT " & vbCrLf & _
                      "	   ON CP.Registro_Id = NFxT.Titulo_Id " & vbCrLf & _
                      "	 LEFT JOIN NotasFiscais NF " & vbCrLf & _
                      "	   ON NFxT.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                      "   AND NFxT.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                      "   AND NFxT.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                      "   AND NFxT.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                      "   AND NFxT.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                      "   AND NFxT.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                      "   AND NFxT.Nota_Id         = NF.Nota_Id" & vbCrLf & _
                      "  LEFT JOIN FaturaDeFreteXTitulo FFxT " & vbCrLf & _
                      "    ON CP.Registro_Id = FFxT.Titulo_Id " & vbCrLf & _
                      " WHERE CP.Registro_Id = " & CStr(Session("ssRegistros" & HID.Value).Item(index)) & vbCrLf & _
                      "   and CP.Situacao not in (2,3,4,5,6,10) " & vbCrLf

                Dim dsFilho As New DataSet
                dsFilho = Banco.ConsultaDataSet(Sql, "ContasAPagarXFilho")

                If Not dsFilho Is Nothing AndAlso dsFilho.Tables(0).Rows.Count > 0 Then
                    For Each drFilho As DataRow In dsFilho.Tables(0).Rows
                        Sql = " DELETE FROM razao" & vbCrLf & _
                              "  WHERE Titulo = " & CStr(Session("ssRegistros" & HID.Value).Item(index)) & vbCrLf
                        SqlArray.Add(Sql)

                        Sql = " DELETE FROM Adiantamentos" & vbCrLf & _
                              "  WHERE Titulo = " & CStr(Session("ssRegistros" & HID.Value).Item(index)) & vbCrLf
                        SqlArray.Add(Sql)

                        Sql = " DELETE FROM AdiantamentosXBaixas" & vbCrLf & _
                              "  WHERE Titulo = " & CStr(Session("ssRegistros" & HID.Value).Item(index)) & vbCrLf

                        SqlArray.Add(Sql)

                        Sql = " UPDATE ContasAPagar" & vbCrLf & _
                              "    SET Provisao = " & DdlProvisoes.SelectedValue & vbCrLf & _
                              "        ,TipoPagto   = '" & DdlTiposDePagamentos.SelectedValue & "'" & vbCrLf & _
                              "        ,DataMoeda   = '" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                              "        ,Baixa       = '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                              "        ,Prorrogacao = '" & CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf

                        If DdlTributos.SelectedIndex > 0 Then
                            Sql &= ", Tributo = '" & DdlTributos.SelectedValue & "'" & vbCrLf
                        End If

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
                            '(0)Agencia Cliente (1)Digito Agencia Cliente (2)Conta Cliente (3)Digito Conta Cliente (4)Conta Contabil
                            Sql &= ", AgenciaPagadora       = '" & Conta(0) & "', DigitoAgenciaPagadora = '" & Conta(1) & "', ContaPagadora         = '" & Conta(2) & "', DigitoContaPagadora   = '" & Conta(3) & "', ContaContabilPagadora = '" & Conta(4) & "'"
                        End If

                        If DdlBancos.Text <> "" Then
                            Sql &= ", BancoCliente         = " & DdlBancos.SelectedValue               'Banco Cliente
                            Sql &= ", AgenciaCliente       ='" & txtAgencia.Text.Trim & "'"            'Agencia do Destinatario
                            Sql &= ", DigitoAgenciaCliente ='" & txtDigitoAgencia.Text.Trim & "'"      'Digito da Agencia do Destinatário
                            Sql &= ", ContaCliente         ='" & txtContaCorrente.Text.Trim & "'"      'Conta Corrente do Destinatário
                            Sql &= ", DigitoContaCliente   ='" & txtDigitoDaConta.Text.Trim & "'"      'Digito da Conta Corrente do Destinatário
                            Sql &= ", TipoContaCliente     ='" & ddlTipoConta.SelectedValue & "'"      'Tipo da Conta do Destinatário
                        End If

                        Sql &= ", UsuarioAlteracao = '" & Session("ssNomeUsuario") & "'"                 'Usuario que Incluiu
                        Sql &= ", UsuarioAlteracaoData = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"  'Data da Inclusao
                        Sql &= ", ContratoBancario = '" & Funcoes.EliminarCaracteresEspeciais(txtContratoBanco.Text) & "'" 'Contrato Bancário

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
                                    If Not Adiantamento(objPedidoAP, drFilho("Registro_Id"), VlrBrutoOficial, VlrBrutoMoeda) Then
                                        MsgBox(Me.Page, Mensagem)
                                        Exit Sub
                                    End If
                                End If

                                If objCarteiraAP.BaixaAdiantamento Or objCarteiraAdtoAP.BaixaAdiantamento Then
                                    If Trim(drFilho("Adiantamento")) <> "" And Trim(drFilho("Adiantamento")) <> "0" Then
                                        If Not AdiantamentoAmortizacao(objPedidoAP, drFilho("Registro_Id"), drFilho("Adiantamento"), IIf(objCarteiraAP.BaixaAdiantamento, VlrBrutoOficial, VlrLiquidoOficial), IIf(objCarteiraAP.BaixaAdiantamento, VlrBrutoMoeda, VlrLiquidoMoeda)) Then
                                            MsgBox(Me.Page, Mensagem)
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If


                            Dim Carteira As New CarteiraFinanceira(drFilho("Carteira"))

                            ' Grava Razao Debito
                            '------------------------------------------
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
                            Raz_DebitoCredito = "D"                         'Debito/Credito

                            LanctosContabeis()

                            '-------Descontos-----------

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
                                Raz_DebitoCredito = "C"                         'Debito/Credito

                                LanctosContabeis()
                            End If

                            '----Deducoes--------------

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
                                Raz_DebitoCredito = "C"                         'Debito/Credito
                                LanctosContabeis()
                            End If

                            '----Juros--------------

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
                                Raz_DebitoCredito = "D"                         'Debito/Credito
                                LanctosContabeis()
                            End If

                            '----Acrescimos--------------

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
                                Raz_DebitoCredito = "D"                         'Debito/Credito

                                LanctosContabeis()
                            End If
                        End If
                    Next
                End If
            Next
        End If


        '-Gravação no Razão Contabil------------------

        'ConsultaCarteiras(ddlCarteiras.SelectedValue)               'Consulta Contas Contabeis
        Dim Carteira2 As New CarteiraFinanceira(ddlCarteiras.SelectedValue)

        If DdlProvisoes.SelectedValue = 1 And lblAgrupar.Text = "" Then
            ' Grava Razao Debito
            '------------------------------------------
            Registro = CInt(txtRegistro.Text)
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

            Raz_DebitoCredito = "D"

            LanctosContabeis()

            '-------Descontos-----------

            If CDec(txtDescontos.Text) > 0 Then
                Registro = CInt(txtRegistro.Text)
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
                Raz_DebitoCredito = "C"

                LanctosContabeis()

            End If

            '----Deducoes--------------

            If CDec(txtDeducoes.Text) > 0 Then
                Registro = CInt(txtRegistro.Text)
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
                Raz_DebitoCredito = "C"

                LanctosContabeis()
            End If

            '----Juros--------------

            If CDec(txtJuros.Text) > 0 Then
                Registro = CInt(txtRegistro.Text)
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
                Raz_DebitoCredito = "D"

                LanctosContabeis()
            End If

            '----Acrescimos--------------

            If CDec(txtAcrescimos.Text) > 0 Then
                Registro = CInt(txtRegistro.Text)
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
                Raz_DebitoCredito = "D"

                LanctosContabeis()
            End If

        End If

        '-----------------------
        'Gravar Credito
        '-----------------------
        If DdlProvisoes.SelectedValue = 1 And (lblAgrupar.Text = "" Or lblAgrupar.Text = "AP") Then
            If DdlCarteirasAdto.Text = "" Then
                Registro = CInt(txtRegistro.Text)
                Raz_Empresa = EmpresaPagadora(0)                  'Empresa Pagadora
                Raz_EndEmpresa = EmpresaPagadora(1)               'Endereco Empresa Pagadora

                Dim conta() As String = DdlContaPagadora.SelectedValue.Split("-")
                Raz_Conta = conta(4)                    'Conta Contabil


                Dim objPlaConta As New [Lib].Negocio.PlanoDeConta("", 0, Raz_Conta)
                If objPlaConta.TemCliente Then
                    Raz_Cliente = Cliente(0)                              'Cliente
                    Raz_EndCliente = Cliente(1)
                Else
                    Raz_Cliente = ""                                'Cliente
                    Raz_EndCliente = 0                              'Endereco do Cliente
                End If

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'estava fixo "01"  antes 19/06/12   'Pegar Unidadde de Negocio da EmpresaPagadora

                Raz_ValorOficial = Replace(txtValorCobrado.Text, ".", "") 'Valor Liquido

                If CDec(txtValorCobradoMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtValorCobradoMoeda.Text, ".", "")
                Else
                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, txtValorCobrado.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                End If

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)

                Raz_DebitoCredito = "C"
                LanctosContabeis()
            Else
                Registro = CInt(txtRegistro.Text)
                Raz_Empresa = EmpresaPagadora(0)                                          'Empresa Pagadora
                Raz_EndEmpresa = EmpresaPagadora(1)                                       'Endereco Empresa Pagadora
                Raz_Conta = objCarteiraAdto.CodigoContaCliente                        'Conta Contabil

                campo = txtCodigoFavorecido.Value.Split("-")
                Raz_Cliente = campo(0)                                          'Cliente
                Raz_EndCliente = campo(1)                                       'Endereco do Cliente

                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue 'estava fixo "01"  antes 19/06/12  'Pegar Unidadde de Negocio da EmpresaPagadora

                Raz_ValorOficial = Replace(txtValorCobrado.Text, ".", "") 'Valor Liquido

                If CDec(txtValorCobradoMoeda.Text) > 0 OrElse ddlIndexador.SelectedValue = 99 Then
                    Raz_ValorMoeda = Replace(txtValorCobradoMoeda.Text, ".", "")
                Else
                    Raz_ValorMoeda = Replace(DolarizaBaixa(DtMovimento, txtValorCobrado.Text, IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue)), ".", "")
                End If

                Raz_Historico = Funcoes.EliminarCaracteresEspeciais(txtHistorico.Text.Trim & ". " & txtObservacoes.Text.Trim)

                Raz_DebitoCredito = "C"
                LanctosContabeis()
            End If

            '-------------------------------------------
            'Transferencias Financeiras
            '-------------------------------------------
            Sql = " SELECT EmpresaDebito, EnderecoDebito, EmpresaCredito, EnderecoCredito, EmpresaContabil, EnderecoContabil, " & vbCrLf & _
                  "        ContaContabil, ClienteContabil,EndClienteContabil,DebitoCredito " & vbCrLf & _
                  "   FROM TransferenciasFinanceiras " & vbCrLf & _
                  "  WHERE EmpresaDebito   ='" & Empresa(0) & "'" & vbCrLf & _
                  "    and EnderecoDebito  = " & Empresa(1) & vbCrLf & _
                  "    and EmpresaCredito  ='" & EmpresaPagadora(0) & "'" & vbCrLf & _
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
                LanctosContabeis()
            Next
        End If

        If Not Session("objDestinoContabil" & HID.Value) Is Nothing AndAlso DdlProvisoes.SelectedValue = 1 Then
            Dim objDestinoContabil() As String = Session("objDestinoContabil" & HID.Value).ToString.Split("-")
            Sql = " INSERT INTO TitulosXDesdobrarFornecedor (Registro_Id, Cliente, EndCliente, Pedido, Carteira) " & vbCrLf & _
                  " VALUES (" & Registro & ",'" & objDestinoContabil(0) & "'," & objDestinoContabil(1) & "," & 0 & ",'" & objDestinoContabil(2) & "')" & vbCrLf
            SqlArray.Add(Sql)
        End If

        If Banco.GravaBanco(SqlArray) = False Then
            MsgBox(Me.Page, Session("ssMessage"))
        Else
            If Not GravarPamcard(Registro) Then
                SqlArray.Clear()

                Sql = " DELETE Razao " & vbCrLf & _
                      "  WHERE Titulo = " & Registro
                SqlArray.Add(Sql)

                Sql = "UPDATE ContasAPagar SET " & vbCrLf & _
                      "   Provisao = 2 " & vbCrLf & _
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

    Function GravarPamcard(ByVal titulo As String) As Boolean
        If txtTipoDoDocumento.Value = 4 AndAlso DdlProvisoes.SelectedValue = 1 Then
            Dim sql As String = ""
            sql = "SELECT NFxT.Empresa_Id, NFxT.EndEmpresa_Id, NFxT.Cliente_Id, " & vbCrLf & _
                  "       NFxT.EndCliente_Id, NFxT.EntradaSaida_Id, NFxT.Serie_Id, " & vbCrLf & _
                  "       NFxT.Nota_Id " & vbCrLf & _
                  "  FROM  ContasAPagar CP " & vbCrLf & _
                  " INNER JOIN NotaFiscalXTitulo NFxT " & vbCrLf & _
                  "    ON CP.Registro_Id = NFxT.Titulo_Id " & vbCrLf & _
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

        Sql = "INSERT INTO ContasAPagar " & vbCrLf & _
              " (Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa," & vbCrLf & _
              "  UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente," & vbCrLf & _
              "  ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora," & vbCrLf & _
              "  ContaContabilPagadora, cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao," & vbCrLf & _
              "  ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros," & vbCrLf & _
              "  MoedaAcrescimos, MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, solicitacao," & vbCrLf & _
              "  UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao," & vbCrLf & _
              "  UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, RegistroMestre, Observacoes, SituacaoBancaria," & vbCrLf & _
              "  UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate," & vbCrLf & _
              "  UsuarioLiberacaoCheque, UsuarioLiberacaoChequeDate, CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento)" & vbCrLf & _
              "SELECT " & NovaSeq & ", Sequencia_Id, 2, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa," & vbCrLf & _
              "       UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente," & vbCrLf & _
              "       ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora," & vbCrLf & _
              "       ContaContabilPagadora, cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao," & vbCrLf & _
              "       " & Str(VlrOficial) & ", 0, 0, 0, 0, " & Str(VlrOficial) & "," & vbCrLf & _
              "       " & Str(VlrMoeda) & ", 0, 0, 0, 0, " & Str(VlrMoeda) & "," & vbCrLf & _
              "       Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, solicitacao," & vbCrLf & _
              "       '" & Session("ssNomeUsuario") & "', '" & CDate(Today).ToString("yyyy/MM/dd") & "', '', NULL, '', NULL, ''," & vbCrLf & _
              "       NULL, '', NULL, Grupado, RegistroMestre, Observacoes, SituacaoBancaria," & vbCrLf & _
              "       UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate," & vbCrLf & _
              "       '', NULL, CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento" & vbCrLf & _
              "  FROM CONTASAPAGAR" & vbCrLf & _
              " WHERE Registro_Id = " & TituloOrigem

        SqlArray.Add(Sql)

        Sql = " INSERT INTO NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id) " & vbCrLf & _
              " Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, " & NovaSeq & vbCrLf & _
              "   From NotaFiscalXTitulo" & vbCrLf & _
              "  Where Titulo_Id = " & TituloOrigem
        SqlArray.Add(Sql)

        Sql = " INSERT INTO FaturaDeFreteXTitulo(Empresa_Id, EndEmpresa_Id, Conveniado_Id, EndConveniado_Id, Fatura_Id, Titulo_Id) " & vbCrLf & _
              " Select Empresa_Id, EndEmpresa_Id, Conveniado_Id, EndConveniado_Id, Fatura_Id, " & NovaSeq & vbCrLf & _
              "   From FaturaDeFreteXTitulo" & vbCrLf & _
              "  Where Titulo_Id = " & TituloOrigem
        SqlArray.Add(Sql)

        MensagemParcial = " E Registro Parcial de Número <" & NovaSeq & ">..."
    End Sub

    Private Sub Limpar(ByVal LimparConsulta As Boolean)
        Try
            Session.Remove("objDestinoContabil" & HID.Value)
            Session.Remove("objPedidoSelecionadoCTAPAG" & HID.Value)
            Session.Remove("ssObservacoes" & HID.Value)
            Session.Remove("ssNomeLiberacao" & HID.Value)
            Session.Remove("ssNomeLiberacaoData" & HID.Value)
            Session.Remove("ssRegistros" & HID.Value)
            Session.Remove("ssGrupado" & HID.Value)
            Session.Remove("objFornecedorCP" & HID.Value)
            Session.Remove("objFavorecidoCP" & HID.Value)
            Session.Remove("objContasAPagar" & HID.Value)
            Session.Remove("ControleCP" & HID.Value)
            Session.Remove("ssRetornoDs" & HID.Value)
            Session.Remove("ssRetorno" & HID.Value)

            HID.Value = Guid.NewGuid().ToString
            ucConsultaClientes.SetarHID(HID.Value)
            ucConsultaAdiantamentos.SetarHID(HID.Value)
            ucConsultaPedidos.SetarHID(HID.Value)
            ucDestinoContabil.SetarHID(HID.Value)

            If LimparConsulta AndAlso lblAgrupar.Text = "AP" Then Limpar_ConsultaTitulos(False)

            If Not chkManterLancamento.Checked Then
                DdlCarteirasAdto.Parent.Visible = True
                txtVencimentoAdto.Parent.Visible = True
                txtNumeroAdto.Parent.Visible = True

                SqlArray.Clear()

                txtTipoDoDocumento.Value = 0

                lblUsuarioIncl.Text = ""
                lblUsuarioAlt.Text = ""
                imgUsuarioIncl.Visible = False
                imgUsuarioAlt.Visible = False
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

                DdlTributos.Items.Clear()

                DdlEmpresaPagadora.SelectedIndex = 0
                DdlBancoPagador.Items.Clear()
                DdlContaPagadora.Items.Clear()

                lblCotacao.Text = String.Empty
                lblDescCotacao.Text = String.Empty

                ddlCarteiras.Enabled = True
                ddlCarteiras.SelectedIndex = 0
                DdlTiposDePagamentos.SelectedIndex = 0
                ddlMoeda.SelectedValue = 1
                ddlIndexador.SelectedIndex = 3
                DdlProvisoes.SelectedIndex = 0

                ddlCarteiraDoTitulo.SelectedIndex = 0
                txtContratoFinanceiro.Text = ""
                ViewState.Clear()
                txtHistorico.Text = ""
                ddlSelecionarHist.SelectedValue = String.Empty
                txtHistorico.Enabled = True
                txtCodigoDeBarras.Text = ""

                txtPedido.Text = "0"
                cmdPedido.Enabled = True
                txtCessaoDeCredito.Text = "0"

                txtSolicitacao.Text = ""
                CkbCodigoDeBarras.Checked = False

                txtFavorecido.Text = ""
                txtCodigoFavorecido.Value = ""
                DdlBancos.SelectedIndex = 0
                txtAgencia.Text = ""
                txtDigitoAgencia.Text = ""
                txtContaCorrente.Text = ""
                txtDigitoDaConta.Text = ""
                ddlTipoConta.SelectedIndex = 0
                txtObservacoesDaConta.Text = ""
                txtVencimentoAdto.Text = ""
                txtTaxaAdto.Text = 0
                txtNumeroCheque.Text = ""

                DdlCarteirasAdto.SelectedIndex = 0

                GridContasCorrentes.DataSource = Nothing
                GridContasCorrentes.DataBind()

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

            lnkNovo.Parent.Visible = False
            lnkExcluir.Parent.Visible = False
            lnkRelatorio.Parent.Visible = False

            txtRegistro.Text = ""
            txtRegistro.Enabled = True
            txtProrrogacao.Text = ""
            hdnProrrogacaoOriginal.Value = String.Empty
            lblVencOriginal.Text = ""

            txtDataEntradaSistema.Text = CDate(Today).ToString("dd/MM/yyyy")

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

            txtValorDoDocumento.Text = ""
            HDValorOriginalOficial.Value = 0
            HDValorOriginalMoeda.Value = 0
            txtValorEmMoeda.Text = ""
            txtValorLiquido.Value = 0
            txtValorLiquidoMoeda.Value = 0

            txtNumeroAdto.Text = "0"
            HDSaldoAdiantamento.Value = 0

            txtPedidoConsultaTitulos.Text = ""

            ImgCalcular.Enabled = True

            TabPanel4.Visible = True

            txtValorDoDocumento.ForeColor = Drawing.Color.Black
            txtValorEmMoeda.ForeColor = Drawing.Color.Black
            txtOficial.ForeColor = Drawing.Color.Black
            txtMoeda.ForeColor = Drawing.Color.Black
            cmdPedido.Enabled = True
        Catch ex As Exception
            MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(ex.Message))
        End Try
    End Sub

    Protected Sub GridConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Registro = GridConsultaTitulos.SelectedRow.Cells(3).Text()
            Limpar(False)
            txtRegistro.Text = Registro
            ConsultaContasAPagar()
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

    Private Sub ConsultaContasAPagar()
        Dim Pedido As Integer = 0
        Dim SaldoParcelas As Decimal
        Dim conta As String = ""
        Dim ConferenciaNF As Boolean = False
        TemRegistro = ""

        Try
            If txtRegistro.Text <> "" Then
                Registro = txtRegistro.Text
                txtRegistro.Text = Registro

                Sql = " SELECT CP.Registro_Id, CP.Sequencia_Id, CP.Provisao, CP.Carteira, CP.Tributo, CP.Indexador, " & vbCrLf & _
                      "        CP.Moeda, CP.TipoPagto, CP.Situacao, CP.Lote, CP.Movimento, CP.Vencimento, CP.Prorrogacao, " & vbCrLf & _
                      "        CP.DataMoeda, isnull(CP.Baixa,CP.Prorrogacao) as Baixa, CP.UnidadeDeNegocio, CP.Empresa, CP.EndEmpresa, CP.Cliente, " & vbCrLf & _
                      "        CP.EndCliente, CP.BancoCliente, CP.AgenciaCliente, CP.DigitoAgenciaCliente, CP.ContaCliente, " & vbCrLf & _
                      "        CP.DigitoContaCliente, Isnull(CP.TipoContaCliente,'C') AS TipoContaCliente, CP.ContaContabilCliente, " & vbCrLf & _
                      "        CP.EmpresaPagadora, CP.EndEmpresaPagadora, CP.BancoPagador, isnull(CP.AgenciaPagadora,'') AS AgenciaPagadora, CP.DigitoAgenciaPagadora, " & vbCrLf & _
                      "        CP.ContaPagadora, CP.DigitoContaPagadora, CP.ContaContabilPagadora, CP.Cheque, isnull(CP.Slips,'N') AS Slips, " & vbCrLf & _
                      "        CP.Recibo, CP.Aviso, CP.ReciboDeposito, isnull(CP.EmpresaPedido,'') AS EmpresaPedido, isnull(CP.EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(CP.Pedido, 0) AS Pedido, " & vbCrLf & _
                      "        isnull(CP.PedidoFixacao,0) AS PedidoFixacao, isnull(CP.Procuracao,0) AS Procuracao, CP.ValorDoDocumento, CP.Descontos, CP.Deducoes, " & vbCrLf & _
                      "        CP.Juros, CP.Acrescimos, CP.ValorLiquido, ISNULL(CP.MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, " & vbCrLf & _
                      "        ISNULL(CP.MoedaDescontos, 0) AS MoedaDescontos, ISNULL(CP.MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(CP.MoedaJuros, 0) AS MoedaJuros, " & vbCrLf & _
                      "        ISNULL(CP.MoedaAcrescimos, 0) AS MoedaAcrescimos, ISNULL(CP.MoedaValorLiquido, 0) AS MoedaValorLiquido, CP.Historico, " & vbCrLf & _
                      "        CP.CodigoDeBarras, CP.CodigoDigitado, isnull(CP.CodigoDeBarraPreImpresso,0) AS CodigoDeBarraPreImpresso, CP.Destinatario, CP.EndDestinatario, CP.NomeDoDestinatario, CP.Destinacao, " & vbCrLf & _
                      "        CP.Solicitacao, CP.UsuarioInclusao, CP.UsuarioInclusaoData, CP.UsuarioAlteracao, CP.UsuarioAlteracaoData, " & vbCrLf & _
                      "        CP.UsuarioCancelamento, CP.UsuarioCancelamentoData, isnull(CP.UsuarioLiberacao,'') AS UsuarioLiberacao, CP.UsuarioLiberacaoData, " & vbCrLf & _
                      "        CP.UsuarioBaixa, CP.UsuarioBaixaData, isnull(CP.Grupado,'N') AS Grupado, isnull(CP.RegistroMestre, 0) as RegistroMestre, CP.Observacoes, " & vbCrLf & _
                      "        CP.SituacaoBancaria, ISNULL(CP.NumeroDoCheque,0) AS NumeroDoCheque, isnull(ad.adiantamento_id,CP.Adiantamento) as Adiantamento, CP.VencimentoAdto, ISNULL(CP.TaxaAdto,0) TaxaAdto, " & vbCrLf & _
                      "        isnull(CP.UsuarioLiberacaoBloqueio, '') AS UsuarioLiberacaoBloqueio, CP.UsuarioLiberacaoBloqueioDate, isnull(CP.UsuarioLiberacaoPedido, '') AS UsuarioLiberacaoPedido, " & vbCrLf & _
                      "        CP.UsuarioLiberacaoPedidoDate, isnull(CP.UsuarioLiberacaoCheque, '') AS UsuarioLiberacaoCheque, CP.UsuarioLiberacaoChequeDate, " & vbCrLf & _
                      "        isnull(CP.CarteiraAdto,'') AS CarteiraAdto, isnull(CP.CarteiraDoTitulo,0) AS CarteiraDoTitulo, isnull(CP.ContratoDeFinanciamento,'') AS ContratoDeFinanciamento, " & vbCrLf & _
                      "        ISNULL(NF.TipoDeDocumento,0) AS TipoDeDocumento, isnull(ContratoANTT,'') AS ContratoANTT, isnull(Ped.FinanceiroAberto,1) AS FinanceiroAberto, isnull(Ped.MomentoFinanceiro,0) AS MomentoFinanceiro, " & vbCrLf & _
                      "        ISNULL(NF.NFG,0) AS NFG, ISNULL(NF.Conferencia, 1) AS Conferencia, " & vbCrLf & _
                      "        ISNULL(FFxT.Fatura_Id, 0) AS FaturaDeFrete, ISNULL(CP.ContratoBancario, '') AS ContratoBanco, " & vbCrLf & _
                      "        CASE " & vbCrLf & _
                      "			   WHEN ISNULL(BaixaAdiantamento.ValorOficial,0) = 0 " & vbCrLf & _
                      "				    THEN 'S' " & vbCrLf & _
                      "				    ELSE 'N' " & vbCrLf & _
                      "		   END AS LiberaAdiantamento" & vbCrLf & _
                      "   FROM ContasAPagar CP" & vbCrLf & _
                      "	  LEFT JOIN NotaFiscalXTitulo NFxT " & vbCrLf & _
                      "     ON CP.Registro_Id = NFxT.Titulo_Id " & vbCrLf & _
                      "	  LEFT JOIN NotasFiscais NF " & vbCrLf & _
                      "     ON NFxT.Empresa_Id       = NF.Empresa_Id " & vbCrLf & _
                      "	   AND NFxT.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                      "    AND NFxT.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                      "    AND NFxT.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                      "    AND NFxT.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                      "    AND NFxT.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                      "    AND NFxT.Nota_Id         = NF.Nota_Id" & vbCrLf & _
                      "   LEFT JOIN FaturaDeFreteXTitulo FFxT " & vbCrLf & _
                      "     ON CP.Registro_Id = FFxT.Titulo_Id" & vbCrLf & _
                      "   LEFT JOIN Pedidos Ped" & vbCrLf & _
                      "     ON Ped.Empresa_Id    = CP.EmpresaPedido " & vbCrLf & _
                      "    AND Ped.EndEmpresa_Id = CP.EndEmpresaPedido " & vbCrLf & _
                      "    AND Ped.Pedido_Id     = CP.Pedido" & vbCrLf & _
                      "   Left Join Adiantamentos ad" & vbCrLf & _
                      "     on ad.Titulo = CP.Registro_id" & vbCrLf & _
                      "   LEFT JOIN (SELECT aXb.Empresa_id, aXb.EndEmpresa_Id, aXb.RegistroPedido, sum(aXb.ValorOficial) AS ValorOficial" & vbCrLf & _
                      "	               FROM AdiantamentosXBaixas aXb" & vbCrLf & _
                      "				  GROUP BY aXb.Empresa_id, aXb.EndEmpresa_Id, aXb.RegistroPedido) BaixaAdiantamento" & vbCrLf & _
                      "	    ON BaixaAdiantamento.Empresa_id      = CP.Empresa" & vbCrLf & _
                      "	   AND BaixaAdiantamento.EndEmpresa_id  = CP.EndEmpresa" & vbCrLf & _
                      "	   AND BaixaAdiantamento.RegistroPedido = CP.Pedido" & vbCrLf & _
                      "  WHERE CP.Registro_Id = " & Registro

                Dim dsContasAPagar As New DataSet
                dsContasAPagar = Banco.ConsultaDataSet(Sql, "ContasAPagar")

                If dsContasAPagar Is Nothing OrElse dsContasAPagar.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Registro não encontrado")
                    Exit Sub
                End If

                Dim Dr As DataRow = dsContasAPagar.Tables(0).Rows(0)

                Session("ControleCP" & HID.Value) = Dr("Registro_Id").ToString + ";" + Dr("ValorDoDocumento").ToString + ";" + Dr("MoedaValorDoDocumento").ToString + ";" + Dr("Moeda").ToString

                DdlUnidadeDeNegocioEmpresaCliente.SelectedIndex = DdlUnidadeDeNegocioEmpresaCliente.Items.IndexOf(DdlUnidadeDeNegocioEmpresaCliente.Items.FindByValue(Dr("UnidadeDeNegocio")))
                CargaEmpresaCliente()
                DdlEmpresaCliente.SelectedIndex = DdlEmpresaCliente.Items.IndexOf(DdlEmpresaCliente.Items.FindByValue(Dr("Empresa") & "-" & CStr(Dr("EndEmpresa"))))

                Dim strCliente() As String = ConsultaCLientes(Dr("Cliente"), Dr("EndCliente")).Split("-")

                txtFornecedor.Text = ConsultaCLientes(Dr("Cliente"), Dr("EndCliente")) 'strCliente(0) & " - " & strCliente(1) & "-" & strCliente(2) & " - " & strCliente(3)
                txtCodigoFornecedor.Value = Dr("Cliente") & "-" & CStr(Dr("EndCliente"))

                DdlEmpresaPagadora.SelectedIndex = DdlEmpresaPagadora.Items.IndexOf(DdlEmpresaPagadora.Items.FindByValue(Dr("EmpresaPagadora") & "-" & CStr(Dr("EndEmpresaPagadora"))))
                DdlTiposDePagamentos.SelectedIndex = DdlTiposDePagamentos.Items.IndexOf(DdlTiposDePagamentos.Items.FindByValue(Dr("TipoPagto")))

                If DdlEmpresaPagadora.Text <> "" Then
                    CargaBancoPagador()
                End If

                If Dr("BancoPagador") > 0 AndAlso Dr("AgenciaPagadora").ToString.Length > 0 Then
                    DdlBancoPagador.SelectedIndex = DdlBancoPagador.Items.IndexOf(DdlBancoPagador.Items.FindByValue(Dr("BancoPagador")))

                    BancoPagador()
                    conta = Dr("AgenciaPagadora") & "-" & Dr("DigitoAgenciaPagadora") & "-" & Dr("ContaPagadora") & "-" & Dr("DigitoContaPagadora") & "-" & Dr("ContaContabilPagadora")
                    DdlContaPagadora.SelectedIndex = DdlContaPagadora.Items.IndexOf(DdlContaPagadora.Items.FindByValue(conta))
                End If

                ddlMoeda.SelectedValue = Dr("Moeda")
                ddlIndexador.SelectedValue = Dr("Indexador")

                ddlCarteiraDoTitulo.SelectedValue = Dr("CarteiraDoTitulo")
                DdlProvisoes.SelectedValue = Dr("Provisao")

                ddlCarteiras.SelectedIndex = ddlCarteiras.Items.IndexOf(ddlCarteiras.Items.FindByValue(Dr("Carteira")))
                ConfigurarContas(Dr("Empresa"), Dr("EndEmpresa"), Dr("Moeda"), ddlCarteiras.SelectedValue)

                '*************************************************************************************************
                '*************************************  ADIANTAMENTO  ********************************************
                '*************************************************************************************************
                Dim cart As New CarteiraFinanceira(Dr("Carteira"))

                If cart.isAdiantamento And cart.BaixaAdiantamento Then
                    If ddlMoeda.SelectedValue = 1 Then
                        HDSaldoAdiantamento.Value = Dr("ValorDoDocumento")
                    Else
                        HDSaldoAdiantamento.Value = Dr("MoedaValorDoDocumento")
                    End If
                ElseIf Dr("CarteiraAdto").ToString.Length > 0 Then
                    If ddlMoeda.SelectedValue = 1 Then
                        HDSaldoAdiantamento.Value = Dr("ValorLiquido")
                    Else
                        HDSaldoAdiantamento.Value = Dr("MoedaValorLiquido")
                    End If
                End If


                If cart.isAdiantamento Or cart.BaixaAdiantamento Then
                    DdlCarteirasAdto.Parent.Visible = False
                Else
                    DdlCarteirasAdto.Parent.Visible = True
                End If

                If cart.isAdiantamento And Not cart.BaixaAdiantamento Then
                    txtVencimentoAdto.Parent.Visible = True
                Else
                    txtVencimentoAdto.Parent.Visible = False
                End If

                If cart.isAdiantamento Or cart.BaixaAdiantamento Or Dr("CarteiraAdto").ToString.Length > 0 Then
                    txtNumeroAdto.Parent.Visible = True
                Else
                    txtNumeroAdto.Parent.Visible = False
                End If

                If Dr("CarteiraAdto").ToString.Length > 0 Then
                    DdlCarteirasAdto.SelectedValue = Dr("CarteiraAdto")
                End If

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

                If Dr("UsuarioAlteracao").ToString.Length > 0 Then
                    lblUsuarioIncl.Text = Dr("UsuarioInclusao") & " - " & CDate(Dr("UsuarioInclusaoData")).ToString("dd/MM/yyyy")
                    lblUsuarioAlt.Text = Dr("UsuarioAlteracao") & " - " & CDate(Dr("UsuarioAlteracaoData")).ToString("dd/MM/yyyy")
                Else
                    lblUsuarioIncl.Text = Dr("UsuarioInclusao") & " - " & CDate(Dr("UsuarioInclusaoData")).ToString("dd/MM/yyyy")
                End If
                imgUsuarioIncl.Visible = True
                imgUsuarioAlt.Visible = True

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

                If Dr("Destinatario") = "" Then
                    txtCodigoFavorecido.Value = Dr("Cliente") & "-" & CStr(Dr("EndCliente"))
                    txtFavorecido.Text = ConsultaCLientes(Dr("Cliente"), Dr("EndCliente"))
                Else
                    txtCodigoFavorecido.Value = Dr("Destinatario") & "-" & Dr("EndDestinatario")
                    txtFavorecido.Text = ConsultaCLientes(Dr("Destinatario"), Dr("EndDestinatario"))
                End If

                CargaContasCorrentes()
                DdlBancos.SelectedIndex = DdlBancos.Items.IndexOf(DdlBancos.Items.FindByValue(Dr("BancoCliente")))
                txtAgencia.Text = Dr("AgenciaCliente")
                txtDigitoAgencia.Text = Dr("DigitoAgenciaCliente")
                txtContaCorrente.Text = Dr("ContaCliente")
                txtDigitoDaConta.Text = Dr("DigitoContaCliente")
                ddlTipoConta.SelectedIndex = ddlTipoConta.Items.IndexOf(ddlTipoConta.Items.FindByValue(Dr("TipoContaCliente")))
                txtEmiteCheque.Value = Dr("Cheque")
                txtSlip.Value = Dr("Slips")
                txtNumeroCheque.Text = Dr("NumeroDoCheque")

                txtPedido.Text = Dr("Pedido")
                txtPedidoFixacao.Value = Dr("PedidoFixacao")
                Pedido = Dr("Pedido")

                TemRegistro = "S"

                If Not IsDBNull(Dr("Procuracao")) Then
                    txtCessaoDeCredito.Text = Dr("Procuracao")
                Else
                    txtCessaoDeCredito.Text = ""
                End If

                If Not IsDBNull(Dr("Adiantamento")) Then
                    txtNumeroAdto.Text = Dr("Adiantamento")
                Else
                    txtNumeroAdto.Text = ""
                End If

                If Not IsDBNull(Dr("VencimentoAdto")) Then
                    txtVencimentoAdto.Text = Dr("VencimentoAdto")
                Else
                    txtVencimentoAdto.Text = ""
                End If

                txtTaxaAdto.Text = Dr("TaxaAdto")

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

                Sql = "Select isnull(IndiceFixado,0) as IndiceFixado, DataPedido From Pedidos where Pedido_id = " & Pedido
                Dim dsPedido As DataSet = Banco.ConsultaDataSet(Sql, "Pedidos")

                If Dr("Moeda") = 1 Then
                    txtValorDoDocumento.ForeColor = Drawing.Color.Red
                    txtOficial.ForeColor = Drawing.Color.Blue
                End If

                If Dr("Moeda") = 3 Then
                    txtValorEmMoeda.ForeColor = Drawing.Color.Red
                    txtMoeda.ForeColor = Drawing.Color.Blue

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
                            Dr("ValorDoDocumento") = Funcoes.ConverteParaMoedaOficial(Dr("MoedaValorDoDocumento"), Dr("Indexador"), Dr("Vencimento"))
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
                        ImgCalcular.Enabled = False
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
                        ImgCalcular.Enabled = False
                        DdlProvisoes.Enabled = False
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False

                    ElseIf Dr("Grupado") = "M" Then
                        DdlProvisoes.Enabled = False
                        lnkNovo.Parent.Visible = False
                        ImgCalcular.Enabled = False
                        imgBloqueio.Visible = True
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                        If CType(Dr("Provisao"), eProvisao) = eProvisao.Previsao Then lnkExcluir.Parent.Visible = True

                    ElseIf CType(Dr("Provisao"), eProvisao) = eProvisao.Baixa Then
                        DdlProvisoes.Enabled = False
                        lnkNovo.Parent.Visible = False
                        ImgCalcular.Enabled = False
                        imgBloqueio.Visible = True
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                    ElseIf CType(Dr("Provisao"), eProvisao) = eProvisao.Provisao AndAlso Dr("PedidoFixacao") > 0 Then
                        lnkNovo.Parent.Visible = False
                        ImgCalcular.Enabled = False
                        DdlProvisoes.Enabled = False
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                        MsgBox(Me.Page, "Fixação Pendente Emissão de Nota Fiscal")
                    ElseIf CType(Dr("Provisao"), eProvisao) = eProvisao.Provisao AndAlso Dr("MomentoFinanceiro") = 3 Then
                        lnkNovo.Parent.Visible = False
                        ImgCalcular.Enabled = False
                        DdlProvisoes.Enabled = False
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                        MsgBox(Me.Page, "Titulo para Consumo na Nota Fiscal não pode ser alterado")
                    Else
                        DdlProvisoes.Enabled = True
                        imgBloqueio.Visible = False
                        txtProrrogacao.Enabled = True
                        txtMovimento.Enabled = True
                    End If
                Else
                    cmdPedido.Enabled = True

                    If Dr("Grupado") = "S" Then
                        lnkNovo.Parent.Visible = False
                        ImgCalcular.Enabled = False
                        lnkExcluir.Parent.Visible = False
                        DdlProvisoes.Enabled = False
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                    ElseIf Dr("Grupado") = "M" Then
                        DdlProvisoes.Enabled = False
                        lnkNovo.Parent.Visible = False
                        ImgCalcular.Enabled = False
                        imgBloqueio.Visible = True
                        txtMovimento.Enabled = False
                        txtProrrogacao.Enabled = False
                        If CType(Dr("Provisao"), eProvisao) = eProvisao.Previsao Then lnkExcluir.Parent.Visible = True
                    ElseIf CType(Dr("Provisao"), eProvisao) = eProvisao.Baixa Then
                        DdlProvisoes.Enabled = False
                        lnkNovo.Parent.Visible = False
                        ImgCalcular.Enabled = False
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

                If Dr("Grupado") = "M" Then
                    lblAgrupar.Text = "AP"
                    lnkRelatorioAgrupamento.Parent.Visible = True

                End If

                cmdClientesTitulo.Enabled = True

                If Dr("Pedido") > 0 AndAlso Dr("FinanceiroAberto") = "False" Then
                    lnkNovo.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False
                    MsgBox(Me.Page, "Título com Financeiro Fechado no Pedido não pode ser alterado")
                End If

                ConferenciaNF = CBool(Dr("Conferencia")) AndAlso CBool(Dr("NFG"))


                If lblAgrupar.Text = "AP" Then
                    cmdPedido.Enabled = False

                    Dim Filhos As New ArrayList
                    Dim Mensagem As String = "Agrupamento dos Registros"

                    Sql = "Select Registro_id " & vbCrLf & _
                          "From contasApagar " & vbCrLf & _
                          "Where RegistroMestre = " & txtRegistro.Text

                    Dim dsFilhos As New DataSet
                    dsFilhos = Banco.ConsultaDataSet(Sql, "RegistrosFilhos")

                    If Not dsFilhos Is Nothing AndAlso dsFilhos.Tables(0).Rows.Count > 0 Then
                        For Each drFilho As DataRow In dsFilhos.Tables(0).Rows
                            Filhos.Add(drFilho("Registro_id"))
                            Mensagem &= " - " & drFilho("Registro_id")
                        Next

                        Session("ssRegistros" & HID.Value) = Filhos
                        Session("ssObservacoes" & HID.Value) = Mensagem
                    End If
                End If

                If Pedido > 0 Then
                    cmdClientesTitulo.Enabled = False

                    Sql = "  SELECT Registro_Id AS Registro, Vencimento, Baixa, Historico, ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, 0.0 as Saldo, Provisao " & vbCrLf & _
                          "   FROM ContasAPagar " & vbCrLf & _
                          "  WHERE Pedido = '" & Pedido & "' And Pedido <> 0 AND Situacao = 1 " & vbCrLf & _
                          "  ORDER BY Vencimento" & vbCrLf

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

                'Sql = "  SELECT Adiantamento_Id " & vbCrLf & _
                '      "    FROM AdiantamentosXBaixas " & vbCrLf & _
                '      "   WHERE Titulo = " & txtRegistro.Text & vbCrLf
                ''?
                'For Each DrA As DataRow In Banco.ConsultaDataSet(Sql, "Adiantamentos").Tables(0).Rows
                '    txtNumeroAdto.Text = DrA("Adiantamento_Id")
                'Next

                Sql = " SELECT R.Empresa_Id + '-' + cast(R.EndEmpresa_Id as varchar) as Empresa, R.Conta_Id, R.Cliente_Id, R.EndCliente_Id, " & vbCrLf & _
                      "        PC.Titulo, R.Movimento_Id, R.Lote_Id, ISNULL(R.Produto, '') AS Produto, R.Custo, R.Historico, " & vbCrLf & _
                      "        R.DebitoOficial, R.CreditoOficial, isnull(Conciliacao,'') AS Conciliacao, isnull(DataDaBaixa,CURRENT_TIMESTAMP) AS DataDaBaixa " & vbCrLf & _
                      "   FROM Razao R" & vbCrLf & _
                      "  INNER JOIN PlanoDeContas PC " & vbCrLf & _
                      "     ON R.Conta_Id = PC.Conta_Id " & vbCrLf & _
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
                    Sql = "SELECT Registro_Id, Cliente, EndCliente, Carteira " & vbCrLf & _
                          "  FROM TitulosXDesdobrarFornecedor " & vbCrLf & _
                          " WHERE Registro_Id = " & txtRegistro.Text & vbCrLf

                    Dim dsDesdobrarFornecedor As New DataSet
                    dsDesdobrarFornecedor = Banco.ConsultaDataSet(Sql, "DesdobrarFornecedor")

                    If Not dsDesdobrarFornecedor Is Nothing AndAlso dsDesdobrarFornecedor.Tables(0).Rows.Count > 0 Then
                        Session("objDestinoContabil" & HID.Value) = dsDesdobrarFornecedor.Tables(0).Rows(0).Item("Cliente") & "-" & _
                                                        dsDesdobrarFornecedor.Tables(0).Rows(0).Item("EndCliente") & "-" & _
                                                        dsDesdobrarFornecedor.Tables(0).Rows(0).Item("Carteira")
                    End If
                End If

                txtRegistro.Enabled = False


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
                        ddlCarteiras.Enabled = True
                        DdlTributos.Enabled = True
                        txtHistorico.Enabled = True
                    End If
                End If

                If DdlProvisoes.SelectedValue = 3 AndAlso Not Funcoes.VerificaPermissao("AutorizacoesDePagamentos", "GRAVAR") Then
                    lnkNovo.Parent.Visible = False 'Provisao Bloqueada P/ Alteracao.
                End If

                If Not dsContasAPagar Is Nothing AndAlso dsContasAPagar.Tables(0).Rows.Count > 0 AndAlso Not dsContasAPagar.Tables(0).Rows(0).Item("Situacao") = 1 Then
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
                    txtMestre.Text = "CANCELADO"
                End If

                'NF conferida pelo fiscal
                If Not ConferenciaNF Then
                    ViewState.Add("CONTROLE_FISCAL", "BLOQUEADO")
                End If

                TabContainer1.ActiveTabIndex = 0

            Else
                MsgBox(Me.Page, "Informe o Numero do Registro para consulta")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao consultar registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(ex.Message.ToString())) & ". Entre em contato com o Suporte.")
        End Try
    End Sub

    Protected Sub DdlUnidadeConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresaConsultaTitulos.Items.Clear()

        Sql = " SELECT Cli.Cliente_Id as Codigo, Cli.Endereco_Id, Cli.Reduzido, Cli.Nome, Cli.Cidade, Cli.Estado " & vbCrLf & _
              "   FROM GruposXEmpresas GxE" & vbCrLf & _
              "  INNER JOIN Clientes Cli " & vbCrLf & _
              "     ON GxE.Cliente_Id = Cli.Cliente_Id " & vbCrLf & _
              "    AND GxE.EndCliente_Id = Cli.Endereco_Id" & vbCrLf & _
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

    Private Sub Limpar_ConsultaTitulos(ByVal limparTudo As Boolean)
        If limparTudo Then
            If Not DdlEmpresaConsultaTitulos.SelectedIndex > 0 Then Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeConsultaTitulos, DdlEmpresaConsultaTitulos, True)
            txtClienteConsulta.Text = String.Empty
            txtCodigoClienteConsulta.Value = String.Empty
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

        GridConsultaTitulos.DataSource = Nothing
        GridConsultaTitulos.DataBind()
    End Sub

    Protected Sub DdlBancoPagador_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
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
            Sql = "SELECT Agencia_Id, DigitoAgencia_Id, Conta_Id, DigitoConta_Id, ContaContabil, Observacoes " & vbCrLf & _
                  "  FROM BancosXContas" & vbCrLf & _
                  " WHERE Empresa_Id = '" & Campo(0) & "'" & vbCrLf & _
                  "   AND EndEmpresa_Id  = " & Campo(1) & vbCrLf & _
                  "   AND Banco_Id  = " & DdlBancoPagador.SelectedValue & vbCrLf
            If SomenteAtivas Then
                Sql &= "   AND Ativo = 1 " & vbCrLf
            End If

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "BancosXContas").Tables(0).Rows
                Conta = Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "-" & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & "-" & Dr("ContaContabil")
                Descricao = "AG " & Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "  C/C " & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & " - " & Dr("Observacoes")
                DdlContaPagadora.Items.Add(New ListItem(Descricao, Conta))
            Next

            DdlContaPagadora.Items.Insert(0, "")
            DdlContaPagadora.SelectedIndex = 0
        End If

    End Sub

    Private Sub ValidaValores(ByVal dolarizar As Boolean)
        Dim Zero As Decimal = 0

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

                    If CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text) < 0 Then
                        txtAcrescimos.Text = ((CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text)) * -1).ToString("N2")
                    Else
                        txtDeducoes.Text = ((CDbl(txtValorDoDocumento.Text) - CDbl(txtValorCobrado.Text))).ToString("N2")
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
                Else
                    txtDeducoes.Text = Zero.ToString("N2")
                End If
                If CDec(txtJurosMoeda.Text) > 0 Then
                    txtJuros.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtJurosMoeda.Text), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue), CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd")), 2)
                Else
                    txtJuros.Text = Zero.ToString("N2")
                End If
                If CDec(txtAcrescimosMoeda.Text) > 0 Then
                    txtAcrescimos.Text = FormatNumber(Funcoes.ConverteParaMoedaOficial(CDbl(txtAcrescimosMoeda.Text), IIf(ddlMoeda.SelectedValue = 1, 2, ddlIndexador.SelectedValue), CDate(txtProrrogacao.Text).ToString("yyyy/MM/dd")), 2)
                Else
                    txtAcrescimos.Text = Zero.ToString("N2")
                End If
            End If
        End If

        txtValorCobrado.Text = (CDec(txtValorDoDocumento.Text) + CDec(txtJuros.Text) + CDec(txtAcrescimos.Text) - CDec(txtDescontos.Text) - CDec(txtDeducoes.Text)).ToString("N2")
        txtValorCobradoMoeda.Text = (CDec(txtValorEmMoeda.Text) + CDec(txtJurosMoeda.Text) + CDec(txtAcrescimosMoeda.Text) - CDec(txtDescontosMoeda.Text) - CDec(txtDeducoesMoeda.Text)).ToString("N2")

        If lnkNovo.Parent.Visible = False Then lnkNovo.Parent.Visible = True
    End Sub

    Private Function ConsultaNotasXTitulos(ByVal Titulo As String) As Boolean
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

    Protected Sub ImgCalcular_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ImgCalcular.Click
        Try
            ValidaValores(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlCarteiras_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlCarteiras.SelectedIndexChanged
        If DdlProvisoes.SelectedIndex = 0 Then
            ddlCarteiras.SelectedIndex = 0
            MsgBox(Me.Page, "Selecione Previsão ou Baixa!")
        Else
            Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)
            DdlTributos.Parent.Visible = Not objCarteira.isAdiantamento
            txtNumeroAdto.Parent.Visible = objCarteira.isAdiantamento
            DdlCarteirasAdto.Parent.Visible = Not objCarteira.isAdiantamento
            txtVencimentoAdto.Parent.Visible = (objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento)

            txtVencimentoAdto.Enabled = (objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento)
            txtTaxaAdto.Parent.Visible = (objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento)

            If objCarteira.isAdiantamento Then
                DdlCarteirasAdto.SelectedIndex = 0
            Else
                CargaTributos()
            End If

            If objCarteira.isAdiantamento And Not objCarteira.BaixaAdiantamento Then
                ConsultaSequenciaDeAdiantamento()
            End If

            If Session("objDestinoContabil" & HID.Value) Is Nothing AndAlso DdlProvisoes.SelectedValue = 1 Then
                Dim strCliente() As String = txtCodigoFornecedor.Value.Split("-")
                Dim objCliente As New [Lib].Negocio.Cliente(strCliente(0), strCliente(1))
                If objCliente.DesdobrarFornecedor = True Then
                    ucDestinoContabil.Limpar()
                    Dim parameters As New Dictionary(Of String, Object)
                    parameters.Add("tipo", "P")
                    Popup.ConsultaDeDestinoContabil(Me.Page, "objDestinoContabil" & HID.Value)
                    ucDestinoContabil.Carregar(parameters)
                End If
            End If

            ConfigurarContas(DdlEmpresaCliente.SelectedValue.Split("-")(0), DdlEmpresaCliente.SelectedValue.Split("-")(1), ddlMoeda.SelectedValue, ddlCarteiras.SelectedValue)
        End If
    End Sub

    Private Sub ConsultaSequenciaDeAdiantamento()
        Dim emp() As String = DdlEmpresaCliente.SelectedValue.ToString.Split("-")
        Dim NumAdiantamento As Integer

        If IsNumeric(txtRegistro.Text) AndAlso CInt(txtRegistro.Text) > 0 Then
            Sql = "SELECT isnull(Adiantamento_id,0) as Adiantamento" & vbCrLf & _
                  "  FROM ContasAPagar cp" & vbCrLf & _
                  "  Left join Adiantamentos A" & vbCrLf & _
                  "    on cp.registro_id = a.titulo " & vbCrLf & _
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

        Sql = " SELECT Tributo_Id as Codigo, (Encargos.Descricao + ' - ' + Tributo_Id) as Descricao " & vbCrLf & _
              "   FROM CarteirasXTributos " & vbCrLf & _
              "  INNER JOIN Encargos " & vbCrLf & _
              "     ON CarteirasXTributos.Tributo_ID = Encargos.Encargo_id " & vbCrLf & _
              "  WHERE Carteira_Id = '" & ddlCarteiras.SelectedValue & "'" & vbCrLf & _
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
            Mensagem = "Sábado - Data Inválida para " & Tipo & "..."
            Return False
        End If

        If CDate(pData.Replace("'", "")).DayOfWeek = 0 Then
            Mensagem = "Domingo - Data Inválida para " & Tipo & "..."
            Return False
        End If

        Sql = "  SELECT Descricao" & vbCrLf & _
              "    FROM DatasNaoProgramaveis" & vbCrLf & _
              "   WHERE Empresa_Id = '99999999999999' " & vbCrLf & _
              "     AND EndEmpresa_ID = 0 " & vbCrLf & _
              "     AND Data_ID = '" & CDate(pData).ToString("yyyy/MM/dd") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Mensagem = "Data de " & Tipo & " não programável, Feriado Nacional > " & Dr("Descricao")
            Return False
        Next

        If Empresa <> "" Then
            Sql = "  SELECT Descricao " & vbCrLf & _
                  "    FROM DatasNaoProgramaveis" & vbCrLf & _
                  "   WHERE Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                  "     AND EndEmpresa_ID = " & EndEmpresa & " " & vbCrLf & _
                  "     AND Data_ID = '" & CDate(pData).ToString("yyyy/MM/dd") & "'" & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                Mensagem = "Data de " & Tipo & " não programável, Feriado Municipal > " & Dr("Descricao")
                Return False
            Next
        End If

        If Tipo = "Movimento" Then
            If Funcoes.VerificaAcesso(Empresa, EndEmpresa, pData, "Financeiro") Then
            Else
                Mensagem = "Movimento já Fechado para esta data " & pData & ", para empresa " & Empresa
                Return False
            End If
        End If

        Return True

    End Function

    Private Function DolarizaBaixa(ByVal pData As String, ByVal Valor As String, ByVal Indexador As String) As String
        Dim SqlL As String
        Dim Calculo As Decimal

        SqlL = "SELECT Indice" & _
               "  FROM Cotacoes" & _
               " WHERE Data_Id      ='" & pData & "'" & vbCrLf & _
               "   AND Indexador_Id =" & Indexador

        For Each Dr As DataRow In Banco.ConsultaDataSet(SqlL, "Cot").Tables(0).Rows
            Calculo = CDec(Valor) / IIf(Dr("Indice") = 0, 1, Dr("Indice"))
            'Calculo = CDec(FormatNumber(Calculo, 2))
        Next

        Return Calculo.ToString("N2")
    End Function

    Protected Sub DdlEmpresaPagadora_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        DdlTiposDePagamentos.SelectedIndex = 0
        CargaBancoPagador()
    End Sub

    Protected Sub cmdFavorecido_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo" & HID.Value) = "LivreClasse"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objFavorecidoCP" & HID.Value, "txtNome")
    End Sub

    Private Function Adiantamento(ByVal Pedido As [Lib].Negocio.Pedido, ByVal RegistroTitulo As String, ByVal ValorAdtoOficial As Decimal, ByVal ValorAdtoMoeda As Decimal) As Boolean
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

        Sql = " INSERT INTO Adiantamentos (Empresa_ID, EndEmpresa_ID, Cliente_ID, EndCliente_ID, Adiantamento_Id, RegistroPedido, Titulo, Recibo," & vbCrLf & _
              "                          Safra, Movimento, Vencimento, ValorOficial, ValorMoeda, Indexador, Moeda, Taxa)" & vbCrLf & _
              " VALUES('" & campo(0) & "'," & vbCrLf & _
              "        " & campo(1)

        campo = txtCodigoFornecedor.Value.Split("-")
        Sql &= ", '" & campo(0) & "'" & vbCrLf & _
               ", " & campo(1) & vbCrLf & _
               ", " & RegistroAdiantamento & vbCrLf

        If Pedido Is Nothing OrElse Pedido.Codigo = 0 Then
            Sql &= ", NULL" & vbCrLf
        Else
            Sql &= ", " & Pedido.Codigo & vbCrLf
        End If

        Sql &= ", " & RegistroTitulo & vbCrLf & _
             ", 0" & vbCrLf

        If Pedido Is Nothing OrElse Pedido.Codigo = 0 Then
            Sql &= ", NULL" & vbCrLf
        Else
            Sql &= ", '" & Pedido.CodigoSafra & "'" & vbCrLf
        End If

        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf

        If String.IsNullOrWhiteSpace(txtVencimentoAdto.Text) Then
            Throw New Exception("Titulo " & RegistroTitulo & " com carteira de Adiantamento mais sem a data do Vencimento informado")
        Else
            Sql &= ", '" & CDate(txtVencimentoAdto.Text).ToString("yyyy/MM/dd") & "'"
        End If

        Sql &= ", " & Str(ValorAdtoOficial) & ", " & Str(ValorAdtoMoeda) & vbCrLf & _
               ", " & ddlIndexador.SelectedValue & ", " & ddlMoeda.SelectedValue

        If String.IsNullOrWhiteSpace(txtTaxaAdto.Text) Then
            Sql &= ", 0 )"                                                             'Taxa de Adiantamento 
        Else
            Sql &= ", " & Replace(Replace(txtTaxaAdto.Text, ".", ""), ",", ".") & ")"  'Taxa
        End If

        SqlArray.Add(Sql)

        Sql = " UPDATE ContasApagar " & vbCrLf & _
              "    SET Adiantamento = " & RegistroAdiantamento & vbCrLf & _
              "  WHERE Registro_id = " & RegistroTitulo

        SqlArray.Add(Sql)

        Return True
    End Function

    Private Function AdiantamentoAmortizacao(ByVal Pedido As [Lib].Negocio.Pedido, ByVal RegistroTitulo As String, ByVal NumeroAdto As String, ByVal ValorBxAdtoOficial As Decimal, ByVal ValorBxAdtoMoeda As Decimal) As Boolean
        Dim objAdiantamento As New [Lib].Negocio.Adiantamento()
        Cliente = DdlEmpresaCliente.SelectedValue
        campo = Cliente.Split("-")
        objAdiantamento.CodigoEmpresa = campo(0)
        objAdiantamento.EndEmpresa = campo(1)
        Cliente = txtCodigoFornecedor.Value
        campo = Cliente.Split("-")
        objAdiantamento.CodigoCliente = campo(0)
        objAdiantamento.EndCliente = campo(1)
        objAdiantamento.Codigo = CInt(NumeroAdto)

        objAdiantamento = New [Lib].Negocio.Adiantamento(0, CInt(NumeroAdto), DdlEmpresaCliente.SelectedValue.ToString.Split("-")(0), DdlEmpresaCliente.SelectedValue.ToString.Split("-")(1))

        If objAdiantamento.Codigo = 0 Then
            Mensagem = "Adiantamento informado não foi encontrado, verifique a lista."
            Return False
        End If

        Dim Sequencia As Integer

        Sqla = "  SELECT ISNULL(MAX(Sequencia_Id), 0) + 1 AS Sequencia  " & vbCrLf & _
               "    FROM AdiantamentosXBaixas" & vbCrLf & _
               "   WHERE Empresa_Id      ='" & objAdiantamento.CodigoEmpresa & "'" & vbCrLf & _
               "     AND EndEmpresa_Id   ='" & objAdiantamento.EndEmpresa & "'" & vbCrLf & _
               "     AND Adiantamento_Id = " & NumeroAdto & vbCrLf


        For Each Dr As DataRow In Banco.ConsultaDataSet(Sqla, "Encargos").Tables(0).Rows
            Sequencia = Dr("Sequencia")
        Next

        Sql = " INSERT INTO AdiantamentosXBaixas" & _
                    " (Empresa_ID " & _
                    ", EndEmpresa_ID" & _
                    ", Cliente_ID" & _
                    ", EndCliente_ID" & _
                    ", Adiantamento_Id" & _
                    ", Sequencia_Id" & _
                    ", RegistroPedido" & _
                    ", Titulo" & _
                    ", ValorOficial" & _
                    ", ValorMoeda" & _
                    ", VariacaoOficial" & _
                    ", VariacaoMoeda" & _
                    ", DataBaixa)" & _
                    " VALUES('" & objAdiantamento.CodigoEmpresa & "'" & _
                    "," & objAdiantamento.EndEmpresa

        campo = txtCodigoFornecedor.Value.Split("-")
        Sql &= ", '" & objAdiantamento.CodigoCliente & "'"                          'Cliente
        Sql &= ", " & objAdiantamento.EndCliente                                    'Endereco Cliente

        Sql &= ", " & NumeroAdto                                                    'Numero do Adiantamentos
        Sql &= ", " & Sequencia                                                     'Sequencia

        If Pedido Is Nothing OrElse Pedido.Codigo = 0 Then                          'Numero do Pedido
            Sql &= ", NULL" & vbCrLf
        Else
            Sql &= ", " & Pedido.Codigo & vbCrLf
        End If

        Sql &= ", " & RegistroTitulo                                                'Titulo

        Sql &= ", " & Str(ValorBxAdtoOficial)
        Sql &= ", " & Str(ValorBxAdtoMoeda)


        Sql &= ", 0"
        Sql &= ", 0"
        Sql &= ", '" & CDate(txtMovimento.Text).ToString("yyyy/MM/dd") & "')"

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

            Popup.ConsultaDePedidos(Me.Page, "objPedidoCTAPAG" & HID.Value)
            ucConsultaPedidos.BindGridView(parameters)
        End If
    End Sub

    Private Function PesoPago(ByVal Pedido As String, ByVal Valor As String, ByVal Historico As String) As String
        Dim HistoricoParcial As String = Historico
        Dim Empresa() As String = DdlEmpresaCliente.SelectedValue.ToString.Split("-")

        Sqla = "SELECT pif.Pedido_Id," & vbCrLf & _
               "       pif.Produto_Id," & vbCrLf & _
               "       Produtos.Nome," & vbCrLf & _
               "       Produtos.Agrupar, " & vbCrLf & _
               "       pif.Fixacao_Id," & vbCrLf & _
               "       pif.Quantidade," & vbCrLf & _
               "       isnull(Pedidos.PedidoEfetivo,'') as PedidoEfetivo," & vbCrLf & _
               "       ISNULL((SELECT SUM(T1.ValorOficial) AS Oficial " & vbCrLf & _
               "                 FROM VW_PedidosXItensXFixacoesXEncargos AS T1 " & vbCrLf & _
               "                WHERE T1.Encargo_Id = 'LIQUIDO'" & vbCrLf & _
               "                  AND T1.Pedido_Id  = pif.Pedido_Id), 0) AS Oficial," & vbCrLf & _
               "       ISNULL((SELECT SUM(T1.ValorMoeda) AS Moeda " & vbCrLf & _
               "                 FROM VW_PedidosXItensXFixacoesXEncargos AS T1 " & vbCrLf & _
               "                WHERE T1.Encargo_Id = 'LIQUIDO'" & vbCrLf & _
               "                  AND T1.Pedido_Id  = pif.Pedido_Id), 0) AS Moeda " & vbCrLf & _
               "  FROM VW_PedidosXItensXFixacoes pif" & vbCrLf & _
               " INNER JOIN Produtos" & vbCrLf & _
               "    ON pif.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
               " INNER JOIN Pedidos" & vbCrLf & _
               "    ON pif.Empresa_Id    = Pedidos.Empresa_Id" & vbCrLf & _
               "   AND pif.EndEmpresa_Id = Pedidos.EndEmpresa_Id" & vbCrLf & _
               "   AND pif.Pedido_Id     = Pedidos.Pedido_Id " & vbCrLf & _
               " WHERE Pedidos.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf & _
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

    Private Sub DiaGeral()
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
        parameters.Add("Relatorio", "Relação de Títulos A Pagar")
        Dim objEmpresa As [Lib].Negocio.Cliente = New [Lib].Negocio.Cliente(HttpContext.Current.Session("ssEmpresa"), HttpContext.Current.Session("ssEndEmpresa"))
        parameters.Add("EmpresaNome", objEmpresa.Nome)
        parameters.Add("EmpresaCidade", objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado)
        parameters.Add("EmpresaCodigo", Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
        parameters.Add("UnidadeDeNegocio", Parametros)
        parameters.Add("TipoDaCarteira", "Carteira")


        Funcoes.BindReport(Me.Page, ds, crystal, eExportType.PDF, parameters)
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
        sql = "  SELECT CP.Registro_Id AS Registro, " & vbCrLf & _
              "         convert(varchar(10),CP.Vencimento,103) as Vencimento,convert(varchar(10)," & vbCrLf & _
              "         CP.Prorrogacao,103) as Posterga, Cli.Nome AS Cliente, Historico, " & vbCrLf & _
              "         isnull(CP.MoedaValorLiquido, 0) AS Dolar, " & vbCrLf & _
              "         CP.ValorLiquido AS Valor, " & vbCrLf & _
              "         UsuarioLiberacao as Liberado, " & vbCrLf & _
              "         CP.Carteira As Carteira, " & vbCrLf & _
              "         ComprasXProdutos.Descricao As DescricaoCarteira, " & vbCrLf & _
              "         CP.Situacao as Situacao, " & vbCrLf & _
              "         CP.Empresa as Empresa, " & vbCrLf & _
              "         Empresa.Reduzido as Reduzido, " & vbCrLf & _
              "         CP.UsuarioBaixa as UsuarioBaixa" & vbCrLf & _
              "    FROM ContasAPagar CP " & vbCrLf & _
              "   INNER JOIN Clientes Cli " & vbCrLf & _
              "      ON CP.Cliente = Cli.Cliente_Id " & vbCrLf & _
              "     AND CP.EndCliente = Cli.Endereco_Id" & vbCrLf & _
              "    LEFT JOIN ComprasXProdutos " & vbCrLf & _
              "      ON CP.Carteira = ComprasXProdutos.Produto_id " & vbCrLf & _
              "   INNER JOIN Clientes as Empresa " & vbCrLf & _
              "      ON CP.Empresa = Empresa.Cliente_Id " & vbCrLf & _
              "     AND CP.EndEmpresa = Empresa.Endereco_Id" & vbCrLf & _
              "   WHERE CP.Provisao <> 1 " & vbCrLf & _
              "     AND CP.Situacao = 1" & vbCrLf

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
            sql &= " and CP.Empresa = '" & Campo(0) & "'" & vbCrLf & _
                   " and CP.EndEmpresa = " & Campo(1) & vbCrLf
        End If

        Campo = txtCodigoClienteConsulta.Value.Split("-")
        If Campo(0) <> "" Then
            sql &= " and CP.Cliente = '" & Campo(0) & "'" & vbCrLf & _
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
        sql = " SELECT CP.Registro_Id AS Registro, " & vbCrLf & _
              "        convert(varchar(10),CP.Vencimento,103) as Vencimento,convert(varchar(10)," & vbCrLf & _
              "        CP.Prorrogacao,103) as Posterga, Clientes.Nome AS Cliente, Historico, " & vbCrLf & _
              "        isnull(CP.MoedaValorLiquido, 0) AS Dolar, " & vbCrLf & _
              "        CP.ValorLiquido AS Valor, " & vbCrLf & _
              "        UsuarioLiberacao as Liberado, " & vbCrLf & _
              "        CP.Carteira As Carteira, " & vbCrLf & _
              "        ComprasXProdutos.Descricao As DescricaoCarteira, " & vbCrLf & _
              "        CP.Situacao as Situacao, " & vbCrLf & _
              "        CP.Empresa as Empresa, " & vbCrLf & _
              "        Empresa.Reduzido as Reduzido, " & vbCrLf & _
              "        CP.UsuarioBaixa as UsuarioBaixa" & vbCrLf & _
              "   FROM ContasAPagar CP " & vbCrLf & _
              "  INNER JOIN Clientes " & vbCrLf & _
              "     ON CP.Cliente = Clientes.Cliente_Id " & vbCrLf & _
              "    AND CP.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
              "   LEFT JOIN ComprasXProdutos " & vbCrLf & _
              "     ON CP.Carteira = ComprasXProdutos.Produto_id " & vbCrLf & _
              "  INNER JOIN Clientes as Empresa " & vbCrLf & _
              "     ON CP.Empresa = Empresa.Cliente_Id" & vbCrLf & _
              "    AND CP.EndEmpresa = Empresa.Endereco_Id" & vbCrLf & _
              "  WHERE CP.Provisao <> 1 " & vbCrLf & _
              "    AND CP.Situacao = 1" & vbCrLf

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
            sql &= " and CP.Empresa = '" & Campo(0) & "'" & vbCrLf & _
                   " and CP.EndEmpresa = " & Campo(1) & vbCrLf
        End If

        Campo = txtCodigoClienteConsulta.Value.Split("-")
        If Campo(0) <> "" Then
            sql &= " and CP.Cliente = '" & Campo(0) & "'" & vbCrLf & _
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
            Dim Tdigito As String = ""
            Dim TDigitoAgencia As String = ""

            Sql = "SELECT Baixa, ISNULL(NumeroDoCheque, 0) AS NumeroDoCheque, DigitoAgenciaCliente,  DigitoContaCliente" & vbCrLf & _
                  "  FROM ContasAPagar CP " & vbCrLf & _
                  " WHERE Registro_Id = " & txtRegistro.Text & vbCrLf

            For Each Dr1 As DataRow In Banco.ConsultaDataSet(Sql, "ContasAPagar").Tables(0).Rows
                DataBaixa = IIf(Dr1("Baixa") Is Nothing, "", CDate(Dr1("Baixa")).ToString("dd/MM/yyyy"))
                TnumeroDoCheque = Dr1("NumeroDoCheque")
                Tdigito = Dr1("DigitoContaCliente")
                TDigitoAgencia = Dr1("DigitoAgenciaCliente")
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
            Sql = "SELECT Emp.Cliente_Id ," & vbCrLf & _
                  "       Emp.Nome, Emp.Cidade," & vbCrLf & _
                  "       Emp.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                  "       Emp.Endereco , Emp.Cep," & vbCrLf & _
                  "       Emp.Inscricao, Emp.Telefone," & vbCrLf & _
                  "       Emp.Bairro, Emp.Complemento," & vbCrLf & _
                  "       Emp.Numero " & vbCrLf & _
                  "  FROM Clientes Emp " & vbCrLf & _
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
            Sql = " SELECT Cli.Cliente_Id ," & vbCrLf & _
                  "        Cli.Nome, Cli.Cidade," & vbCrLf & _
                  "        Cli.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                  "        Cli.Endereco , Cli.Cep," & vbCrLf & _
                  "        Cli.Inscricao, Cli.Telefone," & vbCrLf & _
                  "        Cli.Bairro, Cli.Complemento," & vbCrLf & _
                  "        Cli.Numero " & vbCrLf & _
                  "   FROM Clientes Cli " & vbCrLf & _
                  "  WHERE Cli.Cliente_Id = '" & Cliente & "'" & vbCrLf & _
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
            row("TBanco") = DdlBancos.SelectedValue
            row("TAgencia") = txtAgencia.Text
            row("TConta") = txtContaCorrente.Text
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
            row("TDigito") = Tdigito
            row("TDigitoAgencia") = TDigitoAgencia

            ''tvalordodocumento
            dtRecibo.Rows.Add(row)

            Dim param As New Dictionary(Of String, Object)
            param.Add("XNome", ENome)

            Funcoes.BindReport(Me.Page, dsRecibo, "Cr_ReciboPagar", eExportType.PDF, param)

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
            If DdlCarteirasAdto.SelectedIndex = 0 Then
                cart = New CarteiraFinanceira(ddlCarteiras.SelectedValue)
            Else
                cart = New CarteiraFinanceira(DdlCarteirasAdto.SelectedValue)
            End If
            Parametros.Add("ContaContabil", cart.CodigoContaCliente)
            Parametros.Add("ContaContabilDescricao", cart.CodigoContaCliente + " - " + cart.Descricao)

            Parametros.Add("Moeda", ddlMoeda.SelectedValue)
            Parametros.Add("DescMoeda", ddlMoeda.SelectedItem.Text)
            Parametros.Add("Formulario", "Financeiro")

            Session("Parametros" & HID.Value) = Parametros
            ucConsultaAdiantamentos.BindGridView()
            Popup.ConsultaDeAdiantamentos(Me.Page, "Financeiro" & HID.Value)
        End If
    End Sub


    Protected Sub imgBloqueio_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Cliente = DdlEmpresaCliente.SelectedValue
        campo = Cliente.Split("-")

        If Not Funcoes.VerificaPermissao("ContasAPagar", "LIBERAR") Then
            MsgBox(Me.Page, "Usuário sem permissão para liberar Registro")
        ElseIf Not Funcoes.VerificaAcesso(campo(0), campo(1), txtMovimento.Text, "FINANCEIRO") Then
            MsgBox(Me.Page, "Movimento já Fechado para esta data " & txtMovimento.Text & ", para empresa " & campo(0))
        Else
            If txtNumeroAdto.Text.Length > 0 Then
                Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)
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
                        MsgBox(Me.Page, "Registro com Baixa de Adiantamento no Título " & objAdtoXBaixa(0).CodigoTitulo & ", volte o mesmo para PREVISÃO. Com a alteração desse Registro " & txtRegistro.Text & " será gerado um novo número de Adiantamento, no Registro " & objAdtoXBaixa(0).CodigoTitulo & " vincule novamente o novo número do Adiantamento gerado.")
                        Exit Sub
                    End If
                End If
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
            ImgCalcular.Enabled = True
            lnkNovo.Parent.Visible = True
            ddlIndexador.Enabled = True
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
    End Sub

    Protected Sub imgLimparPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLimparPedido.Click
        If Funcoes.VerificaPermissao("ContasAPagar", "LIBERAR") Then
            'txtLiberarPedido.Value = "S"
            'txtPedido.Text = "0"
            'cmdPedido.Enabled = True
            MsgBox(Me.Page, "Desabilitado, entre em contato com o Suporte.")
        Else
            MsgBox(Me.Page, "Usuário sem permissão para remover o Pedido.")
        End If
    End Sub

    Protected Sub imgLimparCheque_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLimparCheque.Click
        If Funcoes.VerificaPermissao("ContasAPagar", "LIBERAR") Then
            txtLiberarCheque.Value = "S"
            txtEmiteCheque.Value = "N"
            txtNumeroCheque.Text = "0"
            txtNumeroCheque.Enabled = True
        Else
            MsgBox(Me.Page, "Usuário sem permissão para liberar Registro")
        End If
    End Sub

    Protected Sub imgLimparAdto_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLimparAdto.Click
        If Funcoes.VerificaPermissao("ContasAPagar", "LIBERAR") Then
            Dim cart As New CarteiraFinanceira(ddlCarteiras.SelectedValue)

            If Not (cart.isAdiantamento And Not cart.BaixaAdiantamento) Then
                txtNumeroAdto.Text = "0"
                HDSaldoAdiantamento.Value = 0
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
            Extrato.Emitir(Me.Page, FinanceiroNovo, DdlEmpresaCliente.SelectedValue.Split("-")(0), DdlEmpresaCliente.SelectedValue.Split("-")(1), "T", _
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


    Protected Sub DdlCarteirasAdto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlCarteirasAdto.SelectedIndexChanged
        If DdlCarteirasAdto.SelectedIndex = 0 Then
            txtNumeroAdto.Parent.Visible = False
            txtNumeroAdto.Text = ""
        Else
            txtNumeroAdto.Parent.Visible = True
            Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlCarteiras.SelectedValue)
            If objCarteira.isAdiantamento AndAlso DdlCarteirasAdto.SelectedIndex > 0 Then
                ddlCarteiras.SelectedIndex = 0
                DdlCarteirasAdto.SelectedIndex = 0
                'txtPedido.Text = "" 
                txtNumeroAdto.Text = ""
                MsgBox(Me.Page, "Carteira de Compensação ou Baixa de Adiantamento não pode ser usada com Carteira Principal de Adiantamento.")
                TabContainer1.ActiveTabIndex = 0
            End If
        End If
    End Sub

    Protected Sub DdlProvisoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If DdlProvisoes.SelectedIndex = 0 Then Exit Sub

        If Not DdlProvisoes.SelectedValue = 1 AndAlso txtProrrogacao.Enabled = False Then
            txtProrrogacao.Enabled = True
        End If

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
                        parameters.Add("tipo", "P")
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

                    Sql = "UPDATE ContasAPagar " & vbCrLf & _
                          "   Set Prorrogacao = '" & CDate(txtNovoVencimento.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                          "     , UsuarioAlteracao = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf & _
                          "     , UsuarioAlteracaoData = GETDATE()" & vbCrLf & _
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

        Sql = " SELECT     Situacao_Id, Descricao, Provisao " & vbCrLf & _
              "   FROM SituacoesBancariasRetorno " & vbCrLf & _
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

    Protected Sub txtProrrogacao_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtProrrogacao.TextChanged
        Try
            MostrarCotacao()

            If (Not String.IsNullOrWhiteSpace(txtRegistro.Text)) AndAlso (Not String.IsNullOrWhiteSpace(ddlMoeda.SelectedValue)) Then
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
                        Dim vlrDiff As Decimal = CDec(txtValorCobradoMoeda.Text) - CDec(txtValorEmMoeda.Text)
                        If (vlrDiff > 0) Then
                            txtAcrescimosMoeda.Text = String.Format("{0:N2}", vlrDiff)
                        Else
                            txtDeducoesMoeda.Text = String.Format("{0:N2}", (vlrDiff * -1))
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("ContasAPagar", "GRAVAR") Then
            Try
                GravaTitulo()
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        Else
            MsgBox(Me.Page, "Usuário sem permissão para incluir registro!")
        End If
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        If Funcoes.VerificaPermissao("ContasAPagar", "LEITURA") Then
            Registro = txtRegistro.Text
            Limpar(True)
            txtRegistro.Text = Registro
            ConsultaContasAPagar()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        chkManterLancamento.Checked = False
        Limpar(True)
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        EmitirRecibo()
    End Sub

    Protected Sub lnkRelatorioAgrupamento_Click(sender As Object, e As EventArgs) Handles lnkRelatorioAgrupamento.Click
        Try
            If String.IsNullOrWhiteSpace(txtRegistro.Text) Then
                MsgBox(Me.Page, "É necessário informar o campo número do registro!")
                Exit Sub
            End If
            viewAgrupamento(txtRegistro.Text.Trim())
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ContasAPagar")
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
        If Funcoes.VerificaPermissao("ContasAPagar", "RELATORIO") Then

            Dim crpt As New ReportDocument()

            Try
                If RbDiaGeral.Checked = True Then
                    DiaGeral()
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

    Protected Sub lnkLimparConsultaTitulo_Click(sender As Object, e As EventArgs) Handles lnkLimparConsultaTitulo.Click
        Limpar_ConsultaTitulos(True)
    End Sub

    Protected Sub lnkAgruparPagamento_Click(sender As Object, e As EventArgs) Handles lnkAgruparPagamento.Click
        Try
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

            Sql = "SELECT UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, Indexador," & vbCrLf & _
                  "       Sum(isnull(MoedaValorDoDocumento,0)) as MoedaValorDoDocumento," & vbCrLf & _
                  "       Sum(ValorDoDocumento) as ValorDoDocumento," & vbCrLf & _
                  "       Sum(Descontos) as Descontos," & vbCrLf & _
                  "       Sum(Deducoes) as Deducoes, " & vbCrLf & _
                  "       Sum(Juros) as Juros," & vbCrLf & _
                  "       Sum(Acrescimos) as Acrescimos," & vbCrLf & _
                  "       Sum(ValorLiquido) as ValorLiquido, " & vbCrLf & _
                  "       Sum(MoedaValorLiquido) as MoedaValorLiquido" & vbCrLf & _
                  "  FROM ContasAPagar " & vbCrLf & _
                  " WHERE Registro_Id = 99999999 " & vbCrLf

            Dim i As Integer = 0
            While i < GridConsultaTitulos.Rows.Count
                If CType(GridConsultaTitulos.Rows(i).FindControl("ChkGridTitulos"), CheckBox).Checked = True Then
                    Mensagem &= " - " & GridConsultaTitulos.Rows(i).Cells(3).Text()
                    Sql &= " or Registro_Id = " & GridConsultaTitulos.Rows(i).Cells(3).Text()
                    Registros.Add(GridConsultaTitulos.Rows(i).Cells(3).Text())
                End If
                i += 1
            End While

            Sql &= " Group By UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, Indexador " & vbCrLf
            Dim ds As New DataSet

            ds = Banco.ConsultaDataSet(Sql, "ContasAPagar")

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Agrupamento não pode ser realizado.")
                Exit Sub
            End If

            Sql = "Select cp.Registro_Id, cp.Carteira, cXp.Descricao, cp.Tributo, isnull(cXp.ContaClientes,'') as ContaClientes, " & vbCrLf & _
                  "       case " & vbCrLf & _
                  "       	when LEN(isnull(cXp.ContaClientes,'')) = 0 and LEN(isnull(cp.Tributo,'')) = 0 " & vbCrLf & _
                  "           then 'S' " & vbCrLf & _
                  "       	  else 'N' " & vbCrLf & _
                  "       end as FaltaTributo, " & vbCrLf & _
                  "       case " & vbCrLf & _
                  "       	when isnull(cXp.BaixaAdiantamento,0) = 1 " & vbCrLf & _
                  "           then 'S' " & vbCrLf & _
                  "       	  else 'N' " & vbCrLf & _
                  "       end as BaixaAdiantamento" & vbCrLf & _
                  "  from ContasAPagar cp " & vbCrLf & _
                  " inner join ComprasXProdutos cXp " & vbCrLf & _
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
                CargaContasCorrentes()

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
                    Else
                        Dr("ValorLiquido") = Funcoes.ConverteParaMoedaOficial(Dr("MoedaValorLiquido"), ddlIndexador.SelectedValue, CDate(txtProrrogacao.Text).ToString("yyyy-MM-dd"))
                    End If
                End If

                txtValorDoDocumento.Text = CDec(Dr("ValorLiquido")).ToString("N2")
                txtValorEmMoeda.Text = CDec(Dr("MoedaValorLiquido")).ToString("N2")
                txtValorCobrado.Text = CDec(Dr("ValorLiquido")).ToString("N2")

                txtDescontos.Text = "0,00"
                txtDeducoes.Text = "0,00"
                txtJuros.Text = "0,00"
                txtAcrescimos.Text = "0,00"

                txtValorDoDocumento.Enabled = False
                txtValorEmMoeda.Enabled = False
                txtDescontos.Enabled = False
                txtDeducoes.Enabled = False
                txtJuros.Enabled = False
                txtAcrescimos.Enabled = False
                txtValorCobrado.Enabled = False
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
                Sql = "SELECT Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf & _
                      "       UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente,Isnull(TipoContaCliente,'C') AS TipoContaCliente, " & vbCrLf & _
                      "       ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf & _
                      "       ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, isnull(EmpresaPedido,'') AS EmpresaPedido, isnull(EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(Pedido,0) AS Pedido, isnull(PedidoFixacao,0) AS PedidoFixacao, isnull(Procuracao,0) AS Procuracao, ValorDoDocumento, " & vbCrLf & _
                      "       Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, ISNULL(MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, ISNULL(MoedaDescontos, 0) " & vbCrLf & _
                      "       AS MoedaDescontos, ISNULL(MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(MoedaJuros, 0) AS MoedaJuros, ISNULL(MoedaAcrescimos, 0) AS MoedaAcrescimos, " & vbCrLf & _
                      "       ISNULL(MoedaValorLiquido, 0) AS MoedaValorLiquido, (case when len(convert(nvarchar(4000),Observacoes)) > 0 then Historico + ' ' + convert(nvarchar(4000),Observacoes ) else Historico end) as Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, " & vbCrLf & _
                      "       solicitacao, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, " & vbCrLf & _
                      "       UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, isnull(RegistroMestre, 0) as RegistroMestre, Observacoes, SituacaoBancaria,T.Descricao, ISNULL(NumeroDoCheque, 0) AS NumeroDoCheque  " & vbCrLf & _
                      "  FROM ContasAPagar " & vbCrLf & _
                      " INNER JOIN TIPOSDEPAGAMENTOS AS T " & vbCrLf & _
                      "    ON T.TipoDePagamento_ID = TipoPagto " & vbCrLf & _
                      " WHERE Registro_Id = " & Registro

                dsCP = Banco.ConsultaDataSet(Sql, "ContasAPagar")
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
                Sql = " SELECT Emp.Cliente_Id ," & vbCrLf & _
                      "        Emp.Nome, Emp.Cidade," & vbCrLf & _
                      "        Emp.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                      "        Emp.Endereco , Emp.Cep," & vbCrLf & _
                      "        Emp.Inscricao, Emp.Telefone," & vbCrLf & _
                      "        Emp.Bairro, Emp.Complemento," & vbCrLf & _
                      "        Emp.Numero " & vbCrLf & _
                      "   FROM Clientes Emp" & vbCrLf & _
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
                Sql = " SELECT Cli.Cliente_Id ," & vbCrLf & _
                      "        Cli.Nome, Cli.Cidade," & vbCrLf & _
                      "        Cli.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                      "        Cli.Endereco , Cli.Cep," & vbCrLf & _
                      "        Cli.Inscricao, Cli.Telefone," & vbCrLf & _
                      "        Cli.Bairro, Cli.Complemento," & vbCrLf & _
                      "        Cli.Numero " & vbCrLf & _
                      "   FROM Clientes Cli " & vbCrLf & _
                      "  WHERE Cli.Cliente_Id = '" & Cliente & "'" & vbCrLf & _
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

        Sql = "SELECT Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf & _
              "       UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, Isnull(TipoContaCliente,'C') AS TipoContaCliente, " & vbCrLf & _
              "       ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf & _
              "       ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, isnull(EmpresaPedido,'') AS EmpresaPedido, isnull(EndEmpresaPedido,0) AS EndEmpresaPedido, isnull(Pedido,0) AS Pedido, isnull(PedidoFixacao,0) AS PedidoFixacao, isnull(Procuracao,0) AS Procuracao, ValorDoDocumento, " & vbCrLf & _
              "       Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, ISNULL(MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, ISNULL(MoedaDescontos, 0) " & vbCrLf & _
              "       AS MoedaDescontos, ISNULL(MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(MoedaJuros, 0) AS MoedaJuros, ISNULL(MoedaAcrescimos, 0) AS MoedaAcrescimos, " & vbCrLf & _
              "       ISNULL(MoedaValorLiquido, 0) AS MoedaValorLiquido, (case when len(convert(nvarchar(4000),Observacoes)) > 0 then Historico + ' ' + convert(nvarchar(4000),Observacoes ) else Historico end) as Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, " & vbCrLf & _
              "       solicitacao, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, " & vbCrLf & _
              "       UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, isnull(RegistroMestre, 0) as RegistroMestre, Observacoes, SituacaoBancaria,isnull(T.Descricao, 'TIPO PGTO NAO DEFINIDO') AS Descricao, ISNULL(NumeroDoCheque, 0) AS NumeroDoCheque  " & vbCrLf & _
              "  FROM ContasAPagar " & vbCrLf & _
              "  LEFT JOIN TIPOSDEPAGAMENTOS AS T ON T.TipoDePagamento_ID = TipoPagto " & vbCrLf & _
              " WHERE Slips = 'N'" & vbCrLf

        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            Sql &= "   and UnidadeDeNegocio = '" & Cliente & "' "
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            Sql &= "   AND Empresa = '" & Campo(0) & "'" & vbCrLf & _
                   "   AND EndEmpresa = " & Campo(1) & vbCrLf
        End If

        If Left(Session("ssEmpresa").ToString, 8) = "05272759" Then
            Sql &= "   AND CASE " & vbCrLf & _
                   "   	       WHEN len(ISNULL(usuariobaixa,'')) = 0" & vbCrLf & _
                   "               THEN" & vbCrLf & _
                   "   		           CASE " & vbCrLf & _
                   "   			           WHEN LEn(ISNULL(UsuarioAlteracao,'')) = 0" & vbCrLf & _
                   "   			               THEN UsuarioInclusao" & vbCrLf & _
                   "   			               ELSE UsuarioAlteracao" & vbCrLf & _
                   "   		           END" & vbCrLf & _
                   "   	           ELSE UsuarioBaixa" & vbCrLf & _
                   "        END = '" & Session("ssNomeUsuario") & "' " & vbCrLf
        End If

        Sql &= " ORDER BY Registro_Id " & vbCrLf

        Dim dsRegistro As New DataSet
        dsRegistro = Banco.ConsultaDataSet(Sql, "RegistroContasAPagar")

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
                Sql = "  SELECT Emp.Cliente_Id ," & vbCrLf & _
                      "         Emp.Nome, Emp.Cidade," & vbCrLf & _
                      "         Emp.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                      "         Emp.Endereco , Emp.Cep," & vbCrLf & _
                      "         Emp.Inscricao, Emp.Telefone," & vbCrLf & _
                      "         Emp.Bairro, Emp.Complemento," & vbCrLf & _
                      "         Emp.Numero " & vbCrLf & _
                      "    FROM Clientes Emp " & vbCrLf & _
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

                Sql = " SELECT Cli.Cliente_Id ," & vbCrLf & _
                      "        Cli.Nome, Cli.Cidade," & vbCrLf & _
                      "        Cli.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                      "        Cli.Endereco , Cli.Cep," & vbCrLf & _
                      "        Cli.Inscricao, Cli.Telefone," & vbCrLf & _
                      "        Cli.Bairro, Cli.Complemento," & vbCrLf & _
                      "        Cli.Numero " & vbCrLf & _
                      "   FROM Clientes Cli " & vbCrLf & _
                      "  WHERE Cli.Cliente_Id = '" & Cliente & "'" & vbCrLf & _
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

                Sql = " UPDATE ContasAPagar" & vbCrLf & _
                      "    SET Slips = 'S' " & vbCrLf & _
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
                ConsultaContasAPagar()
                DdlProvisoes.SelectedValue = 1 'marca para baixar
                lnkNovo_Click(Nothing, Nothing) 'Baixa

                'Titulos Liquidados
                Sql = " UPDATE ContasAPagar " & vbCrLf & _
                      "    SET SituacaoRemessaBancaria = 204" & vbCrLf & _
                      "  WHERE Registro_id in (" & GridRetornoTitulos.Rows(i).Cells(2).Text() & ")" & vbCrLf
                SqlArray.Add(Sql)
                Banco.GravaBanco(SqlArray)
            Else

                'Titulos Rejeitados
                Sql = "  UPDATE ContasAPagar " & vbCrLf & _
                      "     SET SituacaoRemessaBancaria = 203" & vbCrLf & _
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


    Private Function ValidaContaBancaria() As Boolean
        If DdlBancos.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Banco não foi selecionado.")
            Return False
        ElseIf txtAgencia.Text.Length = 0 Then
            MsgBox(Me.Page, "Agência não foi informada.")
            Return False
        ElseIf txtContaCorrente.Text.Length = 0 Then
            MsgBox(Me.Page, "Conta não foi informada.")
            Return False
        ElseIf txtDigitoDaConta.Text.Length = 0 Then
            MsgBox(Me.Page, "Dígito da Conta não foi informado.")
            Return False
        Else
            Return True
        End If
    End Function

#Region "Dados Bancarios"
    Protected Sub lnkGravarDadosBancarios_Click(sender As Object, e As EventArgs) Handles lnkGravarDadosBancarios.Click
        Dim SqlArray As New ArrayList
        Dim Mensagem As String

        If ValidaContaBancaria() Then
            Dim Campo = txtCodigoFavorecido.Value.Split("-")

            Sql = " INSERT Into ClientesXContasBancarias(Cliente_Id, Endereco_Id, Banco_Id, Agencia_Id, DigitoAgencia_Id, ContaCorrente_Id, DigitoConta_Id, TipoConta, Observacoes) " & vbCrLf & _
                  " Values ('" & Campo(0) & "'," & CStr(Campo(1)) & ", " & DdlBancos.SelectedValue & ", " & txtAgencia.Text.Trim & ", '" & txtDigitoAgencia.Text.Trim & "', " & vbCrLf & _
                            txtContaCorrente.Text.Trim & "," & "'" & txtDigitoDaConta.Text.Trim & "', '" & ddlTipoConta.SelectedValue & "', '" & txtObservacoesDaConta.Text & "')" & vbCrLf

            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                Mensagem = Session("ssMessage")
                MsgBox(Me.Page, Mensagem)
            Else

                LimparDadosBancarios()
                CargaContasCorrentes()
                TabContainer1.ActiveTabIndex = 0
            End If
        End If

    End Sub

    Protected Sub lnkAlterarDadosBancarios_Click(sender As Object, e As EventArgs) Handles lnkAlterarDadosBancarios.Click
        Dim SqlArray As New ArrayList
        Dim Mensagem As String

        If ValidaContaBancaria() Then

            Dim Campo = txtCodigoFavorecido.Value.Split("-")

            Sql = "UPDATE ClientesXContasBancarias SET" & vbCrLf & _
                  "    Observacoes ='" & txtObservacoesDaConta.Text & "'" & vbCrLf & _
                  "   ,TipoConta   ='" & ddlTipoConta.SelectedValue & "'" & vbCrLf & _
                  " WHERE Cliente_Id       ='" & Campo(0) & "'" & vbCrLf & _
                  "   AND Endereco_Id      = " & CStr(Campo(1)) & vbCrLf & _
                  "   AND Banco_Id         = " & DdlBancos.SelectedValue & vbCrLf & _
                  "   AND Agencia_Id       = " & txtAgencia.Text.Trim & vbCrLf & _
                  "   AND DigitoAgencia_ID ='" & txtDigitoAgencia.Text.Trim & "'" & vbCrLf & _
                  "   AND ContaCorrente_Id = " & txtContaCorrente.Text.Trim & vbCrLf & _
                  "   AND DigitoConta_Id   ='" & txtDigitoDaConta.Text.Trim & "'" & vbCrLf
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                Mensagem = Session("ssMessage")
                MsgBox(Me.Page, Mensagem)
            Else
                LimparDadosBancarios()
                CargaContasCorrentes()
                TabContainer1.ActiveTabIndex = 0
            End If
        End If
    End Sub

    Protected Sub lnkExcluirDadosBancarios_Click(sender As Object, e As EventArgs) Handles lnkExcluirDadosBancarios.Click
        Dim SqlArray As New ArrayList
        Dim Mensagem As String

        If ValidaContaBancaria() Then

            Dim Campo = txtCodigoFavorecido.Value.Split("-")

            Sql = "DELETE ClientesXContasBancarias" & vbCrLf & _
                  " WHERE Cliente_Id       ='" & Campo(0) & "'" & vbCrLf & _
                  "   AND Endereco_Id      = " & CStr(Campo(1)) & vbCrLf & _
                  "   AND Banco_Id         = " & DdlBancos.SelectedValue & vbCrLf & _
                  "   AND Agencia_Id       = " & txtAgencia.Text.Trim & vbCrLf & _
                  "   AND DigitoAgencia_ID ='" & txtDigitoAgencia.Text.Trim & "'" & vbCrLf & _
                  "   AND ContaCorrente_Id = " & txtContaCorrente.Text.Trim & vbCrLf & _
                  "   AND DigitoConta_Id   ='" & txtDigitoDaConta.Text.Trim & "'" & vbCrLf

            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                Mensagem = Session("ssMessage")
                MsgBox(Me.Page, Mensagem)
            Else
                LimparDadosBancarios()
                CargaContasCorrentes()
            End If
        End If
    End Sub

    Protected Sub lnkLimparDadosBancarios_Click(sender As Object, e As EventArgs) Handles lnkLimparDadosBancarios.Click
        LimparDadosBancarios()
    End Sub
#End Region


    Private Function getDataSetAgrupamento(ByVal registro As String) As DataSet
        Dim ds As New DataSet
        Dim sql As String = "SELECT CP2.Registro_Id as Mestre_Id, " & vbCrLf & _
                            "       CP2.Provisao as Mestre_Provisao, " & vbCrLf & _
                            "       CP2.Vencimento as Mestre_Vencimento, " & vbCrLf & _
                            "       CP2.Prorrogacao as Mestre_Prorrogacao, " & vbCrLf & _
                            "       CP2.ValorDoDocumento as Mestre_ValorDoDocumento, " & vbCrLf & _
                            "       CP2.Deducoes as Mestre_Deducoes, " & vbCrLf & _
                            "       CP2.Acrescimos as Mestre_Acrescimos, " & vbCrLf & _
                            "       CP2.ValorLiquido as Mestre_ValorLiquido, " & vbCrLf & _
                            "       CP2.Cliente as Mestre_Cliente," & vbCrLf & _
                            "       CLI2.Nome as Mestre_NomeCliente," & vbCrLf & _
                            "       CP1.Registro_Id, " & vbCrLf & _
                            "       CP1.Provisao, " & vbCrLf & _
                            "       CP1.Grupado, " & vbCrLf & _
                            "       CP1.Vencimento, " & vbCrLf & _
                            "       CP1.Prorrogacao, " & vbCrLf & _
                            "       CP1.ValorDoDocumento, " & vbCrLf & _
                            "       CP1.Deducoes, " & vbCrLf & _
                            "       CP1.Acrescimos, " & vbCrLf & _
                            "       CP1.ValorLiquido, " & vbCrLf & _
                            "       CP1.Cliente as Cliente," & vbCrLf & _
                            "       CLI1.Nome as NomeCliente," & vbCrLf & _
                            "       CAST(FFxI.Nota_Id as VARCHAR) + '-' + CAST(FFxI.Serie_Id as varchar) as DACTE " & vbCrLf & _
                            "  FROM ContasAPagar cp1 " & vbCrLf & _
                            " INNER JOIN Clientes	cli1" & vbCrLf & _
                            "    ON CLI1.Cliente_Id = CP1.Cliente" & vbCrLf & _
                            "   AND CLI1.Endereco_Id = CP1.EndCliente" & vbCrLf & _
                            "  LEFT JOIN ContasAPagar CP2  " & vbCrLf & _
                            "    ON (CP1.RegistroMestre = CP2.Registro_Id) " & vbCrLf & _
                            "  LEFT JOIN Clientes cli2 " & vbCrLf & _
                            "    ON CLI2.Cliente_Id = cp2.Cliente " & vbCrLf & _
                            "   AND CLI2.Endereco_Id = cp2.EndCliente " & vbCrLf & _
                            "  LEFT JOIN FaturaDeFreteXTitulo FFxT" & vbCrLf & _
                            "    ON CP1.Registro_Id = FFxT.Titulo_Id" & vbCrLf & _
                            "  LEFT JOIN FaturasDeFretesXItens FFxI" & vbCrLf & _
                            "    ON FFxT.Empresa_Id       = FFxI.EmpresaPagadora_Id " & vbCrLf & _
                            "   AND FFxT.EndEmpresa_Id    = FFxI.EndEmpresaPagadora_Id" & vbCrLf & _
                            "   AND FFxT.Conveniado_Id    = FFxI.Conveniado_Id" & vbCrLf & _
                            "   AND FFxT.EndConveniado_Id = FFxI.EndConveniado_Id" & vbCrLf & _
                            "   AND FFxT.Fatura_Id        = FFxI.Fatura_Id" & vbCrLf & _
                            " WHERE 1=1 " & vbCrLf & _
                            "   AND CP1.Situacao = 1 " & vbCrLf & _
                            "   AND CP1.Grupado <> 'M' " & vbCrLf & _
                            "   AND CP1.Provisao = 1 " & vbCrLf & _
                            "   AND CP2.Registro_Id IS NOT NULL " & vbCrLf

        If Not String.IsNullOrWhiteSpace(registro) Then
            sql &= "AND cp1.Registro_Id in (SELECT Registro_Id FROM ContasApagar WHERE 1=1 AND RegistroMestre  = '" & registro & "')"
        End If

        ds = Banco.ConsultaDataSet(sql, "ContasAPagar")
        Return ds
    End Function

    Private Function getCabecalho() As String
        Dim cab As String = String.Empty

        If Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
            Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0), DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(1))
            cab = objEmpresa.Nome & "(" & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo) & ")" & vbCrLf & _
            objEmpresa.Endereco & ", " & objEmpresa.Numero & vbCrLf & _
            objEmpresa.CEP & " - " & objEmpresa.Cidade & "-" & objEmpresa.CodigoEstado
        End If
        Return cab
    End Function

    Private Sub viewAgrupamento(ByVal registroMestre As String)
        Try
            If Funcoes.VerificaPermissao("ContasAPagar", "RELATORIO") Then
                Dim parameters As New Dictionary(Of String, Object)
                parameters.Add("Titulo", "Relatório Agrupamento De Contas À Pagar.")

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

        If txtCodigoClienteConsulta.Value.Length > 0 Then
            Parametros &= "Cliente " & IIf(chkClientes.Checked, " - Consolidado:", ":") & txtClienteConsulta.Text & vbCrLf
        End If

        Sql = "SELECT Empresas.Reduzido AS ReduzidoEmpresa, Titulos.Empresa, Titulos.EndEmpresa, Empresas.Nome AS NomeEmpresa" & vbCrLf & _
              "       ,Empresas.Nome AS NomeEmpresa, Empresas.Cidade AS CidadeEmpresa, Empresas.Estado AS EstadoEmpresa, " & vbCrLf & _
              "       convert(nvarchar,Titulos.Registro_Id) + ' ' +" & vbCrLf & _
              "       Case" & vbCrLf & _
              "         when isnull(Titulos.Moeda,1) = 1 " & vbCrLf & _
              "           then 'R$'" & vbCrLf & _
              "           Else 'U$'" & vbCrLf & _
              "       end Registro," & vbCrLf & _
              "       '' as Faturamento," & vbCrLf & _
              "       '' as Lote, " & vbCrLf & _
              "       0 as LoteTotal," & vbCrLf & _
              "       0 as LoteEntregue," & vbCrLf & _
              "       Titulos.Pedido," & vbCrLf & _
              "       Titulos.Cliente, Clientes.Nome AS NomeCliente," & vbCrLf & _
              "       Titulos.Movimento, " & vbCrLf & _
              "       Titulos.Prorrogacao AS Vencimento, " & vbCrLf & _
              "       Titulos.Baixa, " & vbCrLf & _
              "       Titulos.Carteira, " & vbCrLf & _
              "       Titulos.Provisao, " & vbCrLf & _
              "       Carteira.Descricao AS NomeCarteira, " & vbCrLf & _
              "       Titulos.Historico" & IIf(chkObservacao.Checked, "+ ' / OBS: ' + cast(Titulos.Observacoes as varchar) as Historico,", ",") & vbCrLf & _
              "       Titulos.solicitacao, " & vbCrLf & _
              "       Titulos.ValorLiquido, " & vbCrLf & _
              "       Titulos.MoedaValorLiquido, " & vbCrLf & _
              "       (SELECT ISNULL(P.PedidoEfetivo,0) FROM Pedidos P WHERE P.Pedido_id = Titulos.Pedido AND P.Empresa_Id = Titulos.Empresa AND P.EndEmpresa_Id = Titulos.EndEmpresa) PedidoEfetivo, " & vbCrLf & _
              "       (SELECT 'P' AS Tipo FROM ContasAPagar WHERE Registro_id = Titulos.Registro_id " & vbCrLf & _
              "         UNION " & vbCrLf & _
              "        SELECT 'R' AS Tipo FROM ContasAReceber WHERE Registro_id = Titulos.Registro_id" & vbCrLf & _
              "       ) AS Tipo" & vbCrLf & _
              "  FROM ContasAPagar AS Titulos" & vbCrLf & _
              "  LEFT Join cotacoes" & vbCrLf & _
              "    ON Cotacoes.Data_id      = Titulos.Prorrogacao " & vbCrLf & _
              "   AND Cotacoes.Indexador_Id = Titulos.Indexador " & vbCrLf & _
              " INNER JOIN Clientes AS Empresas " & vbCrLf & _
              "    ON Titulos.Empresa    = Empresas.Cliente_Id " & vbCrLf & _
              "   AND Titulos.EndEmpresa = Empresas.Endereco_Id " & vbCrLf & _
              " INNER JOIN Clientes " & vbCrLf & _
              "    ON Titulos.Cliente    = Clientes.Cliente_Id " & vbCrLf & _
              "   AND Titulos.EndCliente = Clientes.Endereco_Id " & vbCrLf & _
              "  LEFT OUTER JOIN ComprasXProdutos AS Carteiras " & vbCrLf & _
              "    ON Titulos.Carteira = Carteiras.Produto_Id" & vbCrLf & _
              "  LEFT OUTER JOIN Carteira " & vbCrLf & _
              "    ON Titulos.CarteiraDoTitulo = Carteira.Carteira_Id " & vbCrLf & _
              "   AND (Titulos.Situacao = 1) " & vbCrLf & _
              "   AND (Titulos.Grupado <> 'M') " & vbCrLf & _
              "   AND (Titulos.ContaContabilCliente NOT LIKE '1010101%') " & vbCrLf & _
              "   AND (Titulos.ContaContabilCliente NOT LIKE '1010102%') " & vbCrLf & _
              "  LEFT JOIN NotaFiscalXTitulo NFXT" & vbCrLf & _
              "    ON Titulos.Empresa      = NFXT.Empresa_Id" & vbCrLf & _
              "   AND Titulos.EndEmpresa  = NFXT.EndEmpresa_Id" & vbCrLf & _
              "   AND Titulos.Registro_Id = NFXT.Titulo_Id" & vbCrLf & _
              "  LEFT JOIN FaturaDeFreteXTitulo FFxT" & vbCrLf & _
              "    ON Titulos.Registro_Id = FFxT.Titulo_Id" & vbCrLf & _
              "  LEFT JOIN FaturasDeFretesXItens FFxI" & vbCrLf & _
              "    ON FFxT.Empresa_Id       = FFxI.EmpresaPagadora_Id " & vbCrLf & _
              "   AND FFxT.EndEmpresa_Id    = FFxI.EndEmpresaPagadora_Id" & vbCrLf & _
              "   AND FFxT.Conveniado_Id    = FFxI.Conveniado_Id" & vbCrLf & _
              "   AND FFxT.EndConveniado_Id = FFxI.EndConveniado_Id" & vbCrLf & _
              "   AND FFxT.Fatura_Id        = FFxI.Fatura_Id" & vbCrLf

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

        If chkBaixado.Checked Then
            Sql &= "   AND Titulos.Provisao = 1 " & vbCrLf
            Parametros &= "Titulos Baixados" & vbCrLf
        End If

        If RbCancelado.Checked Then
            Sql &= "   AND Titulos.Situacao = 3 " & vbCrLf
            Parametros &= "Titulos Cancelados" & vbCrLf
        ElseIf RbAtivo.Checked Then
            Sql &= "   AND Titulos.Situacao = 1 " & vbCrLf
            Parametros &= "Titulos Ativos" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlUnidadeConsultaTitulos.SelectedValue) Then
            Sql &= "   AND Titulos.UnidadeDeNegocio = '" & DdlUnidadeConsultaTitulos.SelectedValue & "' " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
            Sql &= "   AND Titulos.Empresa = '" & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                   "   AND Titulos.EndEmpresa = " & DdlEmpresaConsultaTitulos.SelectedValue.Split("-")(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoClienteConsulta.Value) Then
            If chkClientes.Checked Then
                Sql &= "   AND left(Titulos.Cliente, 8) = '" & Left(txtCodigoClienteConsulta.Value.Split("-")(0), 8) & "'" & vbCrLf
            Else
                Sql &= "   AND Titulos.Cliente = '" & txtCodigoClienteConsulta.Value.Split("-")(0) & "'" & vbCrLf & _
                       "   AND Titulos.EndCliente = " & txtCodigoClienteConsulta.Value.Split("-")(1) & vbCrLf
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtNumNota.Text) Then
            Sql &= "   AND (ISNULL(NFXT.Nota_Id,0) in(" & txtNumNota.Text & ") OR ISNULL(FFxI.Fatura_Id,0) in(" & txtNumNota.Text & ") OR ISNULL(FFxI.Nota_Id,0) in(" & txtNumNota.Text & "))"
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
            If Funcoes.VerificaPermissao("ContasAPagar", "EXCLUIR") Then
                If validaExclusao() Then
                    Cliente = DdlEmpresaCliente.SelectedValue
                    campo = Cliente.Split("-")
                    If Funcoes.VerificaAcesso(campo(0), campo(1), txtMovimento.Text, "Financeiro") = False Then
                        MsgBox(Me.Page, "Movimento já Fechado para esta data " & txtMovimento.Text & ", para empresa " & campo(0))
                    Else
                        Dim TemAdiantamento As Boolean = False
                        Sqla = " SELECT  isnull(Adiantamento, 'N') as Adiantamento FROM ComprasXProdutos" & _
                                            " WHERE Produto_Id = '" & ddlCarteiras.SelectedValue & "'"
                        For Each Dra As DataRow In Banco.ConsultaDataSet(Sqla, "Encargos").Tables(0).Rows
                            If Dra("Adiantamento") = "S" Then
                                TemAdiantamento = True
                            End If
                        Next

                        If DdlCarteirasAdto.SelectedIndex > 0 Then TemAdiantamento = True

                        SqlArray.Clear()

                        If txtRegistro.Text = "" Then
                            MsgBox(Me.Page, "Informe o número do Registro para Excluir...")
                        ElseIf txtPedido.Text <> "" And txtPedido.Text > 0 And TemAdiantamento = False Then
                            MsgBox(Me.Page, "Título vinculado à Pedido não pode ser excluído...")
                        ElseIf txtNumeroCheque.Text > 0 Then
                            MsgBox(Me.Page, "Título com emissão de cheque não pode ser excluído...")
                        ElseIf DdlProvisoes.SelectedValue = 1 AndAlso Not lblAgrupar.Text = "AP" Then
                            MsgBox(Me.Page, "Título baixado não pode ser excluído")
                        ElseIf ConsultaNotasXTitulos(txtRegistro.Text) Then   '**********************************************Nao excluie pois tem vinculo em NotasxTitulos.
                            'Registro Não pode se Excluido, Existe uma Nota " & Nota & " Vinculada a este.
                        Else
                            Sql = " UPDATE ContasAPagar " & vbCrLf & _
                                  " SET Situacao = 3" & vbCrLf

                            Sql &= ", UsuarioCancelamento = '" & Session("ssNomeUsuario") & "'"                     'Usuario que Cancelou
                            Sql &= ", UsuarioCancelamentoData = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"   'Data do Cancelamento

                            If txtLiberarTitulo.Value = "S" Then
                                Sql &= ", UsuarioLiberacaoBloqueio = '" & Session("ssNomeUsuario") & "'" 'Usuario que Liberou Titulo
                                Sql &= ", UsuarioLiberacaoBloqueioDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"       'Data da Liberação do Titulo
                            End If

                            If txtLiberarPedido.Value = "S" Then
                                Sql &= ", UsuarioLiberacaoPedido = '" & Session("ssNomeUsuario") & "'"   'Usuario que Liberou Pedido
                                Sql &= ", UsuarioLiberacaoPedidoDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"         'Data da Liberação do Pedido
                            End If

                            If txtLiberarCheque.Value = "S" Then
                                Sql &= ", UsuarioLiberacaoCheque = '" & Session("ssNomeUsuario") & "'"   'Usuario que Liberou Cheque
                                Sql &= ", UsuarioLiberacaoChequeDate = '" & CDate(Today).ToString("yyyy/MM/dd") & "'"         'Data da Liberação do Cheque
                            End If

                            Sql &= " WHERE Registro_ID = " & txtRegistro.Text & vbCrLf

                            SqlArray.Add(Sql)

                            Sql = "DELETE FROM razao " & vbCrLf & _
                                 " WHERE Titulo = " & txtRegistro.Text & vbCrLf
                            SqlArray.Add(Sql)

                            Sql = "DELETE FROM Adiantamentos " & vbCrLf & _
                                  " WHERE Titulo = " & txtRegistro.Text & vbCrLf
                            SqlArray.Add(Sql)

                            Sql = "DELETE FROM AdiantamentosXBaixas" & vbCrLf & _
                                  " WHERE Titulo = " & txtRegistro.Text & vbCrLf
                            SqlArray.Add(Sql)

                            Sql = "DELETE TitulosXDesdobrarFornecedor " & vbCrLf & _
                                  " WHERE Registro_Id = " & txtRegistro.Text & vbCrLf

                            SqlArray.Add(Sql)

                            Sql = "SELECT Registro_id " & vbCrLf & _
                                  "  FROM contasApagar " & vbCrLf & _
                                  " WHERE RegistroMestre = " & txtRegistro.Text

                            Dim dsMestre As New DataSet
                            dsMestre = Banco.ConsultaDataSet(Sql, "Registros")

                            If Not dsMestre Is Nothing AndAlso dsMestre.Tables(0).Rows.Count > 0 Then
                                For Each drFilho As DataRow In dsMestre.Tables(0).Rows
                                    Sql = " UPDATE ContasAPagar" & vbCrLf & _
                                          "    SET Situacao = 1" & vbCrLf & _
                                          "       ,Provisao = 2" & vbCrLf & _
                                          "       ,Grupado = 'N'" & vbCrLf & _
                                          "       ,RegistroMestre = 0" & vbCrLf & _
                                          "       ,UsuarioCancelamento = '" & Session("ssNomeUsuario") & "'" & vbCrLf & _
                                          "       ,UsuarioCancelamentoData = '" & CDate(Today).ToString("yyyy/MM/dd") & "'" & vbCrLf

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

                                    Sql = "DELETE FROM Razao " & vbCrLf & _
                                          " WHERE Titulo = " & drFilho("Registro_id") & vbCrLf

                                    SqlArray.Add(Sql)

                                    Sql = "DELETE FROM Adiantamentos" & vbCrLf & _
                                          " WHERE Titulo = " & drFilho("Registro_id") & vbCrLf

                                    SqlArray.Add(Sql)

                                    Sql = "DELETE FROM AdiantamentosXBaixas" & vbCrLf & _
                                          " WHERE Titulo = " & drFilho("Registro_id") & vbCrLf

                                    SqlArray.Add(Sql)
                                Next
                            End If

                            If Banco.GravaBanco(SqlArray) = False Then
                                MsgBox(Me.Page, "Erro ao Salvar Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                            Else
                                MsgBox(Me.Page, "Registro < " & txtRegistro.Text & " > Excluido com sucesso.", eTitulo.Sucess)
                                Limpar(True)
                            End If
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub
End Class