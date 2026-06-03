Imports System.IO
Imports System.Data
Imports System.Net
Imports System.Net.Sockets
Imports System.Drawing.Printing
Imports System.ServiceModel
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PesagemDaOrdemDeCarregamento
    Inherits BasePage

#Region "Atributos / Propriedades"
    Private sql As String
    Private Client As TcpClient
#End Region

#Region "Eventos"

#Region "Page_Load"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If IsConnect AndAlso Not IsPostBack Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Expedicao.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("PesagemDaOrdemDeCarregamento", "ACESSAR") Then
                CarregarBalanca()
                IncluirEventos()
                BuscarSituacao()
                Dim dsImpressora As New DataSet
                sql = "Select Descricao FROM RotinasDeImpressao WHERE Rotina_Id = 'LAUDO'"
                dsImpressora = Banco.ConsultaDataSet(sql, "Impressora")

                If dsImpressora Is Nothing Then
                ElseIf dsImpressora.Tables(0).Rows.Count = 0 Then
                Else
                    HttpContext.Current.Session("printerName") = dsImpressora.Tables(0).Rows(0).Item("Descricao")
                End If
                'ddl.Carregar(ddlOperacao, CarregarDDL.Tabela.Operacao, "")
                'CarregarTabelaClassificacao()
                LimparCampos("", "")
                Exit Sub
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Expedicao.aspx", eTitulo.Info)
            End If
        End If
    End Sub

#End Region

#Region "Botões"

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        hdnControlePopup.Value = "Empresa"
        ucConsultaEmpresas.Limpar()
        Popup.ConsultaDeEmpresas(Me.Page, "objEmpresa" & HID.Value)
    End Sub

    'Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCliente.Click
    '    hdnControlePopup.Value = "Cliente"
    '    ucConsultaClientes.Limpar()
    '    ucConsultaClientes.SetarTipoCliente("4,5")
    '    ucConsultaClientes.SetarTituloDIV("Consulta de Clientes")
    '    Popup.ConsultaDeClientes(Me.Page, "objPesagem" & HID.Value)
    '    Popup.CenterDialog(Me.Page, "divConsultaCliente")
    'End Sub

    'Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPedido.Click
    '    Limpar()
    '    hdnControlePopup.Value = "Pedido"
    '    CarregarPopupPedido(True)
    'End Sub

    'Protected Sub btnDepositante_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDepositante.Click
    '    hdnControlePopup.Value = "Depositante"
    '    ucConsultaClientes.Limpar()
    '    ucConsultaClientes.SetarTipoCliente("4,5")
    '    ucConsultaClientes.SetarTituloDIV("Consulta de Depositante")
    '    Popup.ConsultaDeClientes(Me.Page, "objPesagem" & HID.Value)
    '    Popup.CenterDialog(Me.Page, "divConsultaCliente")
    'End Sub

    Protected Sub btnTransportador_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnTransportador.Click
        hdnControlePopup.Value = "Transportador"
        ucConsultaClientes.Limpar()
        ucConsultaClientes.SetarTipoCliente("7")
        ucConsultaClientes.SetarTituloDIV("Consulta de Transportador")
        Popup.ConsultaDeClientes(Me.Page, "objLaudo" & HID.Value)
        Popup.CenterDialog(Me.Page, "divConsultaCliente")
    End Sub

    Protected Sub btnPlaca_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPlaca.Click
        hdnControlePopup.Value = "Placa"
        ucConsultaPlacas.Limpar()
        Popup.ConsultaDePlacas(Me.Page, "objLaudo" & HID.Value)
        Popup.CenterDialog(Me.Page, "divConsultaPlacas")
    End Sub

    'Protected Sub btnDeposito_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDeposito.Click
    '    hdnControlePopup.Value = "Deposito"
    '    ucConsultaClientes.Limpar()
    '    ucConsultaClientes.SetarTipoCliente("3")
    '    ucConsultaClientes.SetarTituloDIV("Consulta de Deposito")
    '    Popup.ConsultaDeClientes(Me.Page, "objPesagem" & HID.Value)
    '    Popup.CenterDialog(Me.Page, "divConsultaCliente")
    'End Sub

    Protected Sub btnIncluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnIncluir.Click
        Try
            If Funcoes.VerificaPermissao("PesagemDaOrdemDeCarregamento", "GRAVAR") Then
                Dim msg As String = ValidarCampos("I")
                If (String.IsNullOrEmpty(msg)) Then
                    If (String.IsNullOrEmpty(txtLaudo.Text)) Then
                        Salvar("I")
                    Else
                        Salvar("U")
                    End If
                Else
                    btnIncluir.Enabled = Not String.IsNullOrWhiteSpace(txtPrimeiraPesagem.Text) AndAlso (String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) OrElse Not CDec(txtSegundaPesagem.Text) > 0)
                    MsgBox(Me.Page, msg)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar pesagem!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPesagem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPesagem.Click
        If Funcoes.VerificaPermissao("PesagemDaOrdemDeCarregamento", "GRAVAR") Then
            'Dim consolida As Boolean = chkConsolidarCliente.Checked
            LimparCampos()
            'chkConsolidarCliente.Checked = consolida
            'btnCliente.Visible = True
            'btnPedido.Visible = True
            'btnDepositante.Visible = True
            'btnDeposito.Visible = True
            btnTransportador.Visible = True
            ddlSituacao.Enabled = False
            btnIncluir.Enabled = False
            btnEncerrar.Enabled = False
            btnCancelar.Enabled = False
            btnReimprimir.Enabled = False
            btnConsultar.Enabled = False
            btnEmpresa.Visible = False
            btnConsultar.Visible = False
            btnPesagem.Enabled = False
            'btnCliente_Click(btnCliente, Nothing)
            btnTransportador_Click(btnTransportador, Nothing)
        Else
            MsgBox(Me.Page, "Usuário sem permissão para pesagem!")
        End If
    End Sub

    Protected Sub btnConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultar.Click
        If Funcoes.VerificaPermissao("PesagemDaOrdemDeCarregamento", "LEITURA") Then
            If (String.IsNullOrEmpty(txtLaudo.Text)) Then
                CarregarGridLaudo()
            Else
                CarregarLaudo(txtLaudo.Text, txtCnpjDaEmpresa.Text.Replace(".", "").Replace("/", "").Replace("-", ""), hdnCodigoEndEmpresa.Value)
                btnPesagem.Enabled = False
                btnEncerrar.Enabled = False
                btnIncluir.Enabled = False
                btnCancelar.Enabled = True
                btnReimprimir.Enabled = True
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para pesagem!")
        End If
    End Sub

    Protected Sub btnLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnLimpar.Click
        LimparCampos()
    End Sub

    'Protected Sub btnCalcular_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCalcular.Click
    '    If Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) AndAlso Not txtSegundaPesagem.Text.Equals("0") Then
    '        'If String.IsNullOrWhiteSpace(ddlSubOperacao.SelectedValue) Then
    '        '    MsgBox(Me.Page, "Operação não foi selecionada!")
    '        '    Exit Sub
    '        'End If

    '        Dim libera As Boolean = True
    '        For i = 0 To gridDescontos.Rows.Count - 1
    '            CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Enabled = False
    '            CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Enabled = False
    '            CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Enabled = False
    '        Next

    '        If libera Then
    '            CalcularDesconto()
    '        Else
    '            For i = 0 To gridDescontos.Rows.Count - 1
    '                CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Enabled = True
    '                CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text = 0
    '                CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text = 0
    '                CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text = 0
    '            Next
    '        End If
    '    Else
    '        For i = 0 To gridDescontos.Rows.Count - 1
    '            CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text = 0
    '            CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text = 0
    '            CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text = 0
    '        Next
    '        txtPesoBruto.Text = txtPrimeiraPesagem.Text
    '        txtLiquido.Text = txtPrimeiraPesagem.Text
    '        MsgBox(Me.Page, "Segunda pesagem deve ser maior que zero!")
    '    End If
    'End Sub

    Protected Sub btnEncerrar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnEncerrar.Click
        If Funcoes.VerificaPermissao("PesagemDaOrdemDeCarregamento", "GRAVAR") Then
            LimparCampos()
            btnIncluir.Enabled = False
            btnPesagem.Enabled = False
            btnEmpresa.Visible = False
            btnEncerrar.Enabled = False
            txtLaudo.Enabled = True
            txtLaudo.Focus()
        Else
            MsgBox(Me.Page, "Usuário sem Permissão para Pesagem")
        End If
    End Sub

    Protected Sub btnReimprimir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnReimprimir.Click
        If Funcoes.VerificaPermissao("PesagemDaOrdemDeCarregamento", "RELATORIO") Then
            ImprimirLaudo()
            'SessaoRecuperaLaudo()
            ''AndAlso objPesagem.Analises IsNot Nothing
            'If objLaudo.SegundaPesagem > 0 Then
            '    Imprimir("2", txtNumeroDeCopias.Text, "A", objLaudo)
            'Else
            '    Imprimir("1", txtNumeroDeCopias.Text, "A", objLaudo)
            'End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para reimpressão da pesagem")
        End If
    End Sub

    Protected Sub btnLerPeso_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnLerPeso.Click
        ChamarModalParaConfirmarPeso(LerPesoWS())
    End Sub

    Protected Sub btnAutorizacao_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAutorizacao.Click
        SessaoRecuperaLaudo()

        Dim peso As String = String.Empty
        Dim erroMsg As String = String.Empty
        Dim i As Integer = 0
        Dim dif As Boolean = False

        If Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) AndAlso
            Convert.ToInt32(txtSegundaPesagem.Text) > 0 Then

            ''Verifica novamente a segunda pesagem
            If hdnEletronico.Value Then
                While i < 5
                    System.Threading.Thread.Sleep(1000)

                    peso = LerPesoWS()

                    If IsNumeric(peso) Then
                        If Math.Abs(CInt(peso) - CInt(txtSegundaPesagem.Text)) > 50 Then
                            dif = True
                            erroMsg = "Ocilação da Balança ultrapassou 50 KGS, refazendo leitura."
                        End If
                    Else
                        dif = True
                        erroMsg = "Não foi possível comparar peso Capturado, refazendo leitura."
                    End If

                    i += 1
                End While

                If dif Then
                    MsgBox(Me.Page, erroMsg)
                    txtSegundaPesagem.Text = ""
                    ChamarModalParaConfirmarPeso(LerPesoWS())
                    Exit Sub
                End If
            End If

            'If gridDescontos.Rows.Count > 0 Then
            '    Dim ddlOp As DropDownList
            '    For i = 0 To gridDescontos.Rows.Count - 1
            '        CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Enabled = True
            '        ddlOp = CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList)
            '        If ddlOp IsNot Nothing Then CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList).Enabled = True
            '    Next
            '    CType(gridDescontos.Rows(0).FindControl("txtPercentual"), TextBox).Focus()

            '    If lblEntradaSaida.Text = "ENTRADA" Then btnCalcular.Enabled = True
            '    If lblEntradaSaida.Text = "SAIDA" Then
            '        If hdnAutorizacaoDeRetirada.Value > 0 Then
            '            CarregarAutorizacaoDeRetirada()
            '        Else
            '            btnCalcular.Enabled = True
            '        End If
            '    End If
            'Else
            '    If objPesagem.Produto.Agrupar = "S" Then
            '        btnIncluir.Enabled = True
            '    ElseIf lblEntradaSaida.Text = "ENTRADA" Then
            '        btnIncluir.Enabled = True
            '    ElseIf lblEntradaSaida.Text = "SAIDA" Then
            '        CarregarAutorizacaoDeRetirada()
            '    End If
            'End If
            btnIncluir.Enabled = True
        Else
            ''Verifica novamente a primeira pesagem
            If hdnEletronico.Value Then
                While i < 5
                    System.Threading.Thread.Sleep(1000)

                    peso = LerPesoWS()

                    If IsNumeric(peso) Then
                        If Math.Abs(CInt(peso) - CInt(txtPrimeiraPesagem.Text)) > 50 Then
                            dif = True
                            erroMsg = "Ocilação da Balança ultrapassou 50 KGS, refazendo leitura."
                        End If
                    Else
                        dif = True
                        erroMsg = "Não foi possível comparar peso Capturado, refazendo leitura."
                    End If

                    i += 1
                End While

                If dif Then
                    MsgBox(Me.Page, erroMsg)
                    txtPrimeiraPesagem.Text = ""
                    '    ChamarModalParaConfirmarPeso(LerPesoWS())
                    'Else
                    '    Dim ddlOp As DropDownList
                    '    For i = 0 To gridDescontos.Rows.Count - 1
                    '        ddlOp = CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList)
                    '        If ddlOp IsNot Nothing Then CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList).Enabled = True
                    '    Next
                End If
            End If
        End If
    End Sub

    'Protected Sub btnRateio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRateio.Click
    '    If pnlRateio.Visible Then
    '        pnlRateio.Visible = False
    '    Else
    '        pnlRateio.Visible = True
    '    End If
    'End Sub

    Protected Sub btnCancelar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancelar.Click
        If Funcoes.VerificaPermissao("PesagemDaOrdemDeCarregamento", "EXCLUIR") Then
            SessaoRecuperaLaudo()

            'If objPesagem.TemNota Then
            '    MsgBox(Me.Page, "Pesagem com Nota Fiscal não pode ser cancelada.")
            If (String.IsNullOrWhiteSpace(txtObservacao.Text) OrElse txtObservacao.Text.Trim.Length < 15) Then
                MsgBox(Me.Page, "O campo Observação é obrigatório e deve conter mais de 15 caracteres. Informe o motivo do cancelamento.")
            Else
                Cancelar()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para cancelar a pesagem")
        End If
    End Sub

    Protected Sub btnTesteBalanca_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnTesteBalanca.Click
        Dim peso As String = LerPesoWS()
        If String.IsNullOrWhiteSpace(peso) Then
            MsgBox(Me.Page, "A variável peso não foi informada!")
            Exit Sub
        End If
        MsgBox(Me.Page, "Leitura da balança: " & peso)
    End Sub

    'Protected Sub imgExtratoPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgExtratoPedido.Click
    '    Extrato.Emitir(Me.Page, FinanceiroNovo, txtCnpjDaEmpresa.Text.Replace(".", "").Replace("/", "").Replace("-", ""), hdnCodigoEndEmpresa.Value, "T",
    '                     txtPedido.Text)
    'End Sub

#End Region

#Region "DropDownList"

    'Protected Sub ddlOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlOperacao.SelectedIndexChanged
    '    CarregarSubOperacao()
    'End Sub

    'Protected Sub ddlSubOperacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlSubOperacao.SelectedIndexChanged
    '    If ddlSubOperacao.SelectedIndex > 0 Then
    '        Dim sOp() As String = ddlSubOperacao.SelectedValue.ToString.Split("-")
    '        Dim suboperacao As New [Lib].Negocio.SubOperacao(sOp(0), sOp(1))
    '        If sOp(1) > 0 Then
    '            objPesagem = New Pesagem()
    '            objPesagem.CodigoEmpresa = txtCnpjDaEmpresa.Text.Replace(".", "").Replace("-", "").Replace("/", "")
    '            objPesagem.EnderecoEmpresa = hdnCodigoEndEmpresa.Value
    '            objPesagem.CodigoPedido = txtPedido.Text
    '            objPesagem.Consolidado = chkConsolidarCliente.Checked
    '            Dim dsSaldoPedido As DataSet = objPesagem.SaldoDePedidos(objPesagem, False)
    '            If suboperacao.Devolucao Then
    '                hdnSaldoPedido.Value = dsSaldoPedido.Tables(0).Rows(0).Item("SaldoDevolucao")
    '            Else
    '                hdnSaldoPedido.Value = dsSaldoPedido.Tables(0).Rows(0).Item("Saldo")
    '            End If
    '            txtSaldoPedido.Text = hdnSaldoPedido.Value
    '        End If
    '        If (Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) AndAlso Convert.ToInt32(txtSegundaPesagem.Text) > 0) Then
    '            If Convert.ToInt32(txtPrimeiraPesagem.Text) > Convert.ToInt32(txtSegundaPesagem.Text) AndAlso suboperacao.EntradaSaida = eEntradaSaida.Saida Then
    '                MsgBox(Me.Page, "Operação de saída não pode ser usada para Entrada.")
    '            ElseIf Convert.ToInt32(txtSegundaPesagem.Text) > Convert.ToInt32(txtPrimeiraPesagem.Text) AndAlso suboperacao.EntradaSaida = eEntradaSaida.Entrada Then
    '                MsgBox(Me.Page, "Operação de entrada não pode ser usada para Saída.")
    '            ElseIf suboperacao.EntradaSaida = eEntradaSaida.Entrada Then
    '                lblEntradaSaida.Text = "ENTRADA"
    '            ElseIf suboperacao.EntradaSaida = eEntradaSaida.Saida Then
    '                lblEntradaSaida.Text = "SAIDA"
    '                CarregarAutorizacaoDeRetirada()
    '            End If
    '        Else
    '            If suboperacao.EntradaSaida = eEntradaSaida.Entrada Then
    '                lblEntradaSaida.Text = "ENTRADA"
    '            Else
    '                lblEntradaSaida.Text = "SAIDA"
    '            End If
    '        End If
    '    Else
    '        lblEntradaSaida.Text = "SAIDA"
    '    End If
    'End Sub

#End Region

#Region "CheckBox"

    Protected Sub radArquivo_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles radArquivo.CheckedChanged
        ddlImpressora.ClearSelection()
        ddlImpressora.Enabled = False
    End Sub

    Protected Sub radImpressora_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles radImpressora.CheckedChanged
        ddlImpressora.Enabled = True
    End Sub

#End Region

#Region "TextBox"

    Protected Sub txtLaudo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLaudo.TextChanged
        If String.IsNullOrEmpty(txtCnpjDaEmpresa.Text) Then
            MsgBox(Me.Page, "Número do Laudo não foi informado.")
        Else
            CarregarLaudo(txtLaudo.Text, txtCnpjDaEmpresa.Text.Replace(".", "").Replace("/", "").Replace("-", ""), hdnCodigoEndEmpresa.Value)
            btnPesagem.Enabled = False
            btnEncerrar.Enabled = False
            btnIncluir.Enabled = False
            btnCancelar.Enabled = True
            btnReimprimir.Enabled = True
        End If
    End Sub

#End Region

#Region "GridView"
    Protected Sub Gridlaudo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gdvLaudo.SelectedIndexChanged
        Dim index = gdvLaudo.SelectedIndex
        CarregarLaudo(gdvLaudo.Rows(index).Cells(1).Text, gdvLaudo.Rows(index).Cells(2).Text, gdvLaudo.Rows(index).Cells(3).Text)
        btnPesagem.Enabled = False
        btnEncerrar.Enabled = False
        btnIncluir.Enabled = False
        btnCancelar.Enabled = True
        btnReimprimir.Enabled = True
    End Sub
#End Region

#End Region

#Region "Session"
    Private objLaudo As [Lib].Negocio.LaudoDeCarregamento


    Private Sub SessaoSalvaLaudo()
        Session("ssLaudo" & HID.Value) = objLaudo
    End Sub

    Private Sub SessaoRecuperaLaudo()
        objLaudo = CType(Session("ssLaudo" & HID.Value), [Lib].Negocio.LaudoDeCarregamento)
    End Sub
#End Region

#Region "Métodos"

#Region "Carregar"

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        SessaoRecuperaLaudo()
        If Not Session("objEmpresa" & HID.Value) Is Nothing Then
            Session("objEmpresaLXR") = CType(Session("objEmpresa" & HID.Value), [Lib].Negocio.Cliente)
            Session.Remove("objEmpresa" & HID.Value)
            SetarEmpresa()
        ElseIf Not objLaudo Is Nothing Then
            Select Case hdnControlePopup.Value.Replace(",", "")
                'Case "Cliente"
                '    objPesagem.Cliente = CType(obj, [Lib].Negocio.Cliente)
                '    With objPesagem.Cliente
                '        objPesagem.CodigoCliente = .Codigo
                '        objPesagem.EnderecoCliente = .CodigoEndereco
                '        'hdnCodigoEndCliente.Value = .CodigoEndereco
                '        'txtNomeDoCliente.Text = .Nome
                '        'txtCnpjDoCliente.Text = .CodigoFormatado
                '        'txtEnderecoDoCliente.Text = .Endereco
                '        'txtComplementoDoCliente.Text = .Complemento
                '        'txtBairroDoCliente.Text = .Bairro
                '        'txtMunicipioDoCliente.Text = .Cidade.TrimEnd
                '        'txtCepDoCliente.Text = .CEP
                '        'If (Not String.IsNullOrEmpty7(.Telefone)) Then
                '        '    objPesagem.Cliente.Telefone = objPesagem.Cliente.Telefone.Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "")
                '        '    txtTelefoneDoCliente.Text = String.Format("{0:(##) ####-####}", Long.Parse(objPesagem.Cliente.Telefone.Trim))
                '        'End If
                '        'txtEstadoDoCliente.Text = .Estado.Codigo
                '        'If (Not String.IsNullOrEmpty(.InscricaoEstadual)) Then
                '        '    txtInscricaoDoCliente.Text = .InscricaoEstadual
                '        'End If
                '        'txtDepositante.Text = .CodigoFormatado & "-" & .Nome
                '        'hdnCodigoEndDepositante.Value = .CodigoEndereco
                '        objPesagem.CodigoDepositante = .Codigo
                '        objPesagem.EnderecoDepositante = .CodigoEndereco
                '        Limpar()
                '        CarregarPopupPedido(True)
                '    End With
                Case "Transportador"
                    objLaudo.Transportador = CType(obj, [Lib].Negocio.Cliente)
                    With objLaudo.Transportador
                        objLaudo.CodigoTransportador = .Codigo
                        objLaudo.EnderecoTransportador = .CodigoEndereco
                        txtTransportador.Text = .CodigoFormatado & "-" & .Nome
                        hdnCodigoEndTransportador.Value = .CodigoEndereco
                    End With
                    btnPlaca.Visible = True
                    btnPlaca_Click(btnPlaca, Nothing)
                'Case "Depositante"
                '    objPesagem.Depositante = CType(obj, [Lib].Negocio.Cliente)
                '    With objPesagem.Depositante
                '        objPesagem.CodigoDepositante = .Codigo
                '        objPesagem.EnderecoDepositante = .CodigoEndereco
                '        'txtDepositante.Text = .CodigoFormatado & "-" & .Nome
                '        'hdnCodigoEndDepositante.Value = .CodigoEndereco
                '    End With
                Case "Placa"
                    objLaudo.Placa = CType(obj, [Lib].Negocio.Placa)
                    If Not objLaudo.Placa Is Nothing And Not String.IsNullOrEmpty(objLaudo.Placa.Placa01) Then
                        With objLaudo.Placa
                            objLaudo.CodigoPlaca = .Placa01
                            objLaudo.CodigoViaDeTransporte = .ViaDeTransporte
                            objLaudo.CodigoMotorista = .Motorista.Codigo
                            txtPlaca.Text = .Placa01
                            txtMotorista.Text = .Motorista.Nome & "-" & .Motorista.CodigoFormatado
                            hdnCodigoEndMotorista.Value = .Motorista.CodigoEndereco
                            hdnViaDeTransporte.Value = .ViaDeTransporte
                        End With
                    End If

                    If (Not txtPrimeiraPesagem.Enabled) Then
                        If hdnEletronico.Value Then
                            ChamarModalParaConfirmarPeso(LerPesoWS())
                        Else
                            If Not String.IsNullOrWhiteSpace(txtPrimeiraPesagem.Text) AndAlso CInt(txtPrimeiraPesagem.Text) > 0 Then
                                txtSegundaPesagem.Enabled = True
                                txtSegundaPesagem.Text = ""
                                Popup.SetFocus(Me.Page, txtSegundaPesagem.ClientID, True)
                            Else
                                txtPrimeiraPesagem.Enabled = True
                                txtPrimeiraPesagem.BorderStyle = BorderStyle.Solid
                                txtPrimeiraPesagem.BorderWidth = Unit.Pixel(1)
                                txtPrimeiraPesagem.BorderColor = Drawing.Color.Black
                                txtSegundaPesagem.Enabled = True
                                txtSegundaPesagem.BorderStyle = BorderStyle.Solid
                                txtSegundaPesagem.BorderWidth = Unit.Pixel(1)
                                txtSegundaPesagem.BorderColor = Drawing.Color.Black
                                txtPrimeiraPesagem.Text = ""
                                Popup.SetFocus(Me.Page, txtPrimeiraPesagem.ClientID, True)
                            End If
                        End If
                    End If

                    'Case "Deposito"
                    '    objPesagem.Deposito = CType(obj, [Lib].Negocio.Cliente)
                    '    With objPesagem.Deposito
                    '        objPesagem.CodigoDeposito = .Codigo
                    '        objPesagem.EnderecoDeposito = .CodigoEndereco
                    '        'txtDeposito.Text = .CodigoFormatado & "-" & .Nome
                    '        'hdnCodigoEndDeposito.Value = .CodigoEndereco
                    '    End With
                    '    Dim Empresa = [txtCnpjDaEmpresa].Text.Replace(".", "").Replace("-", "").Replace("/", "")
                    '    Dim Endereco = hdnCodigoEndEmpresa.Value
                    '    'Dim Pedido = txtPedido.Text
                    '    'objPesagem.Pedido = New [Lib].Negocio.Pedido(Empresa, Endereco, Pedido)
                    '    hdnControlePopup.Value = "Transportador"
                    '    ucConsultaClientesDireto.Limpar()
                    '    ucConsultaClientesDireto.SetarTituloDIV("Consulta de Transportador")
                    '    ucConsultaClientesDireto.CarregarTransportadoresPorPedido(objPesagem.Pedido.Transportadores)
                    '    Popup.ConsultaDeClientesDireto(Me.Page, "objPesagem" & HID.Value, "", True, 600)
                    '    Popup.CenterDialog(Me.Page, "divConsultaClienteDireto")
                    'Case "Pedido"
                    '    objPesagem.Pedido = CType(Session("objPesagem" & HID.Value), [Lib].Negocio.Pedido)
                    '    objPesagem.CodigoPedido = objPesagem.Pedido.Codigo
                    '    'imgExtratoPedido.Visible = True
                    '    If objPesagem.CodigoPedido = 0 Then
                    '        MsgBox(Me.Page, "Não existem pedidos para o cliente selecionado!")
                    '    Else
                    '        objPesagem.CodigoEmpresa = txtCnpjDaEmpresa.Text.Replace(".", "").Replace("-", "").Replace("/", "")
                    '        objPesagem.EnderecoEmpresa = hdnCodigoEndEmpresa.Value

                    '        'ADICIONANDO CONTROLE POR AUTORIZAÇÃO DE CARREGAMENTO
                    '        Dim lstAutCarregamento As ListAutCarregamento = Nothing
                    '        If EmbarqueNovo Then
                    '            lstAutCarregamento = New ListAutCarregamento(objPesagem.CodigoPedido, objPesagem.CodigoPlaca)
                    '            If lstAutCarregamento IsNot Nothing AndAlso lstAutCarregamento.Count > 0 Then
                    '                'PESQUISA OS TRANSPORTADORES DA AUTORIZAÇÃO DE CARREGAMENTO
                    '                For Each objAutCarregamento As [Lib].Negocio.AutCarregamento In lstAutCarregamento
                    '                    Dim objPedidoXTransportador As New [Lib].Negocio.PedidoXTransportador(objPesagem.Pedido)
                    '                    objPedidoXTransportador.Codigo = objAutCarregamento.CodigoTransportador
                    '                    objPedidoXTransportador.CodigoEndereco = objAutCarregamento.EndTransportador
                    '                    objPesagem.Pedido.Transportadores.Add(objPedidoXTransportador)
                    '                Next
                    '            End If
                    '        Else
                    '            'PESQUISA OS TRANSPORTADORES DO PEDIDO
                    '            objPesagem.Pedido.Transportadores = New [Lib].Negocio.ListPedidoXTransportador(objPesagem.Pedido)
                    '        End If

                    '        'PESQUISA OS ITENS DO PEDIDO
                    '        'objPesagem.Pedido.Itens = New [Lib].Negocio.ListPedidoXItem(objPesagem.Pedido)

                    '        'PESQUISA OS DEPÓSITOS DO PEDIDO
                    '        objPesagem.Pedido.Depositos = New [Lib].Negocio.ListPedidoxDeposito(objPesagem.Pedido)

                    '        With objPesagem.Pedido
                    '            objPesagem.CodigoPedido = .Codigo
                    '            objPesagem.CodigoProduto = .Itens(0).CodigoProduto
                    '            objPesagem.CodigoOperacao = .CodigoOperacao
                    '            objPesagem.CodigoSubOperacao = .CodigoSubOperacao

                    '            ''*** LIBERADO TEMPORARIAMENTE PARA AGRÍCOLA ATÉ AJUSTARMOS NO DEPOSITO DO PEDIDO - FURLAN - 02/07/2014
                    '            'If String.IsNullOrWhiteSpace(txtPrimeiraPesagem.Text) OrElse Convert.ToInt32(txtPrimeiraPesagem.Text) = 0 Then
                    '            '    If Left(objPesagem.CodigoEmpresa, 8) = "04854422" Then ddlTabelaDeClassificacao.Enabled = True
                    '            'End If
                    '            ''*****************************************************************************************************

                    '            objPesagem.CodigoTabelaDeClassificacao = .Itens(0).Classificacao.Codigo
                    '            'txtPedido.Text = .Codigo
                    '            'ddlOperacao.SelectedValue = .CodigoOperacao
                    '            'CarregarSubOperacao()
                    '            'ddlSubOperacao.SelectedValue = .CodigoOperacao & "-" & .CodigoSubOperacao
                    '            'Dim intIndice As Integer = ddlSubOperacao.Items.IndexOf(ddlSubOperacao.Items.FindByValue(.CodigoOperacao & "-" & .CodigoSubOperacao))
                    '            'If Not intIndice = -1 Then
                    '            '    ddlSubOperacao.SelectedIndex = intIndice
                    '            'End If
                    '            If objPesagem.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                    '                lblEntradaSaida.Text = "ENTRADA"
                    '            Else
                    '                lblEntradaSaida.Text = "SAIDA"
                    '            End If
                    '            'txtProduto.Text = .Itens(0).CodigoProduto & "-" & .Itens(0).Descricao
                    '            'txtSafra.Text = .CodigoSafra
                    '            'ddlTabelaDeClassificacao.SelectedValue = .Itens(0).Classificacao.Codigo
                    '            'objPesagem.Consolidado = chkConsolidarCliente.Checked
                    '            'Dim ds As DataSet = objPesagem.SaldoDePedidos(objPesagem, False)
                    '            'hdnSaldoPedido.Value = ds.Tables(0).Rows(0).Item("Saldo")
                    '            'txtSaldoPedido.Text = hdnSaldoPedido.Value
                    '            'ddlOperacao.Enabled = False
                    '        End With
                    '        Dim deposito As [Lib].Negocio.PedidoXDeposito = objPesagem.Pedido.Depositos.FirstOrDefault(Function(D) D.Tipo = "DE")

                    '        If deposito IsNot Nothing Then
                    '            'txtDeposito.Text = deposito.Deposito.CodigoFormatado & "-" & deposito.Deposito.Nome
                    '            'hdnCodigoEndDeposito.Value = deposito.Deposito.CodigoEndereco
                    '            objPesagem.CodigoDeposito = deposito.Deposito.Codigo
                    '            objPesagem.EnderecoDeposito = deposito.Deposito.CodigoEndereco
                    '            'Transportador
                    '            If (objPesagem.Pedido.Transportadores.Count.Equals(1)) Then
                    '                txtTransportador.Text = objPesagem.Pedido.Transportadores(0).Transportador.CodigoFormatado & "-" & objPesagem.Pedido.Transportadores(0).Transportador.Nome
                    '                hdnCodigoEndTransportador.Value = objPesagem.Pedido.Transportadores(0).Transportador.CodigoEndereco
                    '                objPesagem.CodigoTransportador = objPesagem.Pedido.Transportadores(0).Transportador.Codigo
                    '                objPesagem.EnderecoTransportador = objPesagem.Pedido.Transportadores(0).Transportador.CodigoEndereco
                    '                If EmbarqueNovo Then
                    '                    If lstAutCarregamento IsNot Nothing AndAlso lstAutCarregamento.Count > 0 AndAlso Not String.IsNullOrWhiteSpace(lstAutCarregamento(0).Placa) Then
                    '                        objPesagem.Placa = New [Lib].Negocio.Placa(lstAutCarregamento(0).Placa)
                    '                        With objPesagem.Placa
                    '                            objPesagem.CodigoPlaca = .Placa01
                    '                            objPesagem.CodigoViaDeTransporte = .ViaDeTransporte
                    '                            objPesagem.CodigoMotorista = .Motorista.Codigo
                    '                            txtPlaca.Text = .Placa01
                    '                            txtMotorista.Text = .Motorista.Nome & "-" & .Motorista.CodigoFormatado
                    '                            hdnCodigoEndMotorista.Value = .Motorista.CodigoEndereco
                    '                            hdnViaDeTransporte.Value = .ViaDeTransporte
                    '                            If (Not txtPrimeiraPesagem.Enabled) Then
                    '                                ChamarModalParaConfirmarPeso(LerPesoWS())
                    '                            End If
                    '                        End With
                    '                    End If
                    '                Else
                    '                    btnPlaca_Click(btnPlaca, Nothing)
                    '                End If
                    '            ElseIf (objPesagem.Pedido.Transportadores.Count.Equals(0)) Then
                    '                Session("ssCampo" & HID.Value) = "Transportador"

                    '                hdnControlePopup.Value = "Transportador"
                    '                ucConsultaClientes.Limpar()
                    '                ucConsultaClientes.SetarTipoCliente("7")
                    '                ucConsultaClientes.SetarTituloDIV("Consulta de Transportador")
                    '                Popup.ConsultaDeClientes(Me.Page, "objPesagem" & HID.Value)
                    '                Popup.CenterDialog(Me.Page, "divConsultaCliente")
                    '            Else
                    '                hdnControlePopup.Value = "Transportador"
                    '                ucConsultaClientesDireto.Limpar()
                    '                ucConsultaClientesDireto.SetarTituloDIV("Consulta de Transportador")
                    '                ucConsultaClientesDireto.CarregarTransportadoresPorPedido(objPesagem.Pedido.Transportadores)
                    '                Popup.ConsultaDeClientesDireto(Me.Page, "objPesagem" & HID.Value)
                    '                Popup.CenterDialog(Me.Page, "divConsultaClienteDireto")
                    '            End If
                    '            CarregarAnalises()
                    '        Else
                    '            hdnControlePopup.Value = "Deposito"
                    '            ucConsultaClientesDireto.Limpar()
                    '            ucConsultaClientesDireto.SetarTituloDIV("Consulta de Depósitos")
                    '            ucConsultaClientesDireto.CarregarDepositosPorPedido(objPesagem.Pedido.Depositos)
                    '            Popup.ConsultaDeClientesDireto(Me.Page, "objPesagem" & HID.Value)
                    '            Popup.CenterDialog(Me.Page, "divConsultaClienteDireto")
                    '        End If
                    '    End If
            End Select
        End If
        SessaoSalvaLaudo()
    End Sub

    Public Sub LiberaTransportador()
        If (Not txtPrimeiraPesagem.Enabled) Then
            If hdnEletronico.Value Then
                ChamarModalParaConfirmarPeso(LerPesoWS())
            Else
                txtPrimeiraPesagem.Enabled = True
                txtPrimeiraPesagem.BorderStyle = BorderStyle.Solid
                txtPrimeiraPesagem.BorderWidth = Unit.Pixel(1)
                txtPrimeiraPesagem.BorderColor = Drawing.Color.Black
                txtSegundaPesagem.Enabled = True
                txtSegundaPesagem.BorderStyle = BorderStyle.Solid
                txtSegundaPesagem.BorderWidth = Unit.Pixel(1)
                txtSegundaPesagem.BorderColor = Drawing.Color.Black
                txtPrimeiraPesagem.Text = ""
            End If
            Popup.SetFocus(Me.Page, txtPrimeiraPesagem.ClientID, True)
        Else
            Dim teste As String = "vair vir aqui - testando, avise o suporte"
            MsgBox(Me.Page, teste)
        End If
    End Sub

    Private Sub CarregarBalanca()
        Dim Ip As String = Context.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        If Ip = String.Empty Then
            Ip = Context.Request.ServerVariables("REMOTE_ADDR")
            If Ip = "::1" Then Ip = "127.0.0.1"
        End If

        txtNumeroDeCopias.Text = "1"
        lblAcessoUsuario.Text = Ip

        Dim objListaBalancaEletronica As New [Lib].Negocio.ListBalancaEletronica(Ip, HttpContext.Current.Session("ssNomeUsuario"))

        If objListaBalancaEletronica.Count = 0 Then
            lblAcesso.Text = Ip
            chkLiberado.Checked = False
            MsgBox(Me.Page, "Usuário sem permissão para executar Pesagem, apenas Consulta")
        Else
            If (objListaBalancaEletronica.Count.Equals(1)) Then
                lblModulo.Text = objListaBalancaEletronica.Item(0).BalancaTipo
                lblPorta.Text = objListaBalancaEletronica.Item(0).PortName

                lblAcesso.Text = objListaBalancaEletronica.Item(0).BalancaIp

                hdnBaudRateBalanca.Value = objListaBalancaEletronica.Item(0).BaudRate
                hdnDataBits.Value = objListaBalancaEletronica.Item(0).DataBits
                hdnParity.Value = objListaBalancaEletronica.Item(0).Parity
                hdnStopBits.Value = objListaBalancaEletronica.Item(0).StopBits

                chkLiberado.Checked = True
                hdnEletronico.Value = True
                txtNumeroDeCopias.Text = objListaBalancaEletronica.Item(0).NumeroDeCopias

                If (objListaBalancaEletronica.Item(0).Eletronica.Equals("N")) Then
                    lblPesagemMecanica.Text = "Pesagem Manual"
                    btnTesteBalanca.Visible = False
                    hdnEletronico.Value = False
                End If
            Else
                lblAcesso.Text = Ip
                chkLiberado.Checked = False
                MsgBox(Me.Page, "Usuário com mais de uma balança para o mesmo IP, entre em contato com o Suporte")
            End If
        End If
    End Sub

    'Protected Sub CarregarSubOperacao()
    '    SessaoRecuperaLaudo()

    '    Dim Where As String
    '    Where = "     So.Operacao_id =" & ddlOperacao.SelectedValue & vbCrLf &
    '            " AND So.Laudo = 'S'" & vbCrLf &
    '            " AND So.Situacao = 1" & vbCrLf &
    '            " AND So.Liminar = " & IIf(objPesagem.Pedido.SubOperacao.Liminar, 1, 0) & vbCrLf &
    '            " AND So.Consignacao =  " & IIf(objPesagem.Pedido.SubOperacao.Consignacao, 1, 0) & vbCrLf &
    '            " AND So.PrecoFixo = '" & IIf(objPesagem.Pedido.SubOperacao.PrecoFixo, "S", "N") & "'" & vbCrLf

    '    ddl.Carregar(ddlSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, Where)
    'End Sub

    'Protected Sub CarregarTabelaClassificacao()
    '    ddl.Carregar(ddlTabelaDeClassificacao, CarregarDDL.Tabela.TabelaDeClassificacoes, "")
    'End Sub

    'Protected Sub ddlTabelaDeClassificacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTabelaDeClassificacao.SelectedIndexChanged
    '    CarregarAnalises()

    '    If (Not txtPrimeiraPesagem.Enabled) Then
    '        If hdnEletronico.Value Then
    '            ChamarModalParaConfirmarPeso(LerPesoWS())
    '        Else
    '            If Not String.IsNullOrWhiteSpace(txtPrimeiraPesagem.Text) AndAlso CInt(txtPrimeiraPesagem.Text) > 0 Then
    '                txtSegundaPesagem.Enabled = True
    '                txtSegundaPesagem.Text = ""
    '                Popup.SetFocus(Me.Page, txtSegundaPesagem.ClientID, True)
    '            Else
    '                txtPrimeiraPesagem.Enabled = True
    '                txtPrimeiraPesagem.BorderStyle = BorderStyle.Solid
    '                txtPrimeiraPesagem.BorderWidth = Unit.Pixel(1)
    '                txtPrimeiraPesagem.BorderColor = Drawing.Color.Black
    '                txtSegundaPesagem.Enabled = True
    '                txtSegundaPesagem.BorderStyle = BorderStyle.Solid
    '                txtSegundaPesagem.BorderWidth = Unit.Pixel(1)
    '                txtSegundaPesagem.BorderColor = Drawing.Color.Black
    '                txtPrimeiraPesagem.Text = ""
    '                Popup.SetFocus(Me.Page, txtPrimeiraPesagem.ClientID, True)
    '            End If
    '        End If
    '    End If
    'End Sub

    'Protected Sub CarregarAnalises()
    '    SessaoRecuperaLaudo()
    '    objPesagem.CodigoEmpresa = txtCnpjDaEmpresa.Text.Replace(".", "").Replace("/", "").Replace("-", "")
    '    objPesagem.EnderecoEmpresa = hdnCodigoEndEmpresa.Value
    '    objPesagem.Sequencia = 0
    '    'objPesagem.CodigoTabelaDeClassificacao = ddlTabelaDeClassificacao.SelectedValue
    '    'objPesagem.CodigoProduto = txtProduto.Text.Split("-").ToArray(0).TrimEnd
    '    objPesagem.Analises = New [Lib].Negocio.ListPesagemXAnalises(objPesagem)

    '    btnCalcular.Visible = objPesagem.Analises IsNot Nothing AndAlso objPesagem.Analises.Count > 0
    '    gridDescontos.DataSource = objPesagem.Analises
    '    gridDescontos.DataBind()

    '    If Not txtSegundaPesagem.Text.Equals("0") Then
    '        btnCalcular.Enabled = True
    '        For i = 0 To gridDescontos.Rows.Count - 1
    '            CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Enabled = False
    '        Next
    '    End If
    '    SessaoSalvaLaudo()
    'End Sub

    Protected Sub CarregarGridLaudo()
        objLaudo = New [Lib].Negocio.LaudoDeCarregamento
        objLaudo.CodigoEmpresa = txtCnpjDaEmpresa.Text.Replace(".", "").Replace("/", "").Replace("-", "")
        objLaudo.EnderecoEmpresa = hdnCodigoEndEmpresa.Value
        Dim listLaudo As New ListLaudoDeCarregamento(objLaudo, txtMovimento.Text, txtMovimento.Text)
        gdvLaudo.DataSource = listLaudo
        gdvLaudo.DataBind()

        For i = 0 To gdvLaudo.Rows.Count - 1
            If gdvLaudo.Rows(i).Cells(1).Text = listLaudo(i).Codigo AndAlso Not listLaudo(i).CodigoSituacao = 1 Then
                CType(gdvLaudo.Rows(i).FindControl("imgImprimir"), ImageButton).Visible = False
            Else
                CType(gdvLaudo.Rows(i).FindControl("imgImprimir"), ImageButton).Visible = True
            End If
        Next

        TabContainer1.ActiveTabIndex = 1
    End Sub

    'Public Sub CarregarPopupPedido(ByVal AnoCorrente As Boolean)
    '    Dim parameters As New Dictionary(Of String, Object)
    '    Session("ssCampo" & HID.Value) = "Laudo"
    '    parameters("CodigoEmpresa") = txtCnpjDaEmpresa.Text.Replace(".", "").Replace("-", "").Replace("/", "")
    '    parameters("CodigoEndEmpresa") = hdnCodigoEndEmpresa.Value
    '    'parameters("CodigoCliente") = txtCnpjDoCliente.Text.Replace(".", "").Replace("-", "").Replace("/", "")
    '    'parameters("CodigoEndCliente") = hdnCodigoEndCliente.Value

    '    Dim codigoEmpresa As String = Left(parameters("CodigoEmpresa").ToString(), 8)
    '    If {"49673784", "40938762"}.Contains(codigoEmpresa) Then
    '        parameters("AnoCorrente") = False
    '    Else
    '        parameters("AnoCorrente") = AnoCorrente
    '    End If

    '    'parameters("Consolidado") = IIf(chkConsolidarCliente.Checked, True, False)
    '    ucConsultaPedidos.BindGridView(parameters)
    '    Popup.ConsultaDePedidos(Me.Page, "objPesagem" & HID.Value)
    '    hdnControlePopup.Value = "Pedido"
    'End Sub

    'Public Sub CarregarAutorizacaoDeRetirada()
    '    'Dim objPrd As New [Lib].Negocio.Produto(txtProduto.Text.Split("-")(0))
    '    'If Not objPrd.Agrupar = "N" Then Exit Sub

    '    Dim parametros As New Dictionary(Of String, Object)
    '    parametros("emp") = txtCnpjDaEmpresa.Text.Replace(".", "").Replace("-", "").Replace("/", "")
    '    parametros("endemp") = hdnCodigoEndEmpresa.Value
    '    'parametros("cli") = txtCnpjDoCliente.Text.Replace(".", "").Replace("-", "").Replace("/", "")
    '    'parametros("endcli") = hdnCodigoEndCliente.Value
    '    'parametros("ped") = txtPedido.Text
    '    'Dim sOp() As String = ddlSubOperacao.SelectedValue.ToString.Split("-")
    '    'Dim suboperacao As New [Lib].Negocio.SubOperacao(sOp(0), sOp(1))
    '    'parametros("classe") = suboperacao.Classe
    '    parametros("romaneio") = True
    '    ucConsultaAutorizacaoDeRetirada.BindGridView(parametros)
    '    Popup.ConsultaDeAutorizacaoDeRetirada(Me.Page, "objPesagem" & HID.Value)
    'End Sub

    'Public Sub CarregarAutorizacao(Par As Hashtable)
    '    'Dim sOp() As String = ddlSubOperacao.SelectedValue.ToString.Split("-")
    '    'Dim suboperacao As New [Lib].Negocio.SubOperacao(sOp(0), sOp(1))
    '    'Dim Autorizacao As New [Lib].Negocio.AutorizacaoDeRetirada(txtCnpjDaEmpresa.Text.Replace(".", "").Replace("-", "").Replace("/", ""), hdnCodigoEndEmpresa.Value, Par("Pedido"), Par("Autorizacao")Classe)
    '    If (Autorizacao.SaldoFisico < Convert.ToDouble(txtLiquido.Text)) Then
    '        btnIncluir.Enabled = False
    '        MsgBox(Me.Page, "Pedido com saldo insuficiente para gerar Laudo!")
    '    Else
    '        btnIncluir.Enabled = True
    '        'hdnAutorizacaoDeRetirada.Value = Autorizacao.Autorizacao
    '    End If
    'End Sub

#End Region

    Protected Sub Salvar(ByVal IUD As String)
        'I - INSET, U - UPDATE, D - DELETE
        objLaudo = New [Lib].Negocio.LaudoDeCarregamento
        objLaudo.IUD = IUD

        'EMPRESA
        objLaudo.CodigoEmpresa = txtCnpjDaEmpresa.Text.Replace(".", "").Replace("/", "").Replace("-", "")
        objLaudo.EnderecoEmpresa = hdnCodigoEndEmpresa.Value

        If (IUD.ToString.Equals("I")) Then
            'Data DE ENTRADA/SAÍDA NA BALANÇA
            objLaudo.EntradaBalanca = DateTime.Now
            objLaudo.Sequencia = 0 'NA INCLUSÃO GRAVAR (0)
            objLaudo.UsuarioInclusao = Session("ssNomeUsuario").ToString
            objLaudo.DataInclusao = DateTime.Now
            objLaudo.UsuarioAlteracao = ""
            objLaudo.DataAlteracao = CType(Nothing, DateTime?)

            'If Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) AndAlso Convert.ToInt32(txtSegundaPesagem.Text) > 0 Then objLaudo.CriarRomaneio = True

            'objPesagem.Analises = New [Lib].Negocio.ListPesagemXAnalises(objPesagem)

            'For i = 0 To gridDescontos.Rows.Count - 1
            '    Dim objPesagemXAnalises As New PesagemXAnalises(objPesagem)
            '    objPesagemXAnalises.CodigoAnalise = gridDescontos.Rows(i).Cells(0).Text

            '    If Len(objPesagemXAnalises.Analise.Opcao) > 0 Then
            '        objPesagemXAnalises.Descricao = gridDescontos.Rows(i).Cells(1).Text
            '        objPesagemXAnalises.Percentual = CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList).SelectedValue
            '        objPesagem.Analises.Add(objPesagemXAnalises)
            '    Else
            '        Dim perc As TextBox = CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox)
            '        If IsNumeric(perc.Text) And CDec(perc.Text) > 0 Then
            '            objPesagemXAnalises.Descricao = gridDescontos.Rows(i).Cells(1).Text
            '            objPesagemXAnalises.Percentual = CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text
            '            objPesagemXAnalises.Indice = CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text
            '            objPesagemXAnalises.Desconto = CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text
            '            objPesagem.Analises.Add(objPesagemXAnalises)
            '        End If
            '    End If
            'Next


        ElseIf (IUD.Equals("U")) Then
            'Data DE ENTRADA/SAÍDA NA BALANÇA
            objLaudo.SaidaBalanca = DateTime.Now
            'objPesagem.UsuarioAlteracao = Session("ssNomeUsuario").ToString
            'objPesagem.DataAlteracao = DateTime.Now
            objLaudo.UsuarioAlteracao = Session("ssNomeUsuario").ToString
            objLaudo.DataAlteracao = DateTime.Now
            objLaudo.Codigo = txtLaudo.Text
            objLaudo.Sequencia = 0

            'If Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) AndAlso Convert.ToInt32(txtSegundaPesagem.Text) > 0 Then objLaudo.CriarRomaneio = True

            'objPesagem.Analises = New [Lib].Negocio.ListPesagemXAnalises(objPesagem)

            'For i = 0 To gridDescontos.Rows.Count - 1
            '    Dim objPesagemXAnalises As New PesagemXAnalises(objPesagem)
            '    objPesagemXAnalises.CodigoAnalise = gridDescontos.Rows(i).Cells(0).Text

            '    If Len(objPesagemXAnalises.Analise.Opcao) > 0 Then
            '        objPesagemXAnalises.Descricao = gridDescontos.Rows(i).Cells(1).Text
            '        objPesagemXAnalises.Percentual = CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList).SelectedValue
            '        objPesagem.Analises.Add(objPesagemXAnalises)
            '    Else
            '        Dim perc As TextBox = CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox)
            '        If IsNumeric(perc.Text) And CDec(perc.Text) > 0 Then
            '            objPesagemXAnalises.Descricao = gridDescontos.Rows(i).Cells(1).Text
            '            objPesagemXAnalises.Percentual = CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text
            '            objPesagemXAnalises.Indice = CType(gridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text
            '            objPesagemXAnalises.Desconto = CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text
            '            objPesagem.Analises.Add(objPesagemXAnalises)
            '        Else
            '            objPesagemXAnalises.IUD = "D"
            '            objPesagem.Analises.Add(objPesagemXAnalises)
            '        End If
            '    End If
            'Next

        End If
        'Situação do Pedido
        objLaudo.CodigoSituacao = ddlSituacao.SelectedValue
        'Pedido
        'objPesagem.CodigoPedido = txtPedido.Text
        'Autorização de Retirada
        'objPesagem.CodigoAutorizacao = hdnAutorizacaoDeRetirada.Value
        'Produto
        'objPesagem.CodigoProduto = txtProduto.Text.Split("-").ToArray(0).Trim
        'Cliente
        'objPesagem.CodigoCliente = txtCnpjDoCliente.Text.Replace(".", "").Replace("/", "").Replace("-", "")
        'objPesagem.EnderecoCliente = hdnCodigoEndCliente.Value
        'Depositante
        'objPesagem.CodigoDepositante = txtDepositante.Text.Split("-").ToArray(0).Trim.Replace(".", "").Replace("/", "").Replace("-", "") & txtDepositante.Text.Split("-").ToArray(1).Trim
        'objPesagem.EnderecoDepositante = hdnCodigoEndDepositante.Value

        'Deposito
        'objPesagem.CodigoDeposito = txtDeposito.Text.Split("-").ToArray(0).Trim.Replace(".", "").Replace("/", "").Replace("-", "") & txtDeposito.Text.Split("-").ToArray(1).Trim
        'objPesagem.EnderecoDeposito = hdnCodigoEndDeposito.Value

        'Tabela Classificação
        'objLaudo.CodigoTabelaDeClassificacao = ddlTabelaDeClassificacao.SelectedValue.ToString

        'Operação
        'objPesagem.CodigoOperacao = ddlOperacao.SelectedValue.ToString
        'SubOperação
        'objPesagem.CodigoSubOperacao = ddlSubOperacao.SelectedValue.Split("-").ToArray(1).Trim
        'Pesagem.
        objLaudo.PrimeiraPesagem = CType(txtPrimeiraPesagem.Text, Decimal)
        objLaudo.SegundaPesagem = CType(IIf(String.IsNullOrWhiteSpace(txtSegundaPesagem.Text), 0, txtSegundaPesagem.Text), Decimal)
        'objLaudo.BrutoBalanca = CType(txtPesoBruto.Text, Decimal)
        objLaudo.Liquido = CType(txtLiquido.Text, Decimal)

        If objLaudo.SegundaPesagem > 0 Then objLaudo.SaidaBalanca = DateTime.Now
        '
        'If Not String.IsNullOrEmpty(txtLaudo.Text) Then
        '    If (Convert.ToInt32(txtPrimeiraPesagem.Text) > Convert.ToInt32(txtSegundaPesagem.Text)) Then
        '        objPesagem.EntradaSaida = "E" 'Entrada
        '    Else
        '        objPesagem.EntradaSaida = "S" 'Saida
        '    End If
        '    objPesagem.SaidaBalanca = DateTime.Now 'Data da segunda pesagem.
        'Else
        '    If objPesagem.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
        '        objPesagem.EntradaSaida = "E" 'Entrada
        '    Else
        '        objPesagem.EntradaSaida = "S" 'Saida
        '    End If
        'End If

        'Transportador
        If String.IsNullOrWhiteSpace(txtTransportador.Text) Then
            objLaudo.CodigoTransportador = ""
        Else
            objLaudo.CodigoTransportador = txtTransportador.Text.Split("-").ToArray(0).Trim.Replace(".", "").Replace("/", "") & txtTransportador.Text.Split("-").ToArray(1).Trim
        End If
        objLaudo.EnderecoTransportador = hdnCodigoEndTransportador.Value
        'Motorista
        If String.IsNullOrWhiteSpace(txtMotorista.Text) Then
            objLaudo.CodigoMotorista = ""
        Else
            objLaudo.CodigoMotorista = txtMotorista.Text.Split("-").ToArray(1).Trim.Replace(".", "").Replace("/", "").Replace("-", "") & txtMotorista.Text.Split("-").ToArray(2).Trim
        End If
        objLaudo.EnderecoMotorista = hdnCodigoEndMotorista.Value
        'Placa
        objLaudo.CodigoPlaca = txtPlaca.Text
        'Via de Transporte
        objLaudo.CodigoViaDeTransporte = hdnViaDeTransporte.Value

        'Atualiza o movimento para a data de saída.
        objLaudo.Movimento = Today

        'Nota Produtor
        ''If (Not String.IsNullOrEmpty(txtNotaProdutor.Text)) Then
        ''    objPesagem.NumeroDaNota = txtNotaProdutor.Text
        ''    objPesagem.SerieDaNota = txtSerieNotaProdutor.Text
        ''    objPesagem.PesoFiscal = txtPesoNotaProdutor.Text
        'End If

        'Sempre (0)
        'objLaudo.PesagemDeTerceiros = 0
        'objLaudo.PesagemMecanica = 0

        objLaudo.UsuarioReimpressao = ""
        objLaudo.DataReimpressao = DateTime.Now
        objLaudo.UsuarioCancelamento = ""
        objLaudo.DataCancelamento = DateTime.Now
        objLaudo.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtObservacao.Text)


        ''Default(N) Não utilizado na pesagem
        'objLaudo.TemRomaneio = False

        'Sempre 0
        'objLaudo.RegistroMestre = 0

        If Not objLaudo.Salvar() Then
            MsgBox(Me.Page, "Erro ao Salvar Pesagem: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
            Exit Sub
        End If

        'Será dev. após validação
        'If (objPesagem.SegundaPesagem > 0) Then
        '    Imprimir("2", txtNumeroDeCopias.Text, "A", objLaudo)
        'Else
        '    Imprimir("1", txtNumeroDeCopias.Text, "A", objLaudo)
        'End If
        LimparCampos(IUD, objLaudo.Codigo)
    End Sub

    Protected Function ValidarCampos(ByVal IUD As String) As String
        Dim Erro As String = String.Empty

        'Empresa
        If (String.IsNullOrEmpty(txtCnpjDaEmpresa.Text) Or String.IsNullOrEmpty(hdnCodigoEndEmpresa.Value)) Then
            Erro += "O campo Empresa é obrigatório! \n"
        End If
        'Cliente
        'If (String.IsNullOrEmpty(txtCnpjDoCliente.Text) Or String.IsNullOrEmpty(hdnCodigoEndCliente.Value)) Then
        '    Erro += "O campo Cliente é obrigatório! \n"
        'End If
        ''Pedido
        'If (String.IsNullOrEmpty(txtPedido.Text)) Then
        '    Erro += "O campo Pedido é obrigatório! \n"
        'End If
        'If (String.IsNullOrEmpty(ddlOperacao.SelectedValue)) Then
        '    Erro += "O campo operação é obrigatório! \n"
        'End If
        'If (String.IsNullOrEmpty(ddlSubOperacao.SelectedValue)) Then
        '    Erro += "O campo Sub-Operação é obrigatório \n"
        'End If
        'If (String.IsNullOrEmpty(txtProduto.Text)) Then
        '    Erro += "O campo Produto é obrigatório! \n"
        'End If
        'If (String.IsNullOrEmpty(txtSafra.Text)) Then
        '    Erro += "O campo Safra é obrigatório! \n"
        'End If
        ''Depositante
        'If (String.IsNullOrEmpty(txtDepositante.Text) Or String.IsNullOrEmpty(hdnCodigoEndDepositante.Value)) Then
        '    Erro += "O campo Depositante é obrigatório! \n"
        'End If
        ''Depósito
        'If (String.IsNullOrEmpty(txtDeposito.Text) Or String.IsNullOrEmpty(hdnCodigoEndDeposito.Value)) Then
        '    Erro += "O campo Depósito é obrigatório! \n"
        'End If

        'Transportador
        If Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) AndAlso Convert.ToInt32(txtSegundaPesagem.Text) > 0 Then
            If (String.IsNullOrEmpty(txtTransportador.Text) Or String.IsNullOrEmpty(hdnCodigoEndMotorista.Value)) Then
                Erro += "O campo Transportador é obrigatório! \n"
            End If
            If (String.IsNullOrEmpty(txtPlaca.Text)) Then
                Erro += "O campo Placa é obrigatório! \n"
            End If
            If (String.IsNullOrEmpty(txtMotorista.Text) Or String.IsNullOrEmpty(hdnCodigoEndMotorista.Value)) Then
                Erro += "O campo Motorista é obrigatório! \n"
            End If
        End If

        ''Classificação
        'If (String.IsNullOrEmpty(ddlTabelaDeClassificacao.SelectedValue)) Then
        '    Erro += "O campo Tabela de Classificação é obrigatório! \n"
        'End If

        '1 Pesagem
        If (Not String.IsNullOrWhiteSpace(txtPrimeiraPesagem.Text) And IUD.Equals("I")) Then
            If (Convert.ToInt32(txtPrimeiraPesagem.Text) <= 0) Then
                Erro += "O campo Primeira Pesagem é obrigatório! \n"
            End If
        Else
            Erro += "O campo Primeira Pesagem é obrigatório! \n"
        End If

        '2 Pesagem
        If (Not String.IsNullOrWhiteSpace(txtLaudo.Text)) Then
            If (String.IsNullOrWhiteSpace(txtSegundaPesagem.Text)) Then
                Erro += "O campo Segunda Pesagem é obrigatório! \n"
            ElseIf (Convert.ToInt32(txtSegundaPesagem.Text) <= 0) Then
                Erro += "O campo Segunda Pesagem é obrigatório! \n"
            End If
        End If

        If (Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) AndAlso Convert.ToInt32(txtSegundaPesagem.Text) > 0) Then
            If Convert.ToInt32(txtSegundaPesagem.Text) > Convert.ToInt32(txtPrimeiraPesagem.Text) AndAlso lblEntradaSaida.Text = "SAIDA" Then
                Erro += "Operação de saída não pode ser usada para Entrada \n"
            ElseIf Convert.ToInt32(txtPrimeiraPesagem.Text) > Convert.ToInt32(txtSegundaPesagem.Text) AndAlso lblEntradaSaida.Text = "ENTRADA" Then
                Erro += "Operação de entrada não pode ser usada para Saída \n"
                'ElseIf lblEntradaSaida.Text = "SAIDA" AndAlso hdnAutorizacaoDeRetirada.Value = 0 Then
                '    Dim objPrd As New [Lib].Negocio.Produto(txtProduto.Text.Split("-")(0))
                '    'Verifica Autorização de Retirada se Produto Agrupar Não
                '    If objPrd.Agrupar = "N" Then Erro += "Autorização de Retirada não foi selecionada \n"
            End If
        End If

        'Líquido
        If Convert.ToInt32(txtLiquido.Text) <= 0 Then
            Erro += "O peso liquído: " & Convert.ToInt32(txtLiquido.Text) & " não pode ser menor ou igual a 0 (ZERO) \n"
        End If

        'If Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) AndAlso Convert.ToInt32(txtSegundaPesagem.Text) > 0 Then
        '    Dim pDesconto As Decimal = 0
        '    For i = 0 To gridDescontos.Rows.Count - 1
        '        If Not CInt(gridDescontos.Rows(i).Cells(0).Text) = 6 And Not CInt(gridDescontos.Rows(i).Cells(0).Text) = 12 Then
        '            pDesconto += CType(gridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text
        '        End If
        '    Next
        '    AndAlso pDesconto > 0 

        '    If Convert.ToInt32(txtPesoBruto.Text) = Convert.ToInt32(txtLiquido.Text) Then
        '        Erro += "Limpar o Laudo e Refazer, erro ao Calcular os Descontos \n"
        '    End If
        'End If

        SessaoRecuperaLaudo()

        'Saldo
        'If objPesagem.Produto.ControlarRomaneio Then
        '    If Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) AndAlso (Convert.ToInt32(txtSegundaPesagem.Text) > 0) Then
        '        Dim sOp() As String = ddlSubOperacao.SelectedValue.ToString.Split("-")
        '        Dim suboperacao As New [Lib].Negocio.SubOperacao(sOp(0), sOp(1))

        '        If Not suboperacao.Devolucao AndAlso suboperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso (objPesagem.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR OrElse objPesagem.Pedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS) Then
        '            'PODE PASSAR SEM CHECAR O SALDO
        '        Else
        '            If (Convert.ToInt32(txtLiquido.Text) > Convert.ToInt32(hdnSaldoPedido.Value)) Then
        '                Erro += "O peso liquído: " & Convert.ToInt32(txtLiquido.Text) & " não pode ser maior que o saldo do pedido: " & Convert.ToInt32(hdnSaldoPedido.Value)
        '            End If
        '        End If
        '    End If
        'End If

        'Impressora
        If (radImpressora.Checked) Then
            If (String.IsNullOrWhiteSpace(txtNumeroDeCopias.Text)) Then
                Erro += "O campo NÚMERO DE CÓPIAS é obrigatório \n"
            End If
        End If
        ''Nota Fiscal Produtor
        'If (Not String.IsNullOrWhiteSpace(txtNotaProdutor.Text) Or Not String.IsNullOrWhiteSpace(txtSerieNotaProdutor.Text) Or Not String.IsNullOrWhiteSpace(txtPesoNotaProdutor.Text)) Then
        '    If (String.IsNullOrEmpty(txtNotaProdutor.Text)) Then
        '        Erro += "O campo Nota Fiscal/Série é obrigatório! \n"
        '    End If
        '    If (String.IsNullOrEmpty(txtSerieNotaProdutor.Text)) Then
        '        Erro += "O campo Nota Fiscal/Série é obrigatório! \n"
        '    End If
        '    If (String.IsNullOrEmpty(txtSerieNotaProdutor.Text)) Then
        '        Erro += "O campo Peso é obrigatório! \n"
        '    End If
        'End If
        Return Erro
    End Function

    Public Function LerPesoWS()
        Try
            'FRAMEWORK 2 (WEB SERVICE)
            'Dim ws As New Balanca.Balanca() With { _
            '    .Url = String.Format("http://{0}/LerPeso/Balanca.asmx", lblAcesso.Text) _
            '}

            'Return ws.CapturarPeso(lblModulo.Text, True, True, 4096, 1, lblPorta.Text, hdnBaudRateBalanca.Value)

            'FRAMEWORK 4 (WCF)
            Dim binding As New BasicHttpBinding()
            Dim ipAddress As String = String.Format("http://{0}/LerPeso/Service.svc", lblAcesso.Text)
            Dim endpointAddress As New EndpointAddress(ipAddress)

            Dim ws As New LerPeso.ServiceClient(binding, endpointAddress)

            'Dim pesoTemporario As Integer = 30000
            Return ws.CapturarPeso(lblModulo.Text, True, True, 4096, 1, lblPorta.Text, hdnBaudRateBalanca.Value, hdnDataBits.Value, hdnParity.Value, hdnStopBits.Value)
        Catch ex As Exception
            Return "Não foi possível realizar a pesagem: " & Funcoes.EliminarCaracteresEspeciais(ex.Message)
        End Try
    End Function

    Protected Sub CarregarLaudo(ByVal Pesagem As Integer, ByVal Empresa As String, ByVal Endereco As String)
        LimparCampos()
        objLaudo = New LaudoDeCarregamento(Empresa, Endereco, Pesagem)
        txtLaudo.Text = Pesagem
        SessaoSalvaLaudo()
        'And objLaudo.Cliente IsNot Nothing
        If (objLaudo.Empresa IsNot Nothing) Then
            ddlSituacao.SelectedValue = objLaudo.CodigoSituacao
            'Laudo Cancelado
            If (objLaudo.CodigoSituacao.Equals(1)) Then
                btnCancelar.Enabled = True
            Else
                btnCancelar.Enabled = False
                btnReimprimir.Enabled = False
            End If
            'If objPesagem.Romaneios.Count = 0 Then
            '    txtRomaneio.Text = ""
            'ElseIf objPesagem.Romaneios.Count = 1 Then
            '    txtRomaneio.Text = objPesagem.Romaneios(0).Codigo
            'ElseIf objPesagem.Romaneios.Count > 1 Then
            '    txtRomaneio.Text = "RATEIO"
            'End If

            txtLaudo.Text = objLaudo.Codigo
            txtLaudo.Enabled = False
            'Empresa
            txtCnpjDaEmpresa.Text = objLaudo.Empresa.CodigoFormatado
            txtNomeDaEmpresa.Text = objLaudo.Empresa.Nome
            hdnCodigoEndEmpresa.Value = objLaudo.Empresa.CodigoEndereco
            'Cliente
            'txtNomeDoCliente.Text = objPesagem.Cliente.Nome
            'txtCnpjDoCliente.Text = objPesagem.Cliente.CodigoFormatado
            'hdnCodigoEndCliente.Value = objPesagem.Cliente.CodigoEndereco
            'txtEnderecoDoCliente.Text = objPesagem.Cliente.Endereco
            'txtComplementoDoCliente.Text = objPesagem.Cliente.Complemento
            'txtBairroDoCliente.Text = objPesagem.Cliente.Bairro
            'txtCepDoCliente.Text = objPesagem.Cliente.CEP
            'txtMunicipioDoCliente.Text = objPesagem.Cliente.Cidade
            'txtTelefoneDoCliente.Text = objPesagem.Cliente.Telefone
            'txtEstadoDoCliente.Text = objPesagem.Cliente.Estado.Codigo
            'txtInscricaoDoCliente.Text = objPesagem.Cliente.InscricaoEstadual
            'Pedido
            'txtPedido.Text = objPesagem.Pedido.Codigo
            'imgExtratoPedido.Visible = True
            'CarregarOperacao()

            'If objPesagem.CodigoOperacao > 0 Then ddlOperacao.SelectedValue = objPesagem.CodigoOperacao

            'txtProduto.Text = objPesagem.CodigoProduto & "-" & objPesagem.Produto.Descricao

            'If objPesagem.CodigoSubOperacao > 0 Then
            '    objPesagem.Consolidado = True
            '    Dim dsSaldoPedido As DataSet = objPesagem.SaldoDePedidos(objPesagem, False)

            '    'If dsSaldoPedido.Tables(0).Rows.Count > 0 Then
            '    '    If objPesagem.SubOperacao.Devolucao Then
            '    '        hdnSaldoPedido.Value = dsSaldoPedido.Tables(0).Rows(0).Item("SaldoDevolucao")
            '    '    Else
            '    '        hdnSaldoPedido.Value = dsSaldoPedido.Tables(0).Rows(0).Item("Saldo")
            '    '    End If
            '    'End If
            '    'txtSaldoPedido.Text = hdnSaldoPedido.Value
            'End If

            'hdnAutorizacaoDeRetirada.Value = objPesagem.CodigoAutorizacao
            'txtSafra.Text = objPesagem.Pedido.CodigoSafra

            ''Depositante
            'txtDepositante.Text = objPesagem.Depositante.CodigoFormatado & "-" & objPesagem.Depositante.Nome
            'hdnCodigoEndDepositante.Value = objPesagem.Depositante.CodigoEndereco

            ''Depósito
            'txtDeposito.Text = objPesagem.Deposito.CodigoFormatado & "-" & objPesagem.Deposito.Nome
            'hdnCodigoEndDeposito.Value = objPesagem.Deposito.CodigoEndereco

            'Transportador
            If objLaudo.CodigoTransportador.Length > 0 Then
                '& "-" & objPesagem.Transportador.Nome
                txtTransportador.Text = objLaudo.Transportador.CodigoFormatado
                hdnCodigoEndTransportador.Value = objLaudo.Transportador.CodigoEndereco
            Else
                txtTransportador.Text = ""
                hdnCodigoEndTransportador.Value = 0
            End If

            'Placa
            If objLaudo.CodigoPlaca.Length > 0 Then
                txtPlaca.Text = objLaudo.Placa.Placa01
                hdnViaDeTransporte.Value = objLaudo.Placa.ViaDeTransporte
            Else
                txtPlaca.Text = ""
                hdnViaDeTransporte.Value = 0
            End If

            'Motorista
            If objLaudo.CodigoMotorista.Length > 0 Then
                txtMotorista.Text = objLaudo.Motorista.Nome & "-" & objLaudo.Motorista.CodigoFormatado
                hdnCodigoEndMotorista.Value = objLaudo.Motorista.CodigoEndereco
            Else
                txtMotorista.Text = ""
                hdnCodigoEndMotorista.Value = 0
            End If

            ''Classificação
            'CarregarTabelaClassificacao()
            'ddlTabelaDeClassificacao.SelectedValue = objPesagem.CodigoTabelaDeClassificacao
            'Análises
            txtPrimeiraPesagem.Text = objLaudo.PrimeiraPesagem
            txtSegundaPesagem.Text = objLaudo.SegundaPesagem
            'txtPesoBruto.Text = objPesagem.BrutoBalanca
            'txtDesconto.Text = objPesagem.Desconto
            txtLiquido.Text = objLaudo.Liquido

            'LAUDO ENCERRADO: POSSUI AS DUAS PESAGENS GRAVADAS NA BASE.
            If (txtSegundaPesagem.Text <> "0") Then
                btnEmpresa.Visible = False
                'btnCliente.Visible = False
                'btnPedido.Visible = False
                'btnDepositante.Visible = False
                'btnTesteBalanca.Visible = False
                'btnDeposito.Visible = False
                btnTransportador.Visible = False
                btnPlaca.Visible = False

                'gridDescontos.DataSource = objLaudo.Analises
                'gridDescontos.DataBind()
                'For i = 0 To gridDescontos.Rows.Count - 1
                '    If Now > "2015-12-31" AndAlso gridDescontos.Rows(i).Cells(1).Text = "INTACTA" Then
                '        Tipo.Style.Item("width") = "50%"
                '        INTACTA.Visible = True

                '        If CInt(CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text) = 0 Then
                '            rbIntactaNao.Checked = True
                '            rbIntactaSim.Checked = False
                '            rbIntactaPositivo.Checked = False
                '        End If

                '        If CInt(CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text) = 1 Then
                '            rbIntactaNao.Checked = False
                '            rbIntactaSim.Checked = True
                '            rbIntactaPositivo.Checked = False
                '        End If

                '        If CInt(CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text) = 2 Then
                '            rbIntactaNao.Checked = False
                '            rbIntactaSim.Checked = False
                '            rbIntactaPositivo.Checked = True
                '        End If
                '    End If
                'Next

                ddlSituacao.Enabled = False
                'ddlOperacao.Enabled = False
                'ddlSubOperacao.Enabled = False
                'ddlTabelaDeClassificacao.Enabled = False
                'If objPesagem.TemNota Then txtObservacao.ReadOnly = True

                btnLerPeso.Visible = False
                btnPesagem.Enabled = False
                btnEncerrar.Enabled = False
                If objLaudo.CodigoSituacao = 1 Then btnReimprimir.Enabled = True
                txtPrimeiraPesagem.Enabled = False
                txtPrimeiraPesagem.BorderStyle = BorderStyle.None
                txtPrimeiraPesagem.BorderWidth = Unit.Pixel(0)
                txtSegundaPesagem.Enabled = False
                txtSegundaPesagem.BorderStyle = BorderStyle.None
                txtSegundaPesagem.BorderWidth = Unit.Pixel(0)
                btnLerPeso.Visible = False
                btnIncluir.Enabled = False
                MsgBox(Me.Page, "Laudo: " & txtLaudo.Text & IIf(objLaudo.CodigoSituacao.Equals(1), " Encerrado!", " Cancelado!"))
            Else
                If (Convert.ToInt32(txtSegundaPesagem.Text) > 0) Then
                    If (Convert.ToInt32(txtPrimeiraPesagem.Text) > Convert.ToInt32(txtSegundaPesagem.Text)) Then
                        txtPesoBruto.Text = Convert.ToInt32(txtPrimeiraPesagem.Text) - Convert.ToInt32(txtSegundaPesagem.Text)
                    Else
                        txtPesoBruto.Text = Convert.ToInt32(txtSegundaPesagem.Text) - Convert.ToInt32(txtPrimeiraPesagem.Text)
                    End If
                End If

                'Desnecessario
                'Dim Analises = objPesagem.Analises
                'Analises.ForEach(Function(A)
                '                     A.Indice = 0
                '                     A.Percentual = 0
                '                     Return True
                '                 End Function)
                'gridDescontos.DataSource = objPesagem.Analises
                'gridDescontos.DataBind()

                'For i = 0 To gridDescontos.Rows.Count - 1
                '    CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Enabled = False
                'Next

                'If (objPesagem.Analises.Count > 1) Then
                '    btnCalcular.Visible = True
                '    btnCalcular.Enabled = True
                '    btnIncluir.Enabled = False
                'Else
                '    btnIncluir.Enabled = True
                'End If

                'btnPesagem.Enabled = False
                'btnEncerrar.Enabled = False

                'If (gridDescontos.Rows.Count > 0) Then
                '    If Len(objPesagem.Analises(0).Analise.Opcao) > 0 Then
                '        CType(gridDescontos.Rows(0).FindControl("ddlOpcao"), DropDownList).Focus()
                '    Else
                '        CType(gridDescontos.Rows(0).FindControl("txtPercentual"), TextBox).Focus()
                '    End If
                'End If

                'btnCliente.Visible = True
                'btnPedido.Visible = True
                'btnDepositante.Visible = True
                'btnDeposito.Visible = True
                btnTransportador.Visible = True
                btnPlaca.Visible = True
                btnCancelar.Enabled = True
                If objLaudo.CodigoSituacao = 1 Then btnReimprimir.Enabled = True
            End If

            txtObservacao.Text = objLaudo.Observacoes

            'Devolução ????
            'NOTA FISCAL PRODUTOR
            'If objPesagem.NumeroDaNota Then
            'txtNotaProdutor.Text = objPesagem.NumeroDaNota
            'txtSerieNotaProdutor.Text = objPesagem.SerieDaNota
            'txtPesoNotaProdutor.Text = objPesagem.PesoFiscal
            'If objPesagem.SegundaPesagem > 0 Then
            '    txtNotaProdutor.Enabled = False
            '    txtSerieNotaProdutor.Enabled = False
            '    txtPesoNotaProdutor.Enabled = False
            'End If
            'End If

            'MOVIMENTO
            txtMovimento.Text = objLaudo.Movimento
            txtMovimento.Enabled = False
            btnEmpresa.Visible = False
            Dim entradaSaida As eEntradaSaida = IIf(Convert.ToInt32(txtPrimeiraPesagem.Text) > Convert.ToInt32(txtSegundaPesagem.Text), eEntradaSaida.Entrada, eEntradaSaida.Saida)
            objLaudo.SaidaBalanca = DateTime.Now 'DATA DA SEGUNDA PESAGEM.

            'CARREGA AS SUBOPERAÇÕES CONFORME REGRA DE E E S.
            'Dim subOperacoes As New [Lib].Negocio.ListSubOperacao(" So.Laudo =  'S' AND So.Operacao_id = " & ddlOperacao.SelectedValue)
            'ddlSubOperacao.Items.Clear()
            'If objPesagem.SegundaPesagem > 0 Then
            '    For Each objSubOperacao As SubOperacao In subOperacoes.FindAll(Function(p) p.EntradaSaida = entradaSaida)
            '        ddlSubOperacao.Items.Add(New ListItem(objSubOperacao.CodigoOperacao.ToString("00") & "-" & objSubOperacao.Codigo.ToString("00") & " " & objSubOperacao.Descricao,
            '                                              objSubOperacao.CodigoOperacao & "-" & objSubOperacao.Codigo))
            '    Next
            'Else
            '    For Each objSubOperacao As SubOperacao In subOperacoes
            '        ddlSubOperacao.Items.Add(New ListItem(objSubOperacao.CodigoOperacao.ToString("00") & "-" & objSubOperacao.Codigo.ToString("00") & " " & objSubOperacao.Descricao,
            '                                              objSubOperacao.CodigoOperacao & "-" & objSubOperacao.Codigo))
            '    Next
            'End If

            'If objPesagem.CodigoSubOperacao > 0 Then
            '    Dim intIndice As Integer = ddlSubOperacao.Items.IndexOf(ddlSubOperacao.Items.FindByValue(objPesagem.CodigoOperacao & "-" & objPesagem.CodigoSubOperacao))
            '    If Not intIndice = -1 Then
            '        ddlSubOperacao.SelectedIndex = intIndice
            '    End If
            'End If

            If objLaudo.SegundaPesagem > 0 Then
                lblEntradaSaida.Text = IIf(String.IsNullOrWhiteSpace(objLaudo.SaidaBalanca), "ENTRADA", "SAIDA")
                'ElseIf objPesagem.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                '    lblEntradaSaida.Text = "ENTRADA"
                'ElseIf objPesagem.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                '    lblEntradaSaida.Text = "SAIDA"
            Else
                lblEntradaSaida.Text = "ENTRADA"
            End If

            'HISTÓRICO
            lblHistorico.Text = "Inclusão: " & objLaudo.Movimento.ToString("dd/MM/yyyy") & " - " & objLaudo.UsuarioInclusao.ToString() & "    " & IIf(Not String.IsNullOrWhiteSpace(objLaudo.UsuarioAlteracao), "Alteração: " & objLaudo.DataAlteracao & " - " & objLaudo.UsuarioAlteracao.ToString(), "") &
                "    " & IIf(Not String.IsNullOrWhiteSpace(objLaudo.UsuarioReimpressao), "Reimpressão: " & objLaudo.DataReimpressao & " - " & objLaudo.UsuarioReimpressao.ToString(), "")

            'LIMPA A GRID DA CONSULTA DE LAUDO
            gdvLaudo.DataSource = Nothing
            gdvLaudo.DataBind()

            If (Not String.IsNullOrWhiteSpace(hdnEletronico.Value) AndAlso hdnEletronico.Value AndAlso Not txtPrimeiraPesagem.Enabled AndAlso txtSegundaPesagem.Text.Equals("0")) AndAlso objLaudo.CodigoSituacao = 1 Then
                If String.IsNullOrWhiteSpace(objLaudo.CodigoTransportador) OrElse String.IsNullOrWhiteSpace(objLaudo.CodigoPlaca) Then
                    Session("ssCampo" & HID.Value) = "Transportador"
                    Dim teste As String = Session("ssCampo" & HID.Value)
                    hdnControlePopup.Value = "Transportador"
                    ucConsultaClientes.Limpar()
                    ucConsultaClientes.SetarTipoCliente("7")
                    ucConsultaClientes.SetarTituloDIV("Consulta de Transportador")
                    Popup.ConsultaDeClientes(Me.Page, "objLaudo" & HID.Value)
                    Popup.CenterDialog(Me.Page, "divConsultaCliente")
                Else
                    ChamarModalParaConfirmarPeso(LerPesoWS())
                End If
            Else
                txtPrimeiraPesagem.Enabled = False
                If Not String.IsNullOrWhiteSpace(txtPrimeiraPesagem.Text) AndAlso objLaudo.CodigoSituacao = 1 AndAlso (String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) OrElse txtSegundaPesagem.Text.Equals("0")) Then
                    If String.IsNullOrWhiteSpace(objLaudo.CodigoTransportador) OrElse String.IsNullOrWhiteSpace(objLaudo.CodigoPlaca) Then
                        Session("ssCampo" & HID.Value) = "Transportador"
                        Dim teste As String = Session("ssCampo" & HID.Value)
                        hdnControlePopup.Value = "Transportador"
                        ucConsultaClientes.Limpar()
                        ucConsultaClientes.SetarTipoCliente("7")
                        ucConsultaClientes.SetarTituloDIV("Consulta de Transportador")
                        Popup.ConsultaDeClientes(Me.Page, "objLaudo" & HID.Value)
                        Popup.CenterDialog(Me.Page, "divConsultaCliente")
                    Else
                        txtSegundaPesagem.Enabled = True
                        txtSegundaPesagem.Text = ""
                        Popup.SetFocus(Me.Page, txtSegundaPesagem.ClientID, True)
                    End If
                End If
            End If

            If Not chkLiberado.Checked Then
                btnPesagem.Enabled = False
                btnEncerrar.Enabled = False
                btnCancelar.Enabled = False
                btnIncluir.Enabled = False
            End If
        Else
            MsgBox(Me.Page, "Laudo não encontrado!")
        End If
    End Sub

    Public Sub LimparCampos(Optional ByVal Evento As String = "", Optional ByVal Numero As String = "")
        Session.Remove("objLaudo" & HID.Value)
        Session.Remove("ssLaudo" & HID.Value)
        Session.Remove("ssCampo" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
        ucConsultaEmpresas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaClientesDireto.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        ucConsultaPlacas.SetarHID(HID.Value)
        ucConsultaAutorizacaoDeRetirada.SetarHID(HID.Value)
        ucConsultaEstados.SetarHID(HID.Value)
        ucConsultaCodMunicipios.SetarHID(HID.Value)

        hdnControlePopup.Value = String.Empty
        btnIncluir.Enabled = False
        btnCancelar.Enabled = False
        btnReimprimir.Enabled = True
        btnConsultar.Enabled = True
        'btnPesagem.Visible = True
        btnPesagem.Visible = False
        btnConsultar.Visible = True
        btnLimpar.Enabled = True
        btnPesagem.Enabled = True
        'btnEncerrar.Enabled = True
        btnEncerrar.Enabled = False
        'btnTesteBalanca.Visible = True

        If Not chkLiberado.Checked Then
            btnPesagem.Enabled = False
            btnEncerrar.Enabled = False
        End If

        'chkConsolidarCliente.Checked = False
        'btnCliente.Visible = False
        'btnPedido.Visible = False
        'btnDepositante.Visible = False
        'btnDeposito.Visible = False
        btnTransportador.Visible = False
        btnPlaca.Visible = False
        'btnRateio.Visible = False
        btnEmpresa.Visible = True

        'ddlTabelaDeClassificacao.SelectedIndex = 0
        'ddlTabelaDeClassificacao.Enabled = False

        'imgExtratoPedido.Visible = False
        ddlSituacao.SelectedValue = 1

        txtLaudo.Text = String.Empty
        txtLaudo.Enabled = False
        lblEntradaSaida.Text = String.Empty

        'txtRomaneio.Text = String.Empty

        'txtNomeDoCliente.Text = String.Empty
        'txtInscricaoDoCliente.Text = String.Empty
        'txtCnpjDoCliente.Text = String.Empty
        'txtEnderecoDoCliente.Text = String.Empty
        'txtComplementoDoCliente.Text = String.Empty
        'txtBairroDoCliente.Text = String.Empty
        'txtMunicipioDoCliente.Text = String.Empty
        'txtTelefoneDoCliente.Text = String.Empty
        'txtEstadoDoCliente.Text = String.Empty
        'txtInscricaoDoCliente.Text = String.Empty
        'txtCepDoCliente.Text = String.Empty

        'btnPedido.Enabled = True
        'txtPedido.Text = String.Empty
        'hdnSaldoPedido.Value = 0
        'txtSaldoPedido.Text = hdnSaldoPedido.Value
        'hdnAutorizacaoDeRetirada.Value = 0

        'ddlOperacao.SelectedIndex = 0
        'ddlOperacao.Enabled = False
        'ddlSubOperacao.Items.Clear()
        'pnlRateio.Visible = False

        'txtProduto.Text = String.Empty
        'txtSafra.Text = String.Empty

        'btnDepositante.Enabled = True
        'txtDepositante.Text = String.Empty

        'btnDeposito.Enabled = True
        'txtDeposito.Text = String.Empty

        btnTransportador.Enabled = True
        txtTransportador.Text = String.Empty
        hdnCodigoEndTransportador.Value = 0

        btnPlaca.Enabled = True
        txtPlaca.Text = String.Empty
        txtMotorista.Text = String.Empty
        hdnCodigoEndMotorista.Value = 0
        hdnViaDeTransporte.Value = 0

        txtPrimeiraPesagem.Text = "0"
        txtPrimeiraPesagem.Enabled = False
        txtSegundaPesagem.Text = "0"
        txtSegundaPesagem.Enabled = False
        txtPesoBruto.Text = "0"
        txtDesconto.Text = "0"
        txtLiquido.Text = "0"

        'gridDescontos.DataSource = Nothing
        'gridDescontos.DataBind()
        'btnCalcular.Visible = False

        txtObservacao.ReadOnly = False
        txtObservacao.Text = String.Empty
        lblHistorico.Text = String.Empty

        'radTrocaNao.Checked = True

        'txtNotaProdutor.Enabled = True
        'txtSerieNotaProdutor.Enabled = True
        'txtPesoNotaProdutor.Enabled = True
        'txtNotaProdutor.Text = String.Empty
        'txtSerieNotaProdutor.Text = String.Empty
        'txtPesoNotaProdutor.Text = String.Empty

        txtMovimento.Enabled = True
        txtMovimento.Text = Now.ToString("dd/MM/yyyy")

        objLaudo = New [Lib].Negocio.LaudoDeCarregamento()
        objLaudo.IUD = "I"

        objLaudo.UsuarioInclusao = Session("ssNomeUsuario")

        SessaoSalvaLaudo()
        SetarEmpresa()
        TabContainer1.ActiveTabIndex = 0
        LiberaEmpresa()

        If Evento = "E" Then
            MsgBox(Me.Page, "Laudo " & Numero & " Deletado com Sucesso.", eTitulo.Sucess)
        ElseIf Evento = "I" Then
            MsgBox(Me.Page, "Laudo " & Numero & " Incluido com Sucesso.", eTitulo.Sucess)
        ElseIf Evento = "U" Then
            MsgBox(Me.Page, "Laudo " & Numero & " Alterado com Sucesso.", eTitulo.Sucess)
        ElseIf Evento = "C" Then
            MsgBox(Me.Page, "Laudo " & Numero & " Cancelado com Sucesso.", eTitulo.Sucess)
        End If
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            btnEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub GerenciarBarraBotao()
        'Na conulta do laudo habilita o botão encerrar.
        If (Not String.IsNullOrEmpty(txtPrimeiraPesagem.Text) And Not String.IsNullOrEmpty(txtSegundaPesagem.Text)) Then
            btnEncerrar.Enabled = True
            btnIncluir.Enabled = False
            btnPesagem.Enabled = False
        End If
    End Sub

    Protected Sub Limpar()
        Select Case hdnControlePopup.Value.Replace(",", "")
            Case "Cliente"
                'btnPedido.Enabled = True
                'txtPedido.Text = String.Empty
                'ddlOperacao.Enabled = True
                'ddlSubOperacao.Enabled = True
                'ddlSubOperacao.Items.Clear()
                'txtProduto.Text = String.Empty
                'txtSafra.Text = String.Empty
                'btnDeposito.Enabled = True
                'txtDeposito.Text = String.Empty
                btnTransportador.Enabled = True
                txtTransportador.Text = String.Empty
                btnPlaca.Enabled = True
                txtPlaca.Text = String.Empty
                txtMotorista.Text = String.Empty
                'gridDescontos.DataSource = Nothing
                'gridDescontos.DataBind()
                'btnCalcular.Visible = False
                txtObservacao.Text = String.Empty
                'radTrocaNao.Checked = True
                'txtNotaProdutor.Text = String.Empty
                'txtSerieNotaProdutor.Text = String.Empty
                'txtPesoNotaProdutor.Text = String.Empty
                'ddlTabelaDeClassificacao.SelectedIndex = 0
            Case "Pedido"
                'btnDeposito.Enabled = True
                'txtDeposito.Text = ""
                btnTransportador.Enabled = True
                txtTransportador.Text = ""
        End Select
    End Sub

    Protected Sub CalcularDesconto()
        SessaoRecuperaLaudo()
        objLaudo = New [Lib].Negocio.LaudoDeCarregamento()
        'objPesagem.CodigoTabelaDeClassificacao = ddlTabelaDeClassificacao.SelectedValue

        If (Not txtPrimeiraPesagem.ReadOnly) Then
            If (Convert.ToInt32(txtPrimeiraPesagem.Text) > Convert.ToInt32(txtSegundaPesagem.Text)) Then
                txtPesoBruto.Text = Convert.ToInt32(txtPrimeiraPesagem.Text) - Convert.ToInt32(txtSegundaPesagem.Text)
            Else
                txtPesoBruto.Text = Convert.ToInt32(txtSegundaPesagem.Text) - Convert.ToInt32(txtPrimeiraPesagem.Text)
            End If
        End If

        objLaudo.PrimeiraPesagem = txtPrimeiraPesagem.Text
        objLaudo.SegundaPesagem = txtSegundaPesagem.Text
        'objPesagem.BrutoBalanca = txtPesoBruto.Text
        'objPesagem.CodigoProduto = txtProduto.Text.Split("-").ToArray(0).Trim

        'If (Convert.ToInt32(txtPrimeiraPesagem.Text) > Convert.ToInt32(txtSegundaPesagem.Text)) Then
        '    objPesagem.EntradaSaida = "E"
        'Else
        '    objPesagem.EntradaSaida = "S"
        'End If

        'For i = 0 To gridDescontos.Rows.Count - 1
        '    Dim codanalise As Integer = gridDescontos.Rows(i).Cells(0).Text
        '    Dim Analis As PesagemXAnalises = objPesagem.Analises.Where(Function(s) s.CodigoAnalise = codanalise).First

        '    If Analis.Analise.Opcao.Length = 0 Then
        '        Dim Percentual As String = CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text
        '        Analis.Percentual = IIf(IsNumeric(Percentual), Percentual, 0)
        '    Else
        '        Dim ddlop As DropDownList = CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList)
        '        Analis.Percentual = ddlop.SelectedValue
        '    End If
        'Next
        'Dim Erro As String = objPesagem.Analises.CalcularDescontos()

        txtDesconto.Text = "0"

        'If (String.IsNullOrEmpty(Erro)) Then
        '    gridDescontos.DataSource = objPesagem.Analises
        '    gridDescontos.DataBind()

        '    Dim pDesconto As Decimal = 0
        '    For Each aN In objPesagem.Analises
        '        If Not aN.CodigoAnalise = 6 AndAlso Not aN.CodigoAnalise = 12 Then
        '            pDesconto += aN.Desconto
        '        End If
        '    Next

        '    'Teste da soma dos descontos 27/02/2016
        '    Dim desc As Decimal = objPesagem.Analises.Sum(Function(s) s.Desconto)
        '    If desc <> pDesconto Then
        '        MsgBox(Me.Page, "Erro #12 avise a TI")
        '    End If

        '    txtDesconto.Text = pDesconto
        '    txtLiquido.Text = objPesagem.BrutoBalanca - pDesconto

        '    If txtSegundaPesagem.Enabled Then
        '        For i = 0 To gridDescontos.Rows.Count - 1
        '            CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Enabled = False
        '        Next
        '        btnCalcular.Visible = True
        '        btnCalcular.Enabled = True
        '    End If

        '    btnIncluir.Enabled = True
        '    If objPesagem.EntradaSaida = "S" AndAlso
        '        Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) AndAlso
        '        Convert.ToInt32(txtSegundaPesagem.Text) > 0 Then CarregarAutorizacaoDeRetirada()
        'Else
        '    MsgBox(Me.Page, Erro)
        'End If
    End Sub

    Private Sub SetarEmpresa()
        If Session("objEmpresaLXR") Is Nothing Then Session("objEmpresaLXR") = New [Lib].Negocio.Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))

        SessaoRecuperaLaudo()

        objLaudo.Empresa = CType(Session("objEmpresaLXR"), [Lib].Negocio.Cliente)

        With objLaudo.Empresa
            objLaudo.CodigoEmpresa = .Codigo
            objLaudo.EnderecoEmpresa = .CodigoEndereco
            hdnCodigoEndEmpresa.Value = .CodigoEndereco
            txtNomeDaEmpresa.Text = .Nome
            txtCnpjDaEmpresa.Text = .CodigoFormatado
            txtMunicipioDaEmpresa.Text = .Cidade
            txtEstadoDaEmpresa.Text = .CodigoEstado
        End With
        SessaoSalvaLaudo()
    End Sub

    Public Sub Imprimir(ByVal Tipo As String, ByVal NrCopias As String, ByVal Impressao As String, ByRef objLaudo As [Lib].Negocio.LaudoDeCarregamento)
        Dim ds As New DataSet
        Dim i As Integer
        Dim myDataRow As DataRow

        sql = "SELECT Pesagem.Pesagem_Id AS Laudo, Pesagem.Produto, Pesagem.Placa, Pesagem.EntradaSaida, Pesagem.PrimeiraPesagem, " & vbCrLf &
              "		  Pesagem.SegundaPesagem, Pesagem.BrutoBalanca, Pesagem.BrutoBalanca - Pesagem.Liquido AS Descontos, Pesagem.Liquido, Pesagem.EntradaPatio, " & vbCrLf &
              "		  Pesagem.EntradaBalanca, Pesagem.SaidaBalanca, Pesagem.Movimento, Pesagem.NumeroDaNota AS NotaFiscal, Pesagem.SerieDaNota AS SerieNota, " & vbCrLf &
              "		  Pesagem.PesoFiscal, Pesagem.Observacoes, Produtos.Nome AS NomeProduto, Clientes.Cliente_Id AS CodigoCliente, Clientes.Endereco_Id AS EndCliente, " & vbCrLf &
              "		  Clientes.Nome AS NomeCliente, Clientes.Reduzido AS ReduzidoCliente, Clientes.Endereco AS EnderecoCliente, Clientes.Cidade AS CidadeCliente, " & vbCrLf &
              "		  Clientes.Estado AS EstadoCliente, isnull(Transportes.Cliente_Id,'') AS CodigoTransportador, isnull(Transportes.Endereco_Id,0) AS EndTransportador, " & vbCrLf &
              "       isnull(Transportes.Nome,'') AS NomeTransportador, isnull(Transportes.Reduzido,'') AS ReduzidoTransportador, isnull(Transportes.Endereco,'') AS EnderecoTransportador, " & vbCrLf &
              "       isnull(Transportes.Cidade,'') AS CidadeTransportador, isnull(Transportes.Estado,'') AS EstadoTransportador, " & vbCrLf &
              "		  Depositos.Cliente_Id AS CodigoDeposito, Depositos.Endereco_Id AS EndDeposito, Depositos.Nome AS NomeDeposito, Depositos.Reduzido AS ReduzidoDeposito, " & vbCrLf &
              "       Depositos.Endereco AS EnderecoDeposito, Depositos.Cidade AS CidadeDeposito, Depositos.Estado AS EstadoDeposito, Depositos.Inscricao AS InscricaoDeposito, " & vbCrLf &
              "		  isnull(Placas.Placa01,'') AS Placa01, isnull(Placas.Placa02,'') AS Placa02, isnull(Placas.Placa03,'') AS Placa03, isnull(Placas.CidadePlaca,'') AS CidadePlaca, isnull(Placas.EstadoPlaca,'') AS EstadoPlaca, " & vbCrLf &
              "		  ISNULL(Motorista.Nome,'') AS NomeMotorista, ISNULL(Motorista.Cidade,'') AS CidadeMotorista, ISNULL(Motorista.Estado,'') AS EstadoMotorista, " & vbCrLf &
              "		  ISNULL(Motorista.Habilitacao,'') AS Habilitacao, " & vbCrLf &
              "		  CASE " & vbCrLf &
              "		  	WHEN LEN(ISNULL(Motorista.Cliente_Id,'')) = 0 " & vbCrLf &
              "		  	  THEN '' " & vbCrLf &
              "		  	  ELSE CASE " & vbCrLf &
              "		  			 WHEN LEN(Motorista.Cliente_Id) < 11 " & vbCrLf &
              "		  			   THEN '' " & vbCrLf &
              "		  			   ELSE SUBSTRING(Motorista.Cliente_Id, 1, 3) + '.' + SUBSTRING(Motorista.Cliente_Id, 4, 3) + '.' + SUBSTRING(Motorista.Cliente_Id, 7, 3) + '-' + SUBSTRING(Motorista.Cliente_Id, 10, 2) " & vbCrLf &
              "		  			END " & vbCrLf &
              "		  END AS CpfMotorista, " & vbCrLf &
              "		  isnull(EstadoPlacaMotorista.Descricao,'') AS NomeEstadoMotorista, " & vbCrLf &
              "		  isnull(EstadoPlaca.Descricao,'') AS NomeEstadoPlaca, " & vbCrLf &
              "		  Pedido, Depositos.Numero AS NumeroDeposito, " & vbCrLf &
              "		  Depositos.Complemento AS ComplementoDeposito, Depositos.Bairro AS BairroDeposito, Clientes.Numero AS NumeroCliente, Clientes.Complemento AS ComplementoCliente, " & vbCrLf &
              "		  Clientes.Bairro AS BairroCliente, Clientes.Inscricao AS InscricaoCliente, Empresa.Cliente_Id  AS CodigoEmpresa, Empresa.Endereco_Id AS EndEmpresa, " & vbCrLf &
              "		  Empresa.Nome AS NomeEmpresa, Empresa.Reduzido AS ReduzidoEmpresa, Empresa.Endereco AS EnderecoEmpresa, Empresa.Cidade  AS CidadeEmpresa, " & vbCrLf &
              "		  Empresa.Estado AS EstadoEmpresa, Empresa.Inscricao AS InscricaoEmpresa, Empresa.Numero AS NumeroEmpresa, Empresa.Complemento AS ComplementoEmpresa, " & vbCrLf &
              "		  Empresa.Bairro AS BairroEmpresa, isnull(rXp.Romaneio_id,0) AS Romaneio " & vbCrLf &
              "  FROM Pesagem " & vbCrLf &
              "	INNER JOIN Produtos " & vbCrLf &
              "	   ON Pesagem.Produto = Produtos.Produto_Id " & vbCrLf &
              "	INNER JOIN Clientes " & vbCrLf &
              "	   ON Pesagem.Cliente = Clientes.Cliente_Id " & vbCrLf &
              "	  AND Pesagem.EndCliente = Clientes.Endereco_Id " & vbCrLf &
              "	 LEFT JOIN Clientes AS Transportes " & vbCrLf &
              "	   ON Pesagem.Transportador = Transportes.Cliente_Id " & vbCrLf &
              "	  AND Pesagem.EndTransportador = Transportes.Endereco_Id " & vbCrLf &
              "  LEFT JOIN Clientes AS Motorista " & vbCrLf &
              "	   ON Pesagem.Motorista = Motorista.Cliente_Id " & vbCrLf &
              "	  AND Pesagem.EndMotorista = Motorista.Endereco_Id " & vbCrLf &
              "  LEFT OUTER JOIN Estados AS EstadoPlacaMotorista " & vbCrLf &
              "    ON Motorista.Estado = EstadoPlacaMotorista.Estado_Id " & vbCrLf &
              "	INNER JOIN Clientes AS Depositos " & vbCrLf &
              "	   ON Pesagem.Deposito = Depositos.Cliente_Id " & vbCrLf &
              "	  AND Pesagem.EndDeposito = Depositos.Endereco_Id " & vbCrLf &
              "	 LEFT JOIN Placas " & vbCrLf &
              "	   ON Pesagem.Placa = Placas.Placa_Id " & vbCrLf &
              "	 LEFT JOIN Estados AS EstadoPlaca " & vbCrLf &
              "	   ON Placas.EstadoPlaca = EstadoPlaca.Estado_Id " & vbCrLf &
              "	Inner Join Clientes as Empresa " & vbCrLf &
              "	   on Empresa.cliente_id = Pesagem.Empresa_id " & vbCrLf &
              "	  and Empresa.endereco_id = Pesagem.EndEmpresa_id " & vbCrLf &
              "	 LEFT JOIN RomaneiosXPesagens rXp " & vbCrLf &
              "	   on rXp.Empresa_id     = Pesagem.Empresa_id " & vbCrLf &
              "	  and rXp.EndEmpresa_Id = Pesagem.EndEmpresa_id " & vbCrLf &
              "	  and rXp.Pesagem_id    = Pesagem.Pesagem_id " & vbCrLf &
              "	  and rXp.Sequencia_Id  = Pesagem.Sequencia_Id " & vbCrLf &
              " WHERE (Pesagem.Empresa_Id = '" & objLaudo.CodigoEmpresa & "') " & vbCrLf &
              "   AND (Pesagem.EndEmpresa_Id = " & objLaudo.EnderecoEmpresa & ") " & vbCrLf &
              "   AND (Pesagem.Pesagem_Id = " & objLaudo.Codigo & ") " & vbCrLf &
              "   AND (Pesagem.Sequencia_Id = 0)"

        ds = Banco.ConsultaDataSet(sql, "Laudo")


        'sql = "Select CLA.Analise_Id AS Analise, " & vbCrLf &
        '      "       case " & vbCrLf &
        '      "         when Len(isnull(Analises.Opcao,'')) > 0 " & vbCrLf &
        '      "           then Analises.Descricao + ' - ' + (select valor from fnStringToArray(Analises.Opcao,';') where left(valor,len(convert(int, pxa.percentual))+1) = convert(varchar,convert(int,pxa.percentual)) + '-')" & vbCrLf &
        '      "           else Analises.Descricao " & vbCrLf &
        '      "       end AS Descricao, " & vbCrLf &
        '      "       isnull(pxa.percentual,0) as Percentual," & vbCrLf &
        '      "       isnull(pxa.Indice,CLA.Indice) as indice," & vbCrLf &
        '      "       isnull(pxa.Desconto,0) as Desconto " & vbCrLf &
        '      "  from classificacoes CLA " & vbCrLf &
        '      "  left join Analises " & vbCrLf &
        '      "    on CLA.Analise_Id = Analises.Analise_Id " & vbCrLf &
        '      "  left join Pesagem p" & vbCrLf &
        '      "	   on p.Empresa_Id      ='" & objPesagem.CodigoEmpresa & "'" & vbCrLf &
        '      "	  and p.EndEmpresa_Id   = " & objPesagem.EnderecoEmpresa & vbCrLf &
        '      "	  and p.Pesagem_Id      = " & objPesagem.Codigo & vbCrLf &
        '      "   and p.sequencia_id    = " & objPesagem.Sequencia & vbCrLf &
        '      "  Left Join PesagemXAnalises pxa" & vbCrLf &
        '      "    ON pxa.Empresa_Id    = p.Empresa_Id " & vbCrLf &
        '      "   AND pxa.EndEmpresa_Id = p.EndEmpresa_Id " & vbCrLf &
        '      "   AND pxa.Pesagem_Id    = p.Pesagem_Id " & vbCrLf &
        '      "   AND pxa.Sequencia_Id  = p.Sequencia_Id " & vbCrLf &
        '      "   And pxa.Analise_id    = cla.Analise_Id" & vbCrLf &
        '      "  Where CLA.Tabela_id    = " & objLaudo.CodigoTabelaDeClassificacao & vbCrLf &
        '      "    and CLA.Produto_id   ='" & objLaudo.CodigoProduto & "'" & vbCrLf &
        '      "    and CLA.Sequencia_Id = 1 " & vbCrLf &
        '      "  ORDER BY CLA.Analise_Id "

        'ds.Merge(Banco.ConsultaDataSet(sql, "Analises"))

        If Left(Session("ssEmpresa"), 8) = "44979506" AndAlso Tipo = "2" Then
            For Each row As DataRow In ds.Tables("Analises").Rows
                row("indice") = 0
            Next
        End If

        If NrCopias > 1 Then
            For i = 1 To NrCopias - 1
                myDataRow = ds.Tables(0).NewRow
                myDataRow("Laudo") = ds.Tables(0).Rows(0).Item("Laudo")
                myDataRow("Produto") = ds.Tables(0).Rows(0).Item("Produto")
                myDataRow("Placa") = ds.Tables(0).Rows(0).Item("Placa")
                myDataRow("EntradaSaida") = ds.Tables(0).Rows(0).Item("EntradaSaida")
                myDataRow("PrimeiraPesagem") = ds.Tables(0).Rows(0).Item("PrimeiraPesagem")
                myDataRow("SegundaPesagem") = ds.Tables(0).Rows(0).Item("SegundaPesagem")
                myDataRow("BrutoBalanca") = ds.Tables(0).Rows(0).Item("BrutoBalanca")
                myDataRow("Descontos") = ds.Tables(0).Rows(0).Item("Descontos")
                myDataRow("Liquido") = ds.Tables(0).Rows(0).Item("Liquido")
                myDataRow("EntradaPatio") = ds.Tables(0).Rows(0).Item("EntradaPatio")
                myDataRow("EntradaBalanca") = ds.Tables(0).Rows(0).Item("EntradaBalanca")
                myDataRow("SaidaBalanca") = ds.Tables(0).Rows(0).Item("SaidaBalanca")
                myDataRow("Movimento") = ds.Tables(0).Rows(0).Item("Movimento")
                myDataRow("NotaFiscal") = ds.Tables(0).Rows(0).Item("NotaFiscal")
                myDataRow("SerieNota") = ds.Tables(0).Rows(0).Item("SerieNota")
                myDataRow("PesoFiscal") = ds.Tables(0).Rows(0).Item("PesoFiscal")
                myDataRow("Observacoes") = ds.Tables(0).Rows(0).Item("Observacoes")
                myDataRow("NomeProduto") = ds.Tables(0).Rows(0).Item("NomeProduto")
                myDataRow("CodigoCliente") = ds.Tables(0).Rows(0).Item("CodigoCliente")
                myDataRow("EndCliente") = ds.Tables(0).Rows(0).Item("EndCliente")
                myDataRow("NomeCliente") = ds.Tables(0).Rows(0).Item("NomeCliente")
                myDataRow("ReduzidoCliente") = ds.Tables(0).Rows(0).Item("ReduzidoCliente")
                myDataRow("EnderecoCliente") = ds.Tables(0).Rows(0).Item("EnderecoCliente")
                myDataRow("CidadeCliente") = ds.Tables(0).Rows(0).Item("CidadeCliente")
                myDataRow("EstadoCliente") = ds.Tables(0).Rows(0).Item("EstadoCliente")
                myDataRow("CodigoTransportador") = ds.Tables(0).Rows(0).Item("CodigoTransportador")
                myDataRow("EndTransportador") = ds.Tables(0).Rows(0).Item("EndTransportador")
                myDataRow("NomeTransportador") = ds.Tables(0).Rows(0).Item("NomeTransportador")
                myDataRow("ReduzidoTransportador") = ds.Tables(0).Rows(0).Item("ReduzidoTransportador")
                myDataRow("EnderecoTransportador") = ds.Tables(0).Rows(0).Item("EnderecoTransportador")
                myDataRow("CidadeTransportador") = ds.Tables(0).Rows(0).Item("CidadeTransportador")
                myDataRow("EstadoTransportador") = ds.Tables(0).Rows(0).Item("EstadoTransportador")
                myDataRow("CodigoDeposito") = ds.Tables(0).Rows(0).Item("CodigoDeposito")
                myDataRow("EndDeposito") = ds.Tables(0).Rows(0).Item("EndDeposito")
                myDataRow("NomeDeposito") = ds.Tables(0).Rows(0).Item("NomeDeposito")
                myDataRow("ReduzidoDeposito") = ds.Tables(0).Rows(0).Item("ReduzidoDeposito")
                myDataRow("EnderecoDeposito") = ds.Tables(0).Rows(0).Item("EnderecoDeposito")
                myDataRow("CidadeDeposito") = ds.Tables(0).Rows(0).Item("CidadeDeposito")
                myDataRow("EstadoDeposito") = ds.Tables(0).Rows(0).Item("EstadoDeposito")
                myDataRow("InscricaoDeposito") = ds.Tables(0).Rows(0).Item("InscricaoDeposito")

                myDataRow("CodigoEmpresa") = ds.Tables(0).Rows(0).Item("CodigoEmpresa")
                myDataRow("EndEmpresa") = ds.Tables(0).Rows(0).Item("EndEmpresa")
                myDataRow("NomeEmpresa") = ds.Tables(0).Rows(0).Item("NomeEmpresa")
                myDataRow("ReduzidoEmpresa") = ds.Tables(0).Rows(0).Item("ReduzidoEmpresa")
                myDataRow("EnderecoEmpresa") = ds.Tables(0).Rows(0).Item("EnderecoEmpresa")
                myDataRow("CidadeEmpresa") = ds.Tables(0).Rows(0).Item("CidadeEmpresa")
                myDataRow("EstadoEmpresa") = ds.Tables(0).Rows(0).Item("EstadoEmpresa")
                myDataRow("InscricaoEmpresa") = ds.Tables(0).Rows(0).Item("InscricaoEmpresa")
                myDataRow("NumeroEmpresa") = ds.Tables(0).Rows(0).Item("NumeroEmpresa")
                myDataRow("ComplementoEmpresa") = ds.Tables(0).Rows(0).Item("ComplementoEmpresa")
                myDataRow("BairroEmpresa") = ds.Tables(0).Rows(0).Item("BairroEmpresa")

                myDataRow("Placa01") = ds.Tables(0).Rows(0).Item("Placa01")
                myDataRow("Placa02") = ds.Tables(0).Rows(0).Item("Placa02")
                myDataRow("Placa03") = ds.Tables(0).Rows(0).Item("Placa03")
                myDataRow("CidadePlaca") = ds.Tables(0).Rows(0).Item("CidadePlaca")
                myDataRow("EstadoPlaca") = ds.Tables(0).Rows(0).Item("EstadoPlaca")
                myDataRow("NomeMotorista") = ds.Tables(0).Rows(0).Item("NomeMotorista")
                myDataRow("CidadeMotorista") = ds.Tables(0).Rows(0).Item("CidadeMotorista")
                myDataRow("EstadoMotorista") = ds.Tables(0).Rows(0).Item("EstadoMotorista")
                myDataRow("Habilitacao") = ds.Tables(0).Rows(0).Item("Habilitacao")
                myDataRow("CpfMotorista") = ds.Tables(0).Rows(0).Item("CpfMotorista")
                myDataRow("NomeEstadoPlaca") = ds.Tables(0).Rows(0).Item("NomeEstadoPlaca")
                myDataRow("NomeEstadoMotorista") = ds.Tables(0).Rows(0).Item("NomeEstadoMotorista")
                myDataRow("Pedido") = ds.Tables(0).Rows(0).Item("Pedido")
                myDataRow("NumeroDeposito") = ds.Tables(0).Rows(0).Item("NumeroDeposito")
                myDataRow("ComplementoDeposito") = ds.Tables(0).Rows(0).Item("ComplementoDeposito")
                myDataRow("BairroDeposito") = ds.Tables(0).Rows(0).Item("BairroDeposito")
                myDataRow("NumeroCliente") = ds.Tables(0).Rows(0).Item("NumeroCliente")
                myDataRow("ComplementoCliente") = ds.Tables(0).Rows(0).Item("ComplementoCliente")
                myDataRow("BairroCliente") = ds.Tables(0).Rows(0).Item("BairroCliente")
                myDataRow("InscricaoCliente") = ds.Tables(0).Rows(0).Item("InscricaoCliente")

                myDataRow("Romaneio") = ds.Tables(0).Rows(0).Item("Romaneio")

                ds.Tables(0).Rows.Add(myDataRow)
            Next i
        End If

        Dim objListLaudo = New [Lib].Negocio.ListLaudoDeCarregamento(objLaudo, "", "")

        Dim t = objListLaudo

        Dim rpt As New ReportDocument()

        Try
            If Tipo = "1" Then
                rpt.FileName = Server.MapPath("~/Reports/Cr_LaudoDePesagemBruto.rpt")
            Else
                rpt.FileName = Server.MapPath("~/Reports/Cr_LaudoDePesagem.rpt")
            End If
            rpt.Load(rpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

            'If Tipo = "1" Then
            '    Dim Crv_LaudoDePesagemBruto As New CrystalDecisions.Web.CrystalReportViewer
            'Else
            '    Dim Crv_LaudoDePesagem As New CrystalDecisions.Web.CrystalReportViewer
            'End If

            Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

            rpt.SetDataSource(ds)

            Dim crparametervalues As ParameterValues
            Dim crparameterdiscretevalue As ParameterDiscreteValue
            Dim crparameterfielddefinitions As ParameterFieldDefinitions
            Dim crparameterfielddefinition As ParameterFieldDefinition

            crparameterfielddefinitions = rpt.DataDefinition.ParameterFields()

            crparameterfielddefinition = crparameterfielddefinitions.Item("Reemissao")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            If Tipo = "3" Then
                crparameterdiscretevalue.Value = "REEMISSÃO"
            Else
                crparameterdiscretevalue.Value = ""
            End If
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            If Impressao = "I" Then
                rpt.PrintOptions.PrinterName = HttpContext.Current.Session("printerName")
                rpt.PrintToPrinter(1, True, 1, NrCopias / 2)
            Else
                If Dir(arquivo).Length > 0 Then Kill(arquivo)

                rpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

                If IO.File.Exists(arquivo) Then
                    ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo & "');", True)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            rpt.Close()
            rpt.Dispose()
        End Try
    End Sub

    Public Sub ChamarModalParaConfirmarPeso(ByVal peso As String)
        Try
            Dim primeiraPesagem As Integer = 1

            If Not IsNumeric(peso) Then
                Throw New Exception(peso)
            End If

            If (String.IsNullOrWhiteSpace(txtLaudo.Text)) Then
                If String.IsNullOrWhiteSpace(peso) Then
                    txtPrimeiraPesagem.Text = "0"
                Else
                    txtPrimeiraPesagem.Text = Convert.ToInt32(peso)
                End If

                txtPesoBruto.Text = txtPrimeiraPesagem.Text
                txtLiquido.Text = txtPrimeiraPesagem.Text
            Else
                If String.IsNullOrWhiteSpace(peso) Then
                    txtSegundaPesagem.Text = "0"
                Else
                    txtSegundaPesagem.Text = Convert.ToInt32(peso)
                End If

                primeiraPesagem = 2
                If (Convert.ToInt32(txtPrimeiraPesagem.Text) > Convert.ToInt32(txtSegundaPesagem.Text)) Then
                    txtPesoBruto.Text = Convert.ToInt32(txtPrimeiraPesagem.Text) - Convert.ToInt32(txtSegundaPesagem.Text)
                    txtLiquido.Text = txtPesoBruto.Text
                Else
                    txtPesoBruto.Text = Convert.ToInt32(txtSegundaPesagem.Text) - Convert.ToInt32(txtPrimeiraPesagem.Text)
                    txtLiquido.Text = txtPesoBruto.Text
                End If
            End If

            'If (Convert.ToInt32(txtSegundaPesagem.Text) > 0) Then
            '    Dim subOperacaoDoPedido As String = ddlSubOperacao.SelectedValue
            '    ddlSubOperacao.Items.Clear()
            '    'Carrega as SubOperações conforme regra de E e S.
            '    Dim entradaSaida As String = IIf(Convert.ToInt32(txtPrimeiraPesagem.Text) > Convert.ToInt32(txtSegundaPesagem.Text), "E", "S")
            '    Dim subOperacoes As New [Lib].Negocio.ListSubOperacao(" So.Operacao_Id = " & ddlOperacao.SelectedValue & " AND So.Laudo =  'S' AND So.EntradaSaida = '" & entradaSaida & "'")
            '    ddlSubOperacao.Items.Insert(0, "")
            '    For Each objSubOperacao As SubOperacao In subOperacoes
            '        ddlSubOperacao.Items.Add(New ListItem(objSubOperacao.Operacao.Codigo.ToString("00") & "-" & objSubOperacao.Codigo.ToString("00") & " " & objSubOperacao.Descricao,
            '                                              objSubOperacao.Operacao.Codigo & "-" & objSubOperacao.Codigo))
            '    Next
            '    'btnRateio.Visible = True 'Rateio não implementado até o momento
            '    Dim intIndice As Integer = ddlSubOperacao.Items.IndexOf(ddlSubOperacao.Items.FindByValue(subOperacaoDoPedido))
            '    If Not intIndice = -1 Then
            '        ddlSubOperacao.SelectedIndex = intIndice
            '    End If

            '    If String.IsNullOrWhiteSpace(ddlSubOperacao.SelectedItem.Text) Then lblEntradaSaida.Text = String.Empty
            'Else
            '    btnIncluir.Enabled = True
            'End If

            If Not String.IsNullOrWhiteSpace(peso) Then
                ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "ConfirmarPeso(" & primeiraPesagem & ", '" & Convert.ToInt32(peso).ToString & "');", True)
            Else
                btnIncluir.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Public Sub Incluir()
    '    SessaoRecuperaLaudo()
    '    'dados da pesagem
    '    objPesagem.IUD = "I" 'Insert
    '    objPesagem.Sequencia = 0 'inicia com zero na inclusão.
    '    objPesagem.EntradaPatio = DateTime.Now
    '    objPesagem.EntradaBalanca = DateTime.Now
    '    objPesagem.UsuarioInclusao = Session("ssNomeUsuario").ToString
    '    objPesagem.Analises = Nothing
    '    objPesagem.CodigoSituacao = ddlSituacao.SelectedValue
    '    'objPesagem.CodigoOperacao = ddlOperacao.SelectedValue.ToString
    '    'objPesagem.CodigoSubOperacao = ddlSubOperacao.SelectedValue.Split("-").ToArray(1).Trim 'dúvida do pedido ou pode ser alterado?
    '    objPesagem.PrimeiraPesagem = CType(txtPrimeiraPesagem.Text, Decimal)
    '    objPesagem.PrimeiraPesagem = txtPrimeiraPesagem.Text
    '    objPesagem.Movimento = txtMovimento.Text
    '    'Nota Produtor
    '    'If (Not String.IsNullOrEmpty(txtNotaProdutor.Text)) Then
    '    '    objPesagem.NumeroDaNota = txtNotaProdutor.Text
    '    '    objPesagem.SerieDaNota = txtSerieNotaProdutor.Text
    '    '    objPesagem.PesoFiscal = txtPesoNotaProdutor.Text
    '    'End If
    '    objPesagem.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtObservacao.Text)
    '    objPesagem.TemRomaneio = False 'Default(N) Não utilizado na pesagem
    '    objPesagem.RegistroMestre = 0 'Sempre 0
    '    objPesagem.PesagemDeTerceiros = 0 'Default(0)
    '    If objPesagem.Salvar() Then
    '        MsgBox(Me.Page, "Pesagem realizada com Sucesso.", eTitulo.Sucess)
    '    Else
    '        MsgBox(Me.Page, "Erro ao Salvar Pesagem: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
    '    End If
    '    LimparCampos()
    'End Sub

    Protected Sub imgImprimir_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        If Funcoes.VerificaPermissao("LaudoDePesagem", "RELATORIO") Then
            ImprimirLaudo()

            'Dim img As ImageButton = CType(sender, ImageButton)
            'Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)
            'objLaudo = New LaudoDeCarregamento(gdvLaudo.Rows(row.RowIndex).Cells(2).Text, gdvLaudo.Rows(row.RowIndex).Cells(3).Text, gdvLaudo.Rows(row.RowIndex).Cells(1).Text)
            'SessaoSalvaLaudo()
            ''AndAlso objLaudo.Analises IsNot Nothing
            'If objLaudo.SegundaPesagem > 0 Then
            '    Imprimir("2", txtNumeroDeCopias.Text, "A", objLaudo)
            'Else
            '    Imprimir("1", txtNumeroDeCopias.Text, "A", objLaudo)
            'End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para reimpressão da pesagem")
        End If
    End Sub


    Public Sub Cancelar()
        objLaudo.IUD = "D" 'Insert
        objLaudo.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtObservacao.Text)
        objLaudo.UsuarioCancelamento = Session("ssNomeUsuario").ToString
        objLaudo.DataCancelamento = Now()

        'For Each ron As [Lib].Negocio.Romaneio In objPesagem.Romaneios
        '    ron.IUD = "D"
        'Next

        If objLaudo.Salvar() Then
            MsgBox(Me.Page, "Pesagem cancelada com Sucesso.", eTitulo.Sucess)
        Else
            MsgBox(Me.Page, "Erro ao Salvar Pesagem: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
        End If
        LimparCampos()
    End Sub

    Private Sub IncluirEventos()
        Session.Remove("objEmpresaLXR")
        Dim fwPrintSetting As New System.Drawing.Printing.PrinterSettings
        Dim nCnt As Integer

        With ddlImpressora.Items
            For nCnt = 0 To (PrinterSettings.InstalledPrinters.Count - 1)
                .Add(PrinterSettings.InstalledPrinters.Item(nCnt))
            Next
        End With

        Funcoes.InserirLinhaEmBranco(ddlImpressora)

        Dim dsImpressora As New DataSet
        sql = "Select Descricao FROM RotinasDeImpressao WHERE Rotina_Id = 'LAUDO'"
        dsImpressora = Banco.ConsultaDataSet(sql, "Impressora")

        If dsImpressora Is Nothing Then
        ElseIf dsImpressora.Tables(0).Rows.Count = 0 Then
        Else
            With ddlImpressora
                .SelectedIndex = .Items.IndexOf(.Items.FindByValue(dsImpressora.Tables(0).Rows(0).Item("Descricao")))
            End With
        End If
    End Sub

    Private Sub BuscarSituacao()
        Dim objSituacao As New [Lib].Negocio.ListSituacao(True)

        ddlSituacao.DataTextField = "Descricao"
        ddlSituacao.DataValueField = "Codigo"
        ddlSituacao.DataSource = objSituacao.ToArray
        ddlSituacao.DataBind()

        Funcoes.InserirLinhaEmBranco(ddlSituacao)
    End Sub

    Private Sub ImprimirLaudo()
        SessaoRecuperaLaudo()

        'If String.IsNullOrWhiteSpace(objOP.ProdutosdaProducao.Count = 0) Then
        '    MsgBox(Me.Page, "Ordem de Produção não foi selecionada!", eTitulo.Info)
        '    Exit Sub
        'ElseIf objOP.ItensDeEspecificacao.Count = 0 Then
        '    MsgBox(Me.Page, "Ordem de Produção não tem Especificações!", eTitulo.Info)
        '    Exit Sub
        'End If

        Dim html = New StringBuilder()

        html.Append("<html>")
        html.Append("<head>")
        html.Append("<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>")
        html.Append("<title>Laudo de Análises</title>")
        html.Append("<style type='text/css'>")

        'styles
        html.Append("@page { size: A4; margin: 11mm 17mm 17mm 17mm; }")
        html.Append("@media print { html, body { width: 210mm; height: 297mm; padding-top: 15; } }")
        html.Append("html, body { width: 99.99%; height: 100%; margin: 0; padding-top: 15; }")
        html.Append("body { display: flex; flex-direction: column; }")
        html.Append("main { width: 100%; flex-grow: 1; }")
        html.Append(".table-report { font-family: 'Times New Roman', Times, serif; border-collapse: collapse; width: 100%; font-size: 12px; }")
        html.Append(".table-report td, .table-report th { border-right: 0.2px solid#696969; border-left: 0.2px solid #696969; border-bottom: 0.2px solid #696969; padding: 3px; }")
        html.Append(".table-report tr > th { background-color: #f2f2f2; }")
        html.Append(".table-report th { text-align: left; color: black; }")
        html.Append(".table-report tr { page-break-inside: avoid; }")
        html.Append("ol { padding: 5px; }")
        html.Append("ol li { padding: 3px; margin-left: 15px; }")
        html.Append("ul { padding: 5px; list-style-type: none; color: black; }")
        html.Append("ul li { margin: 5px; }")
        html.Append(".border-top { border-top: 0.2px solid #696969; }")
        html.Append(".td-align-top { font-weight: bold; text-align: start; vertical-align:top!important; width: 60%; } ")
        html.Append("p { padding: 3px; margin: 0px; } ")
        html.Append(".td-align-bottom { text-align: End; vertical-align:top!important; font-size: 8px; width: 20%;} ")
        html.Append(".font-size-14 { font-size: 14px; }")
        html.Append(".width-50 {width: 50%; } ")
        html.Append(".content {position: relative; min-height: 100px; font-size: 8px; } ")
        html.Append(".content-bottom { position: absolute; bottom: 0; right: 0; } ")
        html.Append(".content-top { position: absolute; top: 0; right: 0; } ")
        'end styles
        html.Append("</style>")
        html.Append("</head>")
        html.Append("<body>")
        html.Append("<main>")

        Dim link As String = "http://www.ngssolucoes.com.br/Download/UltimaVersao/" & objLaudo.Empresa.Imagem

        'Cabeçalho
        html.Append("<table class='table-report'>")
        html.Append("<tr Class='border-top'>")
        html.Append("<td> <img src='" & link & "' width='200' height='100'></td>")
        html.Append("<td Class='td-align-top'>")
        html.Append("<p Class='padding-p'>" & objLaudo.Empresa.Codigo & " - " & objLaudo.Empresa.Nome & "</p>")
        html.Append("<p Class='padding-p'>" & objLaudo.Empresa.Cidade & "/" & objLaudo.Empresa.CodigoEstado & "</p>")
        html.Append("<p Class='padding-p'>" & objLaudo.Empresa.OutrosTelefones & "</p>")

        If objLaudo.CodigoSituacao = eSituacao.Excluido Then
            html.Append("<p Class='padding-p'>" & objLaudo.Empresa.Email & "</p>")
            html.Append("<p Class='padding-p' style='color:#FF0000'>***** ORDEM CANCELADA *****</p>")
        Else
            html.Append("<p Class='padding-p'>" & objLaudo.Empresa.Email & "</p>")
        End If

        html.Append("</td>")
        'html.Append("<td style='width: 20%;'>")
        'html.Append("<div class='content'>")
        'html.Append("<div class='content-top'>")
        'html.Append("<p>Usuário: " & IIf(String.IsNullOrWhiteSpace(objOP.UsuarioAlteracao), objOP.UsuarioInclusao, objOP.UsuarioAlteracao) & "</p>")
        'html.Append("</div>")
        'html.Append("<div class='content-bottom'>")
        'html.Append("<p>Data: " & Now().ToString("dd/MM/yyyy") & "</p>")
        'html.Append("<p>Hora: " & Now().ToString("HH:mm") & "</p>")
        'html.Append("</div>")
        'html.Append("</div>")
        'html.Append("</td>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("<br>")
        html.Append("<br>")

        html.Append("<table class='table-report'>")
        html.Append("<tr style='border-top: 0.2px solid #696969;'>")
        html.Append("<th colspan='6' class='font-size-14' style='text-align:center;'>LAUDO: " & objLaudo.Codigo & "</th>")
        html.Append("</tr>")

        html.Append("<br>")

        html.Append("<tr>")
        html.Append("<td>Ordem</td>")
        html.Append("<td>Movimento</td>")
        html.Append("<td>Entrada</td>")
        html.Append("<td>Saida</td>")
        html.Append("</tr>")

        html.Append("<br>")
        html.Append("<tr>")
        html.Append("<td>" & objLaudo.CodigoOrdemDeCarregamento & "</td>")
        html.Append("<td>" & objLaudo.Movimento.ToString("dd/MM/yyyy") & "</td>")
        html.Append("<td>" & objLaudo.EntradaBalanca.ToString("dd/MM/yyyy") & "</td>")
        html.Append("<td>" & objLaudo.SaidaBalanca.ToString("dd/MM/yyyy") & "</td>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("<table class='table-report'>")
        html.Append("<tr style='border-top: 0.2px solid #696969;'>")
        html.Append("<th colspan='4' class='font-size-14' style='text-align:center;'>PESAGEM</th>")
        html.Append("</tr>")

        html.Append("<br>")
        html.Append("<tr>")
        html.Append("<td>Primeira</td>")
        html.Append("<td>Segunda</td>")
        html.Append("<td>Liquido</td>")
        html.Append("</tr>")

        html.Append("<br>")
        html.Append("<tr>")
        html.Append("<td>" & objLaudo.PrimeiraPesagem & "</td>")
        html.Append("<td>" & objLaudo.SegundaPesagem & "</td>")
        html.Append("<td>" & objLaudo.Liquido & "</td>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("<br>")

        html.Append("<table class='table-report'>")
        html.Append("<tr style='border-top: 0.2px solid #696969;'>")
        html.Append("<th colspan='4' class='font-size-14' style='text-align:center;'>DADOS DO TRANSPORTADOR</th>")
        html.Append("</tr>")

        html.Append("<br>")
        html.Append("<tr>")
        html.Append("<td>Transportador</td>")
        html.Append("<td>Motorista</td>")
        html.Append("<td>Placa</td>")
        html.Append("</tr>")

        html.Append("<br>")
        html.Append("<tr>")
        html.Append("<td>" & objLaudo.CodigoTransportador & " - " & objLaudo.Transportador.Nome & "</td>")
        html.Append("<td>" & objLaudo.CodigoMotorista & " - " & objLaudo.Motorista.Nome & "</td>")
        html.Append("<td>" & objLaudo.CodigoPlaca & "</td>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("<table class='table-report'>")
        html.Append("<tr>")
        html.Append("<td>Observações</td>")
        html.Append("</tr>")

        html.Append("<br>")
        html.Append("<tr>")
        html.Append("<td>" & objLaudo.Observacoes & "</td>")
        html.Append("</tr>")
        html.Append("</table>")


        ''Produto
        'html.Append("<table class='table-report'>")

        'For Each produto As OrdemParaProducaoXProduto In objOP.ProdutosdaProducao

        '    html.Append("<tr style='border-top: 0.2px solid #696969;'>")
        '    html.Append("<th colspan='4' class='font-size-14'>PRODUTO: " & produto.Produto.Nome & "</th>")
        '    html.Append("</tr>")

        '    html.Append("<tr>")
        '    html.Append("<td>Lote Nº</td>")
        '    html.Append("<td>Data</td>")
        '    html.Append("<td>Validade</td>")
        '    html.Append("</tr>")

        '    html.Append("<tr>")
        '    html.Append("<td>" & produto.Lote & "</td>")
        '    html.Append("<td>" & objOP.Movimento.ToString("dd/MM/yyyy") & "</td>")
        '    html.Append("<td>" & objOP.Validade.ToString("dd/MM/yyyy") & "</td>")
        '    html.Append("</tr>")

        'Next

        'html.Append("</table>")

        'html.Append("<br>")
        'html.Append("<br>")

        ''Especificações
        'html.Append("<table class='table-report'>")
        'html.Append("<tr style='border-top: 0.2px solid #696969;'>")
        'html.Append("<th> ESPECIFICAÇÕES</th>")
        'html.Append("<th style='text-align:center;'>Faixa Inicial</th>")
        'html.Append("<th style='text-align:center;'>Faixa Final</th>")
        'html.Append("<th style='text-align:right;'> RESULTADOS</th>")
        'html.Append("</tr>")

        'Dim prdEspecificacao = ""

        'For Each ep In objOP.ItensDeEspecificacao

        '    If Not ep.CodigoProdutoProducao = prdEspecificacao Then
        '        html.Append("<tr>")
        '        html.Append("<td colspan='4' style='font-weight: bold;'> " & ep.ProdutoProducao.Nome & "</td>")
        '        html.Append("</tr>")
        '        prdEspecificacao = ep.CodigoProdutoProducao
        '    End If

        '    Dim descricao = ep.EspecificacaoDoProduto.Descricao
        '    Dim faixaInicial = ep.FaixaInicial
        '    Dim faixaFinal = ep.FaixaFinal
        '    Dim resultado = ep.Resultado
        '    html.Append("<tr>")
        '    html.Append("<td> " & descricao & "</td>")
        '    html.Append("<td> " & faixaInicial & "</td>")
        '    html.Append("<td> " & faixaFinal & "</td>")
        '    html.Append("<td style='text-align:right;'>" & resultado & " </td>")
        '    html.Append("</tr>")
        'Next

        'html.Append("</table>")

        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")

        html.Append("<table>")
        html.Append("<tr>")
        html.Append("<td style='text-align:center;'>_________________________________________</td>")
        html.Append("</tr>")
        html.Append("<tr>")
        html.Append("<td style='text-align:center;'>CONTROLE DE QUALIDADE</td>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("</main>")
        html.Append("</body>")
        html.Append("</html>")

        Dim strCaminho As String = Server.MapPath("~/Files/LaudoDeAnalises.html")
        If Dir(strCaminho).Length > 0 Then Kill(strCaminho)

        Using strm As New StreamWriter(Server.MapPath("~/Files/LaudoDeAnalises.html"), True)
            strm.WriteLine(html)
            strm.Close()
            strm.Dispose()
        End Using

        Dim strNomeArquivo As String = "Files/LaudoDeAnalises.pdf"
        Dim pathPDF = Server.MapPath(strNomeArquivo)
        If Dir(pathPDF).Length > 0 Then Kill(pathPDF)

        Dim htmlContent As String = html.ToString()

        Dim generator = New NReco.PdfGenerator.HtmlToPdfConverter()

        generator.PageFooterHtml = "<div style=""width: 100%; text-align: center; padding-top: 10px; display: inline-block;""><img src=""http://localhost:4586/Images/logo_ngs_rpt.jpg"" width=""60"" height=""20""></div>"

        Dim pdf = generator.GeneratePdf(htmlContent)

        Using fs As FileStream = New FileStream(pathPDF, FileMode.Create)
            fs.Write(pdf, 0, pdf.Length)
            fs.Close()
        End Using

        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & strNomeArquivo & "');", True)

    End Sub

#End Region


    'Protected Sub gridDescontos_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridDescontos.RowDataBound
    '    Select Case e.Row.RowType
    '        Case DataControlRowType.DataRow
    '            Dim PxA As PesagemXAnalises = CType(e.Row.DataItem, PesagemXAnalises)
    '            If PxA.Analise.Opcao.Length > 0 Then
    '                CType(e.Row.FindControl("txtPercentual"), TextBox).Visible = False
    '                CType(e.Row.FindControl("txtIndice"), TextBox).Visible = False
    '                CType(e.Row.FindControl("txtDesconto"), TextBox).Visible = False

    '                RemovedControl(CType(e.Row.FindControl("txtPercentual"), TextBox))
    '                RemovedControl(CType(e.Row.FindControl("txtIndice"), TextBox))
    '                RemovedControl(CType(e.Row.FindControl("txtDesconto"), TextBox))


    '                Dim ddlA As DropDownList = CType(e.Row.FindControl("ddlOpcao"), DropDownList)
    '                ddlA.Visible = True
    '                ddl.Carregar(ddlA, PxA.Analise.Opcao, ";", "-")

    '                ddlA.Enabled = False
    '                ddlA.SelectedValue = CInt(PxA.Percentual)
    '                e.Row.Cells(2).ColumnSpan = 3
    '            End If
    '    End Select
    'End Sub
End Class