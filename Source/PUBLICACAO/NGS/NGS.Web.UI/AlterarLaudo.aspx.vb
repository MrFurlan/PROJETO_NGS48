Imports System.Data
Imports System.Collections.Generic
Imports System.IO
Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio

Partial Class AlterarLaudo
    Inherits BasePage

    Dim Sql As String
    Private strJavaScript As String = ""

#Region "Session"
    Dim objLaudoOriginal As Pesagem
    Dim objLaudo As Pesagem
    Dim objNotaFiscal As NotaFiscal

    Private Sub SessaoRecuperaLaudoOriginal()
        objLaudoOriginal = CType(Session("objLaudoOriginal" & HID.Value), Pesagem)
    End Sub
    Private Sub SessaoSalvaLaudoOriginal()
        Session("objLaudoOriginal" & HID.Value) = objLaudoOriginal
    End Sub

    Private Sub SessaoRecuperaLaudo()
        objLaudo = CType(Session("objLaudo" & HID.Value), Pesagem)
        If objLaudo Is Nothing Then objLaudo = New Pesagem()
    End Sub
    Private Sub SessaoSalvaLaudo()
        Session("objLaudo" & HID.Value) = objLaudo
    End Sub

    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AlterarLaudo", "ACESSAR") Then
                ddl.Carregar(DdlClassificacao, CarregarDDL.Tabela.TabelaDeClassificacoes)
                ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                'AtribuirValoresCampos()
                Limpar()

                lnkConsultar.Focus()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    'Private Sub AtribuirValoresCampos()
    '    If Not Session("objLaudoFISICO" & HID.Value) Is Nothing Then
    '        Limpar()
    '        TxtLaudo.Text = Session("objLaudoFISICO" & HID.Value)
    '        ConsultarLaudo()
    '        Session.Remove("objLaudoFISICO" & HID.Value)
    '    End If

    '    If Not Session("objLaudoALTDEP" & HID.Value) Is Nothing Then
    '        Limpar()
    '        TxtLaudo.Text = Session("objLaudoALTDEP" & HID.Value)
    '        ConsultarLaudo()
    '        Session.Remove("objLaudoALTDEP" & HID.Value)
    '    End If
    'End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        SessaoRecuperaLaudo()
        If Not Session("objClienteAXT" & HID.Value) Is Nothing Then
            Dim cliente As Cliente = CType(Session("objClienteAXT" & HID.Value), [Lib].Negocio.Cliente)
            objLaudo.CodigoCliente = cliente.Codigo
            objLaudo.EnderecoCliente = cliente.CodigoEndereco
            objLaudo.Cliente = cliente

            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cliente)
            TxtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            TxtDepositante.Text = itemCliente.Text
            txtCodigoDepositante.Value = itemCliente.Value
            Session.Remove("objClienteAXT" & HID.Value)

            SessaoSalvaLaudo()

            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            ListarPedidos(Empresa(0), Empresa(1), "", "S")

        ElseIf Not Session("objClienteDEPAxL" & HID.Value) Is Nothing Then
            Dim Depositante As Cliente = CType(Session("objClienteDEPAxL" & HID.Value), [Lib].Negocio.Cliente)
            objLaudo.CodigoDepositante = Depositante.Codigo
            objLaudo.EnderecoDepositante = Depositante.CodigoEndereco
            objLaudo.Depositante = Depositante

            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(Depositante)
            txtCodigoDepositante.Value = itemCliente.Value
            TxtDepositante.Text = itemCliente.Text
            lnkAtualizar.Parent.Visible = True
            Session.Remove("objClienteDEPAxL" & HID.Value)

            SessaoSalvaLaudo()

        ElseIf Not Session("objClienteDEPOAxL" & HID.Value) Is Nothing Then
            Dim Deposito As Cliente = CType(Session("objClienteDEPOAxL" & HID.Value), [Lib].Negocio.Cliente)
            objLaudo.CodigoDeposito = Deposito.Codigo
            objLaudo.EnderecoDeposito = Deposito.CodigoEndereco
            objLaudo.Deposito = Deposito

            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(Deposito)
            txtCodigoDeposito.Value = itemCliente.Value
            TxtDeposito.Text = itemCliente.Text
            lnkAtualizar.Parent.Visible = True
            Session.Remove("objClienteDEPOAxL" & HID.Value)

            SessaoSalvaLaudo()

        ElseIf Not Session("objClienteTRAAxL" & HID.Value) Is Nothing Then
            Dim transportador As Cliente = CType(Session("objClienteTRAAxL" & HID.Value), [Lib].Negocio.Cliente)
            objLaudo.CodigoTransportador = transportador.Codigo
            objLaudo.EnderecoTransportador = transportador.CodigoEndereco
            objLaudo.Transportador = transportador

            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteTRAAxL" & HID.Value), [Lib].Negocio.Cliente))
            TxtTransportador.Text = itemCliente.Text
            txtCodigoTransportador.Value = itemCliente.Value
            Session.Remove("objClienteTRAAxL" & HID.Value)

            SessaoSalvaLaudo()

        ElseIf Not Session("objPlacaALTLDO" & HID.Value) Is Nothing Then
            If CType(Session("objPlacaALTLDO" & HID.Value), Placa).Restricao = "S" Then
                MsgBox(Me.Page, "Placa com restrição não pode ser utilizada. Observação: " & CType(Session("objPlacaALTLDO" & HID.Value), Placa).Observacao)
                Session.Remove("objPlacaALTLDO" & HID.Value)
                ConsultarPlaca()
            Else
                txtCodigoPlaca.Value = Trim(CType(Session("objPlacaALTLDO" & HID.Value), Placa).Placa01.ToUpper)
                txtPlaca.Text = txtCodigoPlaca.Value
                objLaudo.CodigoPlaca = txtCodigoPlaca.Value
                objLaudo.CodigoMotorista = CType(Session("objPlacaALTLDO" & HID.Value), Placa).CpfMotorista
                objLaudo.EnderecoMotorista = CType(Session("objPlacaALTLDO" & HID.Value), Placa).EndCpfMotorista
                Session.Remove("objPlacaALTLDO" & HID.Value)

                SessaoSalvaLaudo()
            End If
        ElseIf Not Session("objAlterarLaudo" & HID.Value) Is Nothing Then
            Dim objPedido As Pedido = CType(Session("objAlterarLaudo" & HID.Value), Pedido)

            ddlUnidadeDeNegocio.SelectedValue = objPedido.CodigoUnidadeNegocio
            CargaEmpresas()
            ddlEmpresa.SelectedValue = objPedido.CodigoEmpresa & "-" & objPedido.EnderecoEmpresa
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
            TxtPedido.Text = objPedido.Codigo

            objLaudo.CodigoEmpresa = objPedido.CodigoEmpresa
            objLaudo.EnderecoEmpresa = objPedido.EnderecoEmpresa
            objLaudo.Empresa = objPedido.Empresa
            objLaudo.CodigoPedido = objPedido.Codigo
            objLaudo.Pedido = objPedido

            objLaudo.CodigoOperacao = objLaudo.Pedido.CodigoOperacao
            objLaudo.CodigoSubOperacao = objLaudo.Pedido.CodigoSubOperacao

            Dim Parametros As New Hashtable
            Parametros.Add("Empresa", objLaudo.CodigoEmpresa)
            Parametros.Add("EndEmpresa", objLaudo.EnderecoEmpresa)
            Parametros.Add("Pedido", objLaudo.CodigoPedido)
            Parametros.Add("Operacao", objLaudo.Pedido.CodigoOperacao)
            Parametros.Add("SubOperacao", objLaudo.Pedido.CodigoSubOperacao)

            DdlOperacao.Items.Clear()
            ddl.Carregar(DdlOperacao, CarregarDDL.Tabela.OperacaoSubOperacaoPermitidasNaNota, "", False, Parametros)
            DdlOperacao.SelectedValue = objLaudo.CodigoOperacao & "-" & objLaudo.CodigoSubOperacao


            objLaudo.CodigoTabelaDeClassificacao = objPedido.Itens(0).CodigoClassificacao
            DdlClassificacao.SelectedValue = objLaudo.CodigoTabelaDeClassificacao

            TxtProduto.Text = objPedido.Itens(0).Produto.Descricao
            objLaudo.CodigoProduto = objPedido.Itens(0).CodigoProduto

            objLaudo.Analises = New [Lib].Negocio.ListPesagemXAnalises(objLaudo)

            gridDescontos.DataSource = objLaudo.Analises
            gridDescontos.DataBind()

            For Each row As GridViewRow In gridDescontos.Rows
                CType(row.FindControl("txtPercentual"), TextBox).Enabled = Not (objLaudo.TemNota OrElse objLaudo.Romaneios.Count > 1)
                CType(row.FindControl("txtIndice"), TextBox).Enabled = False
                CType(row.FindControl("txtDesconto"), TextBox).Enabled = False
            Next

            SessaoSalvaLaudo()

            If txtEntradaSaida.Value = "S" AndAlso txtCodigoAutorizacao.Value = 0 Then
                BuscaAutorizacao(objPedido.Codigo, objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa)
            End If
            Session.Remove("objAlterarLaudo" & HID.Value)
        End If
    End Sub

    Public Sub CarregarAutorizacao(Par As Hashtable)
        If Not Session("objAutorizacaoPesagem" & HID.Value) Is Nothing Then

            SessaoRecuperaLaudo()
            Dim strEmpresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")

            Dim objAutorizacao As New [Lib].Negocio.AutorizacaoDeRetirada(strEmpresa(0), strEmpresa(1), Par("Pedido"), Par("Autorizacao"), True)

            If objAutorizacao.CodigoPedido = CInt(TxtPedido.Text) Then
                Session("AutorizacaoLaudo" & HID.Value) = objAutorizacao

                If txtCodigoAutorizacao.Value = objAutorizacao.Autorizacao Then
                    txtSaldoAutorizacao.Value = objAutorizacao.QuantidadeAutorizadaFisica - (objAutorizacao.QuantidadeEntregueFisica - CDbl(txtLiquido.Text))
                Else
                    txtSaldoAutorizacao.Value = objAutorizacao.QuantidadeAutorizadaFisica - objAutorizacao.QuantidadeEntregueFisica
                End If

                txtAutorizacao.Text = objAutorizacao.Autorizacao
                txtCodigoAutorizacao.Value = objAutorizacao.Autorizacao

                If txtSaldoAutorizacao.Value < CDbl(txtLiquido.Text) Then
                    txtAutorizacao.Text = "0"
                    txtCodigoAutorizacao.Value = 0
                    txtSaldoAutorizacao.Value = 0
                    MsgBox(Me.Page, "Saldo fisico da Autorizacao Insuficiente, Saldo: " & objAutorizacao.SaldoFisico & " Laudo: " & txtLiquido.Text)
                End If
            Else
                MsgBox(Me.Page, "Pedido da Autorização " & objAutorizacao.CodigoPedido.ToString & " não pode ser diferente do Laudo " & TxtPedido.Text)
                BuscaAutorizacao(TxtPedido.Text, strEmpresa(0), strEmpresa(1))
            End If

            objLaudo.CodigoAutorizacao = txtCodigoAutorizacao.Value

            SessaoSalvaLaudo()

            Session.Remove("objAutorizacaoPesagem" & HID.Value)
        End If
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Protected Sub BtnConsultarCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            TxtPedido.Text = ""
            DdlOperacao.Items.Clear()
            TxtDepositante.Text = ""
            Session("ssCampo" & HID.Value) = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteAXT" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnConsultarPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnConsultarPedido.Click
        Try
            If ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "Selecione um cliente.")
            Else
                Session.Remove("AutorizacaoLaudo" & HID.Value)
                Session.Remove("objAutorizacaoPesagem" & HID.Value)
                txtAutorizacao.Text = "0"
                txtCodigoAutorizacao.Value = 0

                Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                ListarPedidos(Empresa(0), Empresa(1), "", "S")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnConsultarDepositante_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo" & HID.Value) = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteDEPAxL" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnConsultarDeposito_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo" & HID.Value) = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteDEPOAxL" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnConsultarTransportador_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo" & HID.Value) = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteTRAAxL" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdCalcular_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCalcular.Click
        Try
            If String.IsNullOrWhiteSpace(DdlOperacao.SelectedValue) Then
                MsgBox(Me.Page, "Operação não foi selecionada.")
            Else
                SessaoRecuperaLaudo()

                If String.IsNullOrWhiteSpace(txtPrimeiraPesagem.Text) Then
                    objLaudo.PrimeiraPesagem = 0
                Else
                    objLaudo.PrimeiraPesagem = CInt(txtPrimeiraPesagem.Text)
                End If

                If String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) Then
                    objLaudo.SegundaPesagem = 0
                Else
                    objLaudo.SegundaPesagem = CInt(txtSegundaPesagem.Text)
                End If

                If objLaudo.EntradaSaida = Left(eEntradaSaida.Saida.ToString, 1) Then
                    'txtPesoBruto.Text = Math.Abs(CInt(txtPrimeiraPesagem.Text) - CInt(txtSegundaPesagem.Text))
                    'txtLiquido.Text = Math.Abs(CInt(txtPrimeiraPesagem.Text) - CInt(txtSegundaPesagem.Text))
                    objLaudo.BrutoBalanca = Math.Abs(objLaudo.PrimeiraPesagem - objLaudo.SegundaPesagem)
                    objLaudo.Liquido = Math.Abs(objLaudo.PrimeiraPesagem - objLaudo.SegundaPesagem)

                    txtPesoBruto.Text = objLaudo.BrutoBalanca
                    txtDesconto.Text = objLaudo.Desconto
                    txtLiquido.Text = objLaudo.Liquido

                    For i = 0 To gridDescontos.Rows.Count - 1
                        Dim codanalise As Integer = gridDescontos.Rows(i).Cells(0).Text
                        Dim Analis As PesagemXAnalises = objLaudo.Analises.Where(Function(s) s.CodigoAnalise = codanalise).First

                        If Analis.Analise.Opcao.Length = 0 Then
                            Dim Percentual As String = CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text
                            Analis.Percentual = IIf(IsNumeric(Percentual), Percentual, 0)
                        Else
                            Dim ddlop As DropDownList = CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList)
                            Analis.Percentual = ddlop.SelectedValue
                        End If
                    Next

                    lnkAtualizar.Parent.Visible = True
                Else
                    For i = 0 To gridDescontos.Rows.Count - 1
                        Dim codanalise As Integer = gridDescontos.Rows(i).Cells(0).Text
                        Dim Analis As PesagemXAnalises = objLaudo.Analises.Where(Function(s) s.CodigoAnalise = codanalise).First

                        If Analis.Analise.Opcao.Length = 0 Then
                            Dim Percentual As String = CType(gridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text
                            Analis.Percentual = IIf(IsNumeric(Percentual), Percentual, 0)
                        Else
                            Dim ddlop As DropDownList = CType(gridDescontos.Rows(i).FindControl("ddlOpcao"), DropDownList)
                            Analis.Percentual = ddlop.SelectedValue
                        End If
                    Next

                    Dim Erro As String = objLaudo.Analises.CalcularDescontos()

                    If (String.IsNullOrEmpty(Erro)) Then
                        SessaoSalvaLaudo()

                        gridDescontos.DataSource = objLaudo.Analises
                        gridDescontos.DataBind()

                        txtPesoBruto.Text = objLaudo.BrutoBalanca
                        txtDesconto.Text = objLaudo.Desconto
                        txtLiquido.Text = objLaudo.Liquido

                        lnkAtualizar.Parent.Visible = True
                    Else
                        MsgBox(Me.Page, Erro)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        ddlUnidadeDeNegocio.Enabled = True
        ddlEmpresa.Enabled = True
        TxtLaudo.Enabled = True
        txtCodigoEmpresa.Value = ""
        TxtLaudo.Text = ""
        TxtRomaneio.Text = "0"
        txtEntradaSaida.Value = ""
        TxtMovimento.Text = ""
        TxtCliente.Text = ""
        txtCodigoCliente.Value = ""
        TxtPedido.Text = ""
        DdlOperacao.Items.Clear()
        txtAutorizacao.Text = "0"
        txtCodigoAutorizacao.Value = 0
        txtSaldoAutorizacao.Value = 0
        TxtDepositante.Text = ""
        txtCodigoDepositante.Value = ""
        TxtDeposito.Text = ""
        txtCodigoDeposito.Value = ""
        TxtTransportador.Text = ""
        txtCodigoTransportador.Value = ""
        TxtProduto.Text = ""
        txtPrimeiraPesagem.Text = ""
        txtSegundaPesagem.Text = ""
        txtPesoBruto.Text = ""
        txtDesconto.Text = ""
        txtLiquido.Text = ""
        txtCodigoPlaca.Value = ""
        txtPlaca.Text = ""
        txtNotaFiscal.Text = ""

        DdlOperacao.Enabled = False
        BtnConsultarCliente.Enabled = False
        BtnConsultarPedido.Enabled = False
        btnAutorizacao.Enabled = False
        btnConsultarNota.Enabled = False
        BtnConsultarDepositante.Enabled = False
        BtnConsultarDeposito.Enabled = False
        BtnConsultarTransportador.Enabled = False

        lnkConsultar.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkVincularNF.Parent.Visible = False
        lnkDesVincularNF.Parent.Visible = False
        cmdCalcular.Enabled = False
        txtPrimeiraPesagem.Enabled = True
        txtSegundaPesagem.Enabled = True
        TxtMovimento.Enabled = True
        DdlClassificacao.Enabled = True

        Session.Remove("AutorizacaoLaudo" & HID.Value)
        Session.Remove("objAutorizacaoPesagem" & HID.Value)

        Session.Remove("Produto" & HID.Value)
        Session.Remove("ProdutoAlterado" & HID.Value)

        Session.Remove("ssCalcular" & HID.Value)
        Session.Remove("ssPesagem" & HID.Value)
        Session.Remove("ssRomaneios" & HID.Value)
        Session.Remove("ssRomaneiosXDescontos" & HID.Value)
        Session.Remove("ssRomaneiosXPesagens" & HID.Value)
        Session.Remove("ssPedido" & HID.Value)
        Session.Remove("objPlacaALTLDO" & HID.Value)

        Session.Remove("ssPesagemXAnalises" & HID.Value)

        Session.Remove("objNFConsultaALTL" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        ucConsultaPlacas.SetarHID(HID.Value)
        ucConsultaAutorizacaoDeRetirada.SetarHID(HID.Value)
        ucConsultaCodMunicipios.SetarHID(HID.Value)
        ucConsultaPedidosXNotas.SetarHID(HID.Value)

        Session("ssCalcular" & HID.Value) = "N"
        Session("ssPesagem" & HID.Value) = ""
        Session("ssRomaneios" & HID.Value) = ""
        Session("ssRomaneiosXDescontos" & HID.Value) = ""
        Session("ssRomaneiosXPesagens" & HID.Value) = ""
        Session("ssPedido" & HID.Value) = ""

        gridDescontos.DataSource = CType(Session("ssPesagemXAnalises" & HID.Value), DataTable)
        gridDescontos.DataBind()

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub ConsultarPlaca()
        If txtCodigoTransportador.Value.Length > 0 Then
            Session("ssCampo" & HID.Value) = "LivreClasse"
            ucConsultaPlacas.Limpar()
            Popup.ConsultaDePlacas(Me.Page, "objPlacaALTLDO" & HID.Value)
            Popup.CenterDialog(Me.Page, "divConsultaPlacas")
            'ucConsultaPlacas.Limpar()
        Else
            MsgBox(Me.Page, "Transportador não foi selecionado!")
        End If
    End Sub


    Private Sub BuscaAutorizacao(ByVal Pedido As String, ByVal Empresa As String, ByVal EndEmpresa As String)
        SessaoRecuperaLaudo()

        Dim objPedido As [Lib].Negocio.Pedido = New [Lib].Negocio.Pedido(Empresa, EndEmpresa, Pedido)
        Dim parameters As New Dictionary(Of String, Object)
        parameters("ped") = Pedido
        parameters("emp") = Empresa
        parameters("endemp") = EndEmpresa
        parameters("cli") = objPedido.CodigoCliente
        parameters("endcli") = objPedido.EnderecoCliente
        parameters("romaneio") = True
        parameters("classe") = objPedido.SubOperacao.Classe
        Popup.ConsultaDeAutorizacaoDeRetirada(Me.Page, "objAutorizacaoPesagem" & HID.Value)
        ucConsultaAutorizacaoDeRetirada.BindGridView(parameters)
    End Sub

    Private Sub ListarPedidos(ByVal Empresa As String, ByVal EndEmpresa As String, ByVal Pedido As String, ByVal TipoOperacao As String)
        SessaoRecuperaLaudo()

        Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
        Session("ssCampo" & HID.Value) = "Pedidos"
        Session("ssPedido" & HID.Value) = TipoOperacao
        Dim parameters As New Dictionary(Of String, Object)
        parameters("unidade") = ddlUnidadeDeNegocio.SelectedValue
        parameters("empresa") = Empresa
        parameters("enderecoEmpresa") = EndEmpresa
        parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
        parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
        Popup.ConsultaDePedidos(Me.Page, "objAlterarLaudo" & HID.Value, "txtNome")
        ucConsultaPedidos.BindGridView(parameters)
    End Sub

    Private Sub ConsultarLaudo()
        Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
        '*************************************************************************
        '********************  Laudo de Pesagem  *********************************
        '*************************************************************************
        objLaudo = New Pesagem(Empresa(0), Empresa(1), TxtLaudo.Text)
        objLaudoOriginal = New Pesagem(Empresa(0), Empresa(1), TxtLaudo.Text)

        SessaoSalvaLaudo()
        SessaoSalvaLaudoOriginal()

        If objLaudo.Codigo = 0 Then
            MsgBox(Me.Page, "Laudo de Pesagem não encontrado!")
            TxtLaudo.Focus()
        Else
            If objLaudo.CodigoSituacao <> 1 Then
                MsgBox(Me.Page, "Laudo de Pesagem Cancelado não pode ser alterado!")
                TxtLaudo.Focus()
            ElseIf objLaudo.SegundaPesagem = 0 Then
                MsgBox(Me.Page, "Laudo sem a segunda pesagem deve ser ajustado no programa da Balança!")
                TxtLaudo.Focus()
                'ElseIf objLaudo.Observacoes.Contains("AGRUPAMENTO") Then
                '    MsgBox(Me.Page, "Laudo agrupado não pode ser alterado!")
                '    TxtLaudo.Focus()
            Else
                txtCodigoEmpresa.Value = objLaudo.CodigoEmpresa & "-" & objLaudo.EnderecoEmpresa

                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objLaudo.Cliente)
                TxtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value

                Dim itemDepositante As ListItem = Funcoes.FormatarListItemCliente(objLaudo.Depositante)
                TxtDepositante.Text = itemDepositante.Text
                txtCodigoDepositante.Value = itemDepositante.Value

                Dim itemDeposito As ListItem = Funcoes.FormatarListItemCliente(objLaudo.Deposito)
                TxtDeposito.Text = itemDeposito.Text
                txtCodigoDeposito.Value = itemDeposito.Value

                Dim itemTransportador As ListItem = Funcoes.FormatarListItemCliente(objLaudo.Transportador)
                TxtTransportador.Text = itemTransportador.Text
                txtCodigoTransportador.Value = itemTransportador.Value

                txtCodigoPlaca.Value = objLaudo.CodigoPlaca.ToUpper
                txtPlaca.Text = objLaudo.CodigoPlaca.ToUpper

                TxtMovimento.Text = objLaudo.Movimento
                TxtPedido.Text = objLaudo.CodigoPedido

                If objLaudo.Romaneios.Count = 1 Then
                    TxtRomaneio.Text = objLaudo.Romaneios(0).Codigo
                    objLaudo.CodigoAutorizacao = objLaudo.Romaneios(0).CodigoAutorizacao
                    txtCodigoAutorizacao.Value = objLaudo.CodigoAutorizacao
                    txtAutorizacao.Text = objLaudo.CodigoAutorizacao
                End If

                ddlUnidadeDeNegocio.SelectedValue = objLaudo.Pedido.CodigoUnidadeNegocio
                CargaEmpresas()

                ddlEmpresa.SelectedValue = objLaudo.Pedido.CodigoEmpresa & "-" & objLaudo.Pedido.EnderecoEmpresa
                ddlUnidadeDeNegocio.Enabled = False
                ddlEmpresa.Enabled = False
                TxtProduto.Text = objLaudo.Pedido.Itens(0).Produto.Descricao

                txtEntradaSaida.Value = Left(objLaudo.EntradaSaida.ToString, 1)

                Dim Parametros As New Hashtable
                Parametros.Add("Empresa", objLaudo.CodigoEmpresa)
                Parametros.Add("EndEmpresa", objLaudo.EnderecoEmpresa)
                Parametros.Add("Pedido", objLaudo.CodigoPedido)
                Parametros.Add("Operacao", objLaudo.Pedido.CodigoOperacao)
                Parametros.Add("SubOperacao", objLaudo.Pedido.CodigoSubOperacao)

                ddl.Carregar(DdlOperacao, CarregarDDL.Tabela.OperacaoSubOperacaoPermitidasNaNota, "", False, Parametros)
                DdlOperacao.SelectedValue = objLaudo.CodigoOperacao & "-" & objLaudo.CodigoSubOperacao

                DdlClassificacao.SelectedValue = objLaudo.CodigoTabelaDeClassificacao
                txtPrimeiraPesagem.Text = objLaudo.PrimeiraPesagem
                txtSegundaPesagem.Text = objLaudo.SegundaPesagem
                txtPesoBruto.Text = objLaudo.BrutoBalanca
                txtDesconto.Text = objLaudo.Desconto
                txtLiquido.Text = objLaudo.Liquido

                gridDescontos.DataSource = objLaudo.Analises
                gridDescontos.DataBind()

                For Each row As GridViewRow In gridDescontos.Rows
                    CType(row.FindControl("txtPercentual"), TextBox).Enabled = Not (objLaudo.TemNota OrElse objLaudo.Romaneios.Count > 1)
                    CType(row.FindControl("txtIndice"), TextBox).Enabled = False
                    CType(row.FindControl("txtDesconto"), TextBox).Enabled = False
                Next

                TxtLaudo.Enabled = False
                DdlOperacao.Enabled = True
                BtnConsultarCliente.Enabled = True
                BtnConsultarPedido.Enabled = True
                btnAutorizacao.Enabled = True
                If txtCodigoAutorizacao.Value > 0 Then
                    btnAutorizacao.Enabled = True
                    Dim objAutorizacao As New [Lib].Negocio.AutorizacaoDeRetirada(Empresa(0), Empresa(1), objLaudo.Pedido.Codigo, txtCodigoAutorizacao.Value, objLaudo.Pedido.SubOperacao.Classe)
                    Session("AutorizacaoLaudo" & HID.Value) = objAutorizacao
                End If
                DdlOperacao.Enabled = True
                BtnConsultarDepositante.Enabled = True
                BtnConsultarDeposito.Enabled = True
                BtnConsultarTransportador.Enabled = True
                btnConsultarPlaca.Enabled = True
                DdlClassificacao.Enabled = True
                txtPrimeiraPesagem.Enabled = True
                txtSegundaPesagem.Enabled = True
                cmdCalcular.Enabled = True
                lnkConsultar.Parent.Visible = False

                If objLaudo.TemNota OrElse objLaudo.Romaneios.Count > 1 Then
                    lnkAtualizar.Parent.Visible = False
                    DdlOperacao.Enabled = False
                    txtPrimeiraPesagem.Enabled = False
                    txtSegundaPesagem.Enabled = False
                    TxtMovimento.Enabled = False
                    cmdCalcular.Enabled = False
                    DdlClassificacao.Enabled = False
                    BtnConsultarCliente.Enabled = False
                    BtnConsultarPedido.Enabled = False
                    btnAutorizacao.Enabled = False
                    BtnConsultarDepositante.Enabled = False
                    BtnConsultarDeposito.Enabled = False

                    BtnConsultarTransportador.Enabled = False
                    btnConsultarPlaca.Enabled = False
                    If objLaudo.Romaneios.Count > 1 Then
                        MsgBox(Me.Page, "Laudo de Pesagem com rateio não pode ser alterado!")
                    ElseIf objLaudo.Observacoes.Contains("AGRUPAMENTO") Then
                        MsgBox(Me.Page, "Laudo agrupado com Nota Fiscal não pode ser alterado!")
                        Limpar()
                        TxtLaudo.Focus()
                    ElseIf objLaudo.TemNota Then

                        Dim sql As String = "select nXr.Empresa_Id, nXr.EndEmpresa_Id, nXr.Cliente_Id, nXr.EndCliente_Id, nXr.EntradaSaida_Id, n.Serie_Id, nXr.Nota_Id" & vbCrLf & _
                                           "from notasfiscaisXromaneios nXr" & vbCrLf & _
                                           "		left join NotasFiscais n" & vbCrLf & _
                                           "				ON n.Empresa_Id       = nXr.Empresa_Id" & vbCrLf & _
                                           "				and n.EndEmpresa_iD   = nXr.EndEmpresa_Id" & vbCrLf & _
                                           "				and n.Cliente_Id      = nXr.Cliente_Id" & vbCrLf & _
                                           "				and n.EndCliente_Id   = nXr.EndCliente_Id" & vbCrLf & _
                                           "				and n.EntradaSaida_Id = nXr.EntradaSaida_Id" & vbCrLf & _
                                           "				and n.Serie_Id        = nXr.Serie_Id" & vbCrLf & _
                                           "				and n.Nota_Id         = nXr.Nota_Id" & vbCrLf & _
                                           "where nXr.Empresa_Id  = '" & objLaudo.CodigoEmpresa & "'" & vbCrLf & _
                                           "and nXr.EndEmpresa_Id = " & objLaudo.EnderecoEmpresa & vbCrLf & _
                                           "and nXr.Cliente_Id    = '" & objLaudo.CodigoCliente & "'" & vbCrLf & _
                                           "and nXr.EndCliente_Id = " & objLaudo.EnderecoCliente & vbCrLf & _
                                           "and nXr.Romaneio_Id   = " & objLaudo.Romaneios(0).Codigo

                        Dim objBanco As New AcessaBanco()
                        Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "NotaFiscalXRomaneio")

                        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count = 1 Then

                            objNotaFiscal = New [Lib].Negocio.NotaFiscal()

                            objNotaFiscal.CodigoEmpresa = ds.Tables(0).Rows(0).Item(0)
                            objNotaFiscal.EnderecoEmpresa = ds.Tables(0).Rows(0).Item(1)
                            objNotaFiscal.CodigoCliente = ds.Tables(0).Rows(0).Item(2)
                            objNotaFiscal.EnderecoCliente = ds.Tables(0).Rows(0).Item(3)
                            objNotaFiscal.EntradaSaida = IIf(ds.Tables(0).Rows(0).Item(4) = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                            objNotaFiscal.Serie = ds.Tables(0).Rows(0).Item(5)
                            objNotaFiscal.Codigo = ds.Tables(0).Rows(0).Item(6)

                            objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNotaFiscal)

                            txtNotaFiscal.Text = objNotaFiscal.Codigo

                            SalvaNotaFiscal()

                            If Funcoes.VerificaPermissao("DesVincularLaudo", "ALTERAR") Then lnkDesVincularNF.Parent.Visible = True

                        End If

                        MsgBox(Me.Page, "Laudo de Pesagem vinculado com Nota Fiscal não pode ser alterado, apenas ajustado o Depósito/Depositante!")
                        If Not objLaudo.Analises Is Nothing AndAlso objLaudo.Analises.Count > 0 Then
                            lnkAtualizar.Parent.Visible = False
                        Else
                            lnkAtualizar.Parent.Visible = True
                        End If
                        BtnConsultarDepositante.Enabled = True
                        BtnConsultarDeposito.Enabled = True
                    End If
                Else
                    If objLaudo.Observacoes.Contains("AGRUPAMENTO") Then
                        MsgBox(Me.Page, "Laudo agrupado não pode ser alterado, apenas vinculado com Nota Fiscal!")

                        lnkAtualizar.Parent.Visible = False
                        DdlOperacao.Enabled = False
                        txtPrimeiraPesagem.Enabled = False
                        txtSegundaPesagem.Enabled = False
                        TxtMovimento.Enabled = False
                        cmdCalcular.Enabled = False
                        DdlClassificacao.Enabled = False
                        BtnConsultarCliente.Enabled = False
                        BtnConsultarPedido.Enabled = False
                        btnAutorizacao.Enabled = False
                        BtnConsultarDepositante.Enabled = False
                        BtnConsultarDeposito.Enabled = False

                        BtnConsultarTransportador.Enabled = False
                        btnConsultarPlaca.Enabled = False

                        btnConsultarNota.Enabled = True
                    ElseIf Not objLaudo.Analises Is Nothing AndAlso objLaudo.Analises.Count > 0 Then
                        lnkAtualizar.Parent.Visible = False
                    Else
                        lnkAtualizar.Parent.Visible = True
                    End If

                    btnConsultarNota.Enabled = True
                End If
            End If
        End If
    End Sub

    Public Sub CarregarNotaFiscal()
        Try
            If Not Session("objNFConsultaALTL" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaALTL" & HID.Value), NotaFiscal))

                If objNotaFiscal.Codigo > 0 Then
                    If Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento.ToString("dd-MM-yyyy"), "NOTAS FISCAIS") Then
                        MsgBox(Me.Page, "Movimento de Notas Fiscais já Fechado para esta data...")
                    ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento.ToString("dd-MM-yyyy"), "PRODUCAO") Then
                        MsgBox(Me.Page, "Movimento de Produção já Fechado para esta data....")
                    ElseIf objNotaFiscal.CodigoRomaneio > 0 AndAlso Not objNotaFiscal.Romaneio.Processo = "NOTA FISCAL" Then
                        MsgBox(Me.Page, "Nota Fiscal não pode ser utilizada pois já está vinculada com uma Pesagem.")
                    Else
                        SessaoRecuperaLaudo()

                        If objLaudo.CodigoOperacao = objNotaFiscal.CodigoOperacao AndAlso objLaudo.CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao Then
                            txtNotaFiscal.Text = objNotaFiscal.Codigo

                            SalvaNotaFiscal()

                            Session.Remove("objNFConsultaALTL" & HID.Value)

                            lnkVincularNF.Parent.Visible = True
                        Else
                            MsgBox(Me.Page, "Operação da Nota Fiscal é diferente da Operação do Laudo, primeiro ajuste a operação do Laudo e depois refaça o processo para vincular!")
                        End If
                    End If
                Else
                    MsgBox(Me.Page, "Houve um problema ao selecionar a Nota Fiscal, refaça o processo!")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidarLaudo() As Boolean
        SessaoRecuperaLaudo()

        'Nao deixar alterar um laudo de entrada pra saida e vice versa
        If objLaudo.EntradaSaida <> IIf(CInt(txtPrimeiraPesagem.Text) > CInt(txtSegundaPesagem.Text), "E", "S") Then
            MsgBox(Me.Page, "O Laudo nao pode ser alterado de Entrada pra saida e vice versa.")
            Return False
        End If

        If objLaudo.Produto.Agrupar = "N" AndAlso CInt(txtSegundaPesagem.Text) > CInt(txtPrimeiraPesagem.Text) AndAlso txtCodigoAutorizacao.Value = 0 Then
            MsgBox(Me.Page, "Autorização de Retirada não foi selecionada.")
            Return False
            'ElseIf CInt(txtSegundaPesagem.Text) > CInt(txtPrimeiraPesagem.Text) AndAlso Session("AutorizacaoLaudo" & HID.Value) Is Nothing Then
            '    MsgBox(Me.Page, "Autorização de Retirada não foi encontrada, selecione novamente.")
            '    Return False
            'ElseIf CInt(txtSegundaPesagem.Text) > CInt(txtPrimeiraPesagem.Text) AndAlso _
            '       Not CType(Session("AutorizacaoLaudo" & HID.Value), [Lib].Negocio.AutorizacaoDeRetirada).CodigoEmpresa = objLaudo.CodigoEmpresa Then
            '    MsgBox(Me.Page, "Empresa da Pesagem não é a mesma da Autorização de Retirada.")
            '    Return False
            'ElseIf CInt(txtSegundaPesagem.Text) > CInt(txtPrimeiraPesagem.Text) AndAlso _
            '       Not CType(Session("AutorizacaoLaudo" & HID.Value), [Lib].Negocio.AutorizacaoDeRetirada).CodigoPedido = CInt(TxtPedido.Text) Then
            '    MsgBox(Me.Page, "Pedido da Pesagem não é o mesmo da Autorização de Retirada.")
            '    Return False
            'ElseIf txtCodigoAutorizacao.Value > 0 AndAlso txtSaldoAutorizacao.Value < CDbl(txtLiquido.Text) Then
            '    MsgBox(Me.Page, "Saldo fisico da Autorizacao Insuficiente, Saldo: " & txtSaldoAutorizacao.Value & " Laudo: " & txtLiquido.Text)
            '    Return False
        ElseIf objLaudo.TemNota Then
            For Each rom In objLaudo.Romaneios
                If rom.NF.Itens(0).CodigoProduto <> objLaudo.CodigoProduto Then
                    MsgBox(Me.Page, "Produto da Nota Fiscal está diferente do alterado no Laudo.")
                    Return False
                End If
            Next
            Return True
            'If Not nf Is Nothing AndAlso nf.Itens.Count > 0 AndAlso Not nf.Itens(0).CodigoProduto = objLaudo.CodigoProduto Then
            '    MsgBox(Me.Page, "Produto da Nota Fiscal está diferente do alterado no Laudo.")
            '    Return False
            'ElseIf Not nf Is Nothing AndAlso nf.Itens.Count = 0 Then
            '    MsgBox(Me.Page, "Laudo com Nota Fiscal não pode ser alterado. Entre em contato com o Suporte.")
            '    Return False
            'Else
            '    Return True
            'End If
        Else
            Return True
        End If
    End Function

    Protected Sub DdlClassificacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaLaudo()

            objLaudo.CodigoTabelaDeClassificacao = DdlClassificacao.SelectedValue

            gridDescontos.DataSource = objLaudo.Analises
            gridDescontos.DataBind()

            SessaoSalvaLaudo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultarNota_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaLaudo()

            objNotaFiscal = New [Lib].Negocio.NotaFiscal()
            objNotaFiscal.CodigoEmpresa = objLaudo.CodigoEmpresa
            objNotaFiscal.EntradaSaida = IIf(objLaudo.EntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
            objNotaFiscal.CodigoSituacao = 1
            objNotaFiscal.CodigoTipoDeDocumento = 1
            objNotaFiscal.CodigoPedido = objLaudo.CodigoPedido
            objNotaFiscal.DataNota = objNotaFiscal.DataNota.AddYears(-2)

            Session("objNotaFiscal" & HID.Value) = objNotaFiscal
            Session("ssCampo" & HID.Value) = "AlterarLaudo"

            ucConsultaPedidosXNotas.BindGridView()
            Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaALTL" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultarPlaca_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ConsultarPlaca()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnAutorizacao_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim strEmpresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            If String.IsNullOrWhiteSpace(TxtPedido.Text) OrElse TxtPedido.Text = 0 Then
                BtnConsultarPedido_Click(BtnConsultarPedido, New EventArgs())
            End If

            BuscaAutorizacao(TxtPedido.Text, strEmpresa(0), strEmpresa(1))
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarLaudo", "LEITURA") Then
                If TxtLaudo.Text.Length = 0 Then
                    MsgBox(Me.Page, "Número da pesagem não foi informado.")
                ElseIf ddlUnidadeDeNegocio.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Unidade de Negócio não foi selecionada.")
                ElseIf ddlEmpresa.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Empresa não foi selecionada.")
                Else
                    ConsultarLaudo()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            SessaoRecuperaLaudoOriginal()
            SessaoRecuperaLaudo()

            If Funcoes.VerificaPermissao("AlterarLaudo", "ALTERAR") Then
                If objLaudo.Codigo > 0 Then
                    If ValidarLaudo() Then
                        Dim sqls As New ArrayList

                        If objLaudoOriginal.CodigoDeposito <> objLaudo.CodigoDeposito Or objLaudoOriginal.EnderecoDeposito <> objLaudo.EnderecoDeposito Then
                            For Each rom In objLaudo.Romaneios
                                If rom.NF IsNot Nothing AndAlso rom.NF.Codigo > 0 Then
                                    Sql = "UPDATE NotasFiscais SET" & _
                                          "    Deposito    ='" & objLaudo.CodigoDeposito & "'" & _
                                          "   ,EndDeposito = " & objLaudo.EnderecoDeposito & _
                                          " WHERE Empresa_id      ='" & rom.NF.CodigoEmpresa & "'" & _
                                          "   AND EndEmpresa_id   = " & rom.NF.EnderecoEmpresa &
                                          "   AND Cliente_Id      ='" & rom.NF.CodigoCliente & "'" & _
                                          "   AND EndCliente_Id   = " & rom.NF.EnderecoCliente & _
                                          "   AND EntradaSaida_Id ='" & Left(rom.NF.EntradaSaida.ToString, 1) & "'" & _
                                          "   AND Serie_Id        ='" & rom.NF.Serie & "'" & _
                                          "   AND Nota_Id         = " & rom.NF.Codigo & ";"
                                    sqls.Add(Sql)


                                    Sql = "UPDATE NotasFiscaisXItens SET" & _
                                          "    QuantidadeFisica = " & CInt(txtLiquido.Text) & _
                                          "   ,Deposito         ='" & objLaudo.CodigoDeposito & "'" & _
                                          "   ,EndDeposito      = " & objLaudo.EnderecoDeposito & _
                                          " WHERE Empresa_id      ='" & rom.NF.CodigoEmpresa & "'" & _
                                          "   AND EndEmpresa_id   = " & rom.NF.EnderecoEmpresa &
                                          "   AND Cliente_Id      ='" & rom.NF.CodigoCliente & "'" & _
                                          "   AND EndCliente_Id   = " & rom.NF.EnderecoCliente & _
                                          "   AND EntradaSaida_Id ='" & Left(rom.NF.EntradaSaida.ToString, 1) & "'" & _
                                          "   AND Serie_Id        ='" & rom.NF.Serie & "'" & _
                                          "   AND Nota_Id         = " & rom.NF.Codigo & ";"
                                    sqls.Add(Sql)
                                End If

                            Next
                        End If

                        objLaudo.UsuarioAlteracao = Session("ssNomeUsuario").ToString

                        If objLaudo.CodigoEmpresa <> objLaudoOriginal.CodigoEmpresa OrElse objLaudo.EnderecoEmpresa <> objLaudoOriginal.EnderecoEmpresa Then
                            objLaudoOriginal.IUD = "D"
                            objLaudoOriginal.SalvarSql(sqls)

                            objLaudo.IUD = "I"
                            objLaudo.Codigo = 0
                            objLaudo.CriarRomaneio = True
                            objLaudo.SalvarSql(sqls)
                        Else
                            objLaudo.IUD = "U"
                            objLaudo.SalvarSql(sqls)
                        End If

                        If Banco.GravaBanco(sqls) Then
                            MsgBox(Me.Page, "Laudo Alterado com Sucesso.")
                            Limpar()
                        Else
                            MsgBox(Me.Page, "Erro a Altererar o Laudo de Pesagem.")
                        End If
                    End If
                Else
                    MsgBox(Me.Page, "Consultar o Laudo para alteração!")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkVincularNF_Click(sender As Object, e As EventArgs) Handles lnkVincularNF.Click
        Try
            If Funcoes.VerificaPermissao("VincularLaudo", "ALTERAR") Then

                SessaoRecuperaLaudo()

                RecuperaNotaFiscal()

                Dim sqls As New ArrayList

                Dim romaneioAnt As Integer = objNotaFiscal.CodigoRomaneio

                Sql = "delete from notasfiscaisXromaneios " & vbCrLf & _
                        "where Empresa_Id  = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                        "and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                        "and Cliente_Id    = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                        "and EndCliente_Id = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                        "and Romaneio_Id   = " & objNotaFiscal.CodigoRomaneio & ";"
                sqls.Add(Sql)

                objNotaFiscal.CodigoRomaneio = Trim(TxtRomaneio.Text)

                Sql = " Insert into NotasFiscaisXRomaneios(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Romaneio_Id)" & vbCrLf &
                      " Values('" & objNotaFiscal.CodigoEmpresa & "'," & objNotaFiscal.EnderecoEmpresa & ",'" & objNotaFiscal.CodigoCliente & "'," & objNotaFiscal.EnderecoCliente & ",'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "','" & objNotaFiscal.Serie & "'," & objNotaFiscal.Codigo & "," & objNotaFiscal.CodigoRomaneio & ");"
                sqls.Add(Sql)

                Sql = "UPDATE NotasFiscaisXItens SET" & vbCrLf &
                        "    QuantidadeFisica   = " & CInt(txtLiquido.Text) & vbCrLf &
                        "   ,Deposito           ='" & objLaudo.CodigoDeposito & "'" & vbCrLf &
                        "   ,EndDeposito        = " & objLaudo.EnderecoDeposito & vbCrLf &
                        " WHERE Empresa_id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                        "   AND EndEmpresa_id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                        "   AND Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                        "   AND EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                        "   AND EntradaSaida_Id ='" & Left(objNotaFiscal.EntradaSaida.ToString, 1) & "'" & vbCrLf &
                        "   AND Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                        "   AND Nota_Id         = " & objNotaFiscal.Codigo & ";"
                sqls.Add(Sql)

                Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                    obs = obs & ". Vinculado Laudo pelo Alterar Laudo em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " - " & HttpContext.Current.Session("ssNomeUsuario")
                Else
                    obs = "Vinculado Laudo pelo Alterar Laudo em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " - " & HttpContext.Current.Session("ssNomeUsuario")
                End If

                Sql = " Update NotasFiscais set " & vbCrLf & _
                      "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                      "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf & _
                      "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf & _
                      "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                      "	   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                      "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                      "    and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                      "    and Nota_Id         = " & objNotaFiscal.Codigo & ";"
                sqls.Add(Sql)

                Sql = "DELETE FROM ROMANEIOSxDESCONTOS " & vbCrLf &
                      " WHERE Empresa_id    ='" & objLaudo.CodigoEmpresa & "'" & vbCrLf &
                      "   AND EndEmpresa_id = " & objLaudo.EnderecoEmpresa & vbCrLf &
                      "   AND ROMANEIO_ID   = " & romaneioAnt & ";"
                sqls.Add(Sql)

                Sql = "DELETE FROM ROMANEIOS " & vbCrLf &
                      " WHERE Empresa_id    ='" & objLaudo.CodigoEmpresa & "'" & vbCrLf &
                      "   AND EndEmpresa_id = " & objLaudo.EnderecoEmpresa & vbCrLf &
                      "   AND ROMANEIO_ID   = " & romaneioAnt & ";"
                sqls.Add(Sql)

                If Banco.GravaBanco(sqls) Then
                    MsgBox(Me.Page, "Laudo Vinculado com Sucesso.")
                    Limpar()
                Else
                    MsgBox(Me.Page, "Erro ao Vincular o Laudo de Pesagem.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para vincular registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkDesVincularNF_Click(sender As Object, e As EventArgs) Handles lnkDesVincularNF.Click

        Try
            SessaoRecuperaLaudo()

            RecuperaNotaFiscal()

            Dim sqls As New ArrayList

            Sql = "Delete from notasfiscaisXromaneios" & vbCrLf &
               "where Empresa_Id  = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
               "and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
               "and Cliente_Id    = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
               "and EndCliente_Id = " & objNotaFiscal.EnderecoCliente & vbCrLf &
               "and Romaneio_Id   = " & objNotaFiscal.CodigoRomaneio & ";"
            sqls.Add(Sql)

            Sql = "UPDATE NotasFiscaisXItens SET" & vbCrLf & _
                         "    QuantidadeFisica = 0" & vbCrLf & _
                         " WHERE Empresa_id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                         "   AND EndEmpresa_id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                         "   AND Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                         "   AND EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                         "   AND EntradaSaida_Id ='" & Left(objNotaFiscal.EntradaSaida.ToString, 1) & "'" & vbCrLf & _
                         "   AND Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                         "   AND Nota_Id         = " & objNotaFiscal.Codigo & ";"
            sqls.Add(Sql)

            Dim obs As String = objNotaFiscal.ObservacoesControleInterno
            If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                obs = obs & ". Desvinculado Laudo pelo Alterar Laudo em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " - " & HttpContext.Current.Session("ssNomeUsuario")
            Else
                obs = "Desvinculado Laudo pelo Alterar Laudo em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " - " & HttpContext.Current.Session("ssNomeUsuario")
            End If

            Sql = " Update NotasFiscais set " & vbCrLf & _
                  "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                  "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf & _
                  "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf & _
                  "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                  "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                  "	   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                  "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                  "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                  "    and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf & _
                  "    and Nota_Id         = " & objNotaFiscal.Codigo & ";"
            sqls.Add(Sql)

            If Banco.GravaBanco(sqls) Then
                MsgBox(Me.Page, "Laudo Desvinculado com Sucesso.")
                Limpar()
            Else
                MsgBox(Me.Page, "Erro ao Desvincular o Laudo de Pesagem.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AlterarLaudo")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub GridDescontos_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridDescontos.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim PxA As PesagemXAnalises = CType(e.Row.DataItem, PesagemXAnalises)
                If PxA.Analise.Opcao.Length > 0 Then
                    CType(e.Row.FindControl("txtPercentual"), TextBox).Visible = False
                    CType(e.Row.FindControl("txtIndice"), TextBox).Visible = False
                    CType(e.Row.FindControl("txtDesconto"), TextBox).Visible = False

                    Dim ddlA As DropDownList = CType(e.Row.FindControl("ddlOpcao"), DropDownList)
                    ddlA.Visible = True
                    ddl.Carregar(ddlA, PxA.Analise.Opcao, ";", "-")

                    ddlA.SelectedValue = CInt(PxA.Percentual)
                    e.Row.Cells(2).ColumnSpan = 3
                End If
        End Select
    End Sub

    Protected Sub DdlOperacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlOperacao.SelectedIndexChanged
        Dim so As String() = DdlOperacao.SelectedValue.ToString.Split("-")
        SessaoRecuperaLaudo()
        objLaudo.CodigoOperacao = so(0)
        objLaudo.CodigoSubOperacao = so(1)
        objLaudo.SubOperacao = Nothing
        SessaoSalvaLaudo()
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEmpresa.SelectedIndexChanged
        Dim emp As String() = ddlEmpresa.SelectedValue.ToString.Split("-")
        SessaoRecuperaLaudo()
        objLaudo.CodigoEmpresa = emp(0)
        objLaudo.EnderecoEmpresa = emp(1)
        SessaoSalvaLaudo()
    End Sub
End Class