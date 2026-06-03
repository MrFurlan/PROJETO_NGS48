Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AutorizacaoDeRetirada
    Inherits BasePage

    Dim objAutorizacao As [Lib].Negocio.AutorizacaoDeRetirada

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("AutorizacaoDeRetirada", "ACESSAR") Then
                    SessaoRecuperaAutorizacao()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub SessaoSalvaAutorizacao()
        Session("Autorizacao" & HID.Value) = objAutorizacao
    End Sub

    Private Sub SessaoRecuperaAutorizacao()
        objAutorizacao = CType(Session("Autorizacao" & HID.Value), [Lib].Negocio.AutorizacaoDeRetirada)
    End Sub

    Public Sub CarregarAutorizacaoAxE(par As Hashtable)
        Try
            If Not Session("objAutorizacaoAxE" & HID.Value) Is Nothing Then
                SessaoRecuperaAutorizacao()

                objAutorizacao = New [Lib].Negocio.AutorizacaoDeRetirada(objAutorizacao.CodigoEmpresa, objAutorizacao.EnderecoEmpresa, par("Pedido"), par("Autorizacao"), objAutorizacao.Pedido.SubOperacao.Classe)
                SessaoSalvaAutorizacao()
                AtualizaFormularioComAClasse()


                'Dim SaldoProduto As New [Lib].Negocio.ListSaldoPedidoXNota(objAutorizacao.CodigoEmpresa, objAutorizacao.EnderecoEmpresa, "", 0, "", "", "", objAutorizacao.CodigoPedido, [Lib].Negocio.ListSaldoPedidoXNota.Situacao.Todos, False)
                Dim Parametros As New Hashtable
                Parametros.Add("Empresa", objAutorizacao.CodigoEmpresa)
                Parametros.Add("EndEmpresa", objAutorizacao.EnderecoEmpresa)
                Parametros.Add("Pedido", objAutorizacao.CodigoPedido)
                Dim SaldoProdutoNovo As New ListSaldoPedido2015(Parametros)

                Dim Autorizacoes As New [Lib].Negocio.ListAutorizacaoDeRetirada(objAutorizacao.CodigoEmpresa, objAutorizacao.EnderecoEmpresa, "", 0, objAutorizacao.CodigoPedido, objAutorizacao.Pedido.SubOperacao.Classe)

                'txtMaxQuantidadeFiscal.Value = Math.Max(Autorizacoes.QtdeEntregueFiscal + SaldoProduto(0).SaldoFiscal - Autorizacoes.QtdeAutorizadaFiscal + objAutorizacao.QuantidadeAutorizadaFiscal, objAutorizacao.QuantidadeAutorizadaFiscal)
                'txtMaxQuantidadeFisica.Value = Math.Max(Autorizacoes.QtdeEntregueFisica + SaldoProduto(0).SaldoFisico - Autorizacoes.QtdeAutorizadaFisica + objAutorizacao.QuantidadeAutorizadaFisica, objAutorizacao.QuantidadeAutorizadaFisica)

                txtMaxQuantidadeFiscal.Value = (SaldoProdutoNovo(0).QtdeContratadoFiscal - Autorizacoes.QtdeAutorizadaFiscal) + objAutorizacao.QuantidadeAutorizadaFiscal
                txtMaxQuantidadeFisica.Value = (SaldoProdutoNovo(0).QtdeContratadoFisico - Autorizacoes.QtdeAutorizadaFisica) + objAutorizacao.QuantidadeAutorizadaFisica

                txtQtdeContratada.Text = SaldoProdutoNovo(0).QtdeContratadoFiscal.ToString("N0")
                txtQtdeContratadaFisica.Text = SaldoProdutoNovo(0).QtdeContratadoFisico.ToString("N0")

                txtQtdeEntregue.Text = (SaldoProdutoNovo(0).QtdeEntregueFiscalDireta + SaldoProdutoNovo(0).QtdeEntregueFiscalAFixar + SaldoProdutoNovo(0).QtdeEntregueFiscalGlobal + SaldoProdutoNovo(0).QtdeEntregueFiscalDeposito).ToString("N0")
                txtQtdeEntregueFisica.Text = (SaldoProdutoNovo(0).QtdeEntregueFisicoDireta + SaldoProdutoNovo(0).QtdeEntregueFisicoAFixar + SaldoProdutoNovo(0).QtdeEntregueFisicoRemessa + SaldoProdutoNovo(0).QtdeEntregueFisicoDeposito).ToString("N0")

                txtSaldoFiscal.Text = (SaldoProdutoNovo(0).QtdeContratadoFiscal - (SaldoProdutoNovo(0).QtdeEntregueFiscalDireta + SaldoProdutoNovo(0).QtdeEntregueFiscalAFixar + SaldoProdutoNovo(0).QtdeEntregueFiscalGlobal + SaldoProdutoNovo(0).QtdeEntregueFiscalDeposito)).ToString("N0")
                txtSaldoFisico.Text = (SaldoProdutoNovo(0).QtdeContratadoFisico - (SaldoProdutoNovo(0).QtdeEntregueFisicoDireta + SaldoProdutoNovo(0).QtdeEntregueFisicoAFixar + SaldoProdutoNovo(0).QtdeEntregueFisicoRemessa + SaldoProdutoNovo(0).QtdeEntregueFisicoDeposito)).ToString("N0")

                txtAutorizadoFiscal.Text = Autorizacoes.QtdeAutorizadaFiscal.ToString("N0")
                txtAutorizadoFisico.Text = Autorizacoes.QtdeAutorizadaFisica.ToString("N0")

                Session.Remove("objAutorizacaoAxE" & HID.Value)
                lnkNovo.Parent.Visible = False
                lnkExcluir.Parent.Visible = True
                lnkAtualizar.Parent.Visible = True
            End If

            If Not Session("SemAutorizacao" & HID.Value) Is Nothing Then
                If Not String.IsNullOrEmpty(txtCodigoCliente.Value.Trim()) AndAlso Not String.IsNullOrEmpty(txtPedido.Text.Trim()) Then
                    Dim ds As New DataSet
                    Dim strCliente = txtCodigoCliente.Value.Split("-")

                    Dim sql As String
                    sql = "Select count(autorizacao_id) as qtde " & vbCrLf & _
                          "  from autorizacaoderetirada " & vbCrLf & _
                          " where empresa_id    ='" & strCliente(0) & "'" & vbCrLf & _
                          "   and endempresa_id ='" & strCliente(1) & "'" & vbCrLf & _
                          "   and pedido_id     ='" & txtPedido.Text & "'"

                    ds = Banco.ConsultaDataSet(sql, "AutorizacaoDeRetirada")

                    If ds.Tables(0).Rows.Count > 0 AndAlso Convert.ToInt32(ds.Tables(0).Rows(0).Item("qtde")) > 0 Then
                        Limpar()
                        Exit Sub
                    End If
                End If

                Session.Remove("SemAutorizacao" & HID.Value)
                MsgBox(Me.Page, "Não há autorizações para os parâmetros selecionados!")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Not Session("objEmpresaAxE" & HID.Value) Is Nothing Then
                SetarEmpresa(CType(Session("objEmpresaAxE" & HID.Value), [Lib].Negocio.Cliente).Codigo, CType(Session("objEmpresaAxE" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco)
                Session.Remove("objEmpresaAxE" & HID.Value)
                'SessaoRecuperaAutorizacao()
                'Dim n As New [Lib].Negocio.Numerador(objAutorizacao.CodigoEmpresa, objAutorizacao.EnderecoEmpresa, 50)
                'objAutorizacao.Autorizacao = n.Sequencia + 1
                'txtAutorizacao.Text = n.Sequencia
                'SessaoSalvaAutorizacao()
            ElseIf Not Session("objClienteAxE" & HID.Value) Is Nothing Then
                Dim objClienteAxE As [Lib].Negocio.Cliente = CType(Session("objClienteAxE" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objClienteAxE)
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteAxE" & HID.Value)
                BuscarPedido()
            ElseIf Not Session("objClienteAxEA" & HID.Value) Is Nothing Then
                Dim objClienteAxEA As [Lib].Negocio.Cliente = CType(Session("objClienteAxEA" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objClienteAxEA)
                txtAutorizante.Text = itemCliente.Text
                txtCodigoAutorizante.Value = itemCliente.Value
                Session.Remove("objClienteAxEA" & HID.Value)
                SalvarAutorizante()
            ElseIf Not Session("objClienteAxECR" & HID.Value) Is Nothing Then
                Dim objClienteAxECR As [Lib].Negocio.Cliente = CType(Session("objClienteAxECR" & HID.Value), [Lib].Negocio.Cliente)
                SessaoRecuperaAutorizacao()
                Dim sqlAutorizacao As String = "select count(autorizacao_id) as qtde " & vbCrLf & _
                                               "   from autorizacaoderetirada " & vbCrLf & _
                                               "   where empresa_id = '" & objAutorizacao.CodigoEmpresa & "'   " & vbCrLf & _
                                               "   and endempresa_id = '" & objAutorizacao.EnderecoEmpresa & "' " & vbCrLf & _
                                               "   and pedido_id = '" & objAutorizacao.CodigoPedido & "' " & vbCrLf & _
                                               "   and clienteretirada = '" & objClienteAxECR.Codigo & "'" & vbCrLf & _
                                               "   and endclienteretirada = " & objClienteAxECR.CodigoEndereco
                Dim ds As New DataSet
                ds = Banco.ConsultaDataSet(sqlAutorizacao, "AutorizacaoDeRetirada")

                If ds.Tables IsNot Nothing AndAlso Convert.ToInt32(ds.Tables(0).Rows(0).Item("qtde")) > 0 Then
                    MsgBox(Me.Page, "Ja existe autorização para este destinatário, selecione-a para alterá-lo.")
                    tabContainer.ActiveTabIndex = 1
                    Exit Sub
                End If

                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(obj)
                txtClienteRetirada.Text = itemCliente.Text
                txtCodigoClienteRetirada.Value = itemCliente.Value
                Session.Remove("objClienteAxECR" & HID.Value)
                SalvarClienteRetirada()
            ElseIf Session("objPedidosXItens" & HID.Value) IsNot Nothing Then
                SessaoRecuperaAutorizacao()
                If CType(Session("objPedidosXItens" & HID.Value), [Lib].Negocio.Pedido).PedidoBloqueado Then
                    MsgBox(Me.Page, "Esse pedido ainda não foi Liberado!")
                    Exit Sub
                End If

                Dim objPedido As [Lib].Negocio.Pedido = CType(Session("objPedidosXItens" & HID.Value), [Lib].Negocio.Pedido)
                CargaListaAutorizacao(objPedido)
                Session.Remove("objPedidosXItens" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaListaAutorizacao(ByVal objPedido As [Lib].Negocio.Pedido)
        Dim lstAutorizacao As New [Lib].Negocio.ListAutorizacaoDeRetirada(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, "", 0, objPedido.Codigo, objPedido.SubOperacao.Classe)
        gridAut.DataSource = lstAutorizacao
        gridAut.DataBind()

        SessaoRecuperaAutorizacao()
        objAutorizacao.Autorizacao = lstAutorizacao.Count + 1
        SessaoSalvaAutorizacao()

        Dim Parametros As New Hashtable
        Parametros.Add("Empresa", objPedido.CodigoEmpresa)
        Parametros.Add("EndEmpresa", objPedido.EnderecoEmpresa)
        Parametros.Add("Pedido", objPedido.Codigo)
        Parametros.Add("Classe", objPedido.Operacao.CodigoClasse)
        Dim ListaItensPedido As New ListSaldoPedido2015(Parametros)
        Dim SaldoProduto As SaldoPedido2015 = ListaItensPedido(0)

        lblOperacaoPedido.Text = "OP: " & objPedido.CodigoOperacao & " - " & objPedido.CodigoSubOperacao & "  " & objPedido.SubOperacao.Descricao
        imgExtratoPedido.Visible = True
        lblProdutoPedido.Text = objPedido.Itens(0).CodigoProduto + " - " + objPedido.Itens(0).Descricao

        If lstAutorizacao IsNot Nothing Then
            If Funcoes.VerificaPermissao("PrestacaoDeServico", "GRAVAR") Then
                txtTaxa.Enabled = True
                lblTaxa.Text = ""
            Else
                txtTaxa.Enabled = False
            End If

            txtPedido.Text = SaldoProduto.CodigoPedido
            objAutorizacao.CodigoPedido = SaldoProduto.CodigoPedido

            txtMaxQuantidadeFiscal.Value = (SaldoProduto.QtdeContratadoFiscal - lstAutorizacao.QtdeAutorizadaFiscal) '+ objAutorizacao.QuantidadeAutorizadaFiscal
            txtMaxQuantidadeFisica.Value = (SaldoProduto.QtdeContratadoFisico - lstAutorizacao.QtdeAutorizadaFisica) '+ objAutorizacao.QuantidadeAutorizadaFisica

            txtQtdeContratada.Text = SaldoProduto.QtdeContratadoFiscal.ToString("N0")
            txtQtdeContratadaFisica.Text = SaldoProduto.QtdeContratadoFisico.ToString("N0")

            txtQtdeEntregue.Text = (SaldoProduto.QtdeEntregueFiscalDireta + SaldoProduto.QtdeEntregueFiscalAFixar + SaldoProduto.QtdeEntregueFiscalGlobal + SaldoProduto.QtdeEntregueFiscalDeposito).ToString("N0")
            txtQtdeEntregueFisica.Text = (SaldoProduto.QtdeEntregueFisicoDireta + SaldoProduto.QtdeEntregueFisicoAFixar + SaldoProduto.QtdeEntregueFisicoRemessa + SaldoProduto.QtdeEntregueFisicoDeposito).ToString("N0")

            txtSaldoFiscal.Text = (SaldoProduto.QtdeContratadoFiscal - (SaldoProduto.QtdeEntregueFiscalDireta + SaldoProduto.QtdeEntregueFiscalAFixar + SaldoProduto.QtdeEntregueFiscalGlobal + SaldoProduto.QtdeEntregueFiscalDeposito)).ToString("N0")
            txtSaldoFisico.Text = (SaldoProduto.QtdeContratadoFisico - (SaldoProduto.QtdeEntregueFisicoDireta + SaldoProduto.QtdeEntregueFisicoAFixar + SaldoProduto.QtdeEntregueFisicoRemessa + SaldoProduto.QtdeEntregueFisicoDeposito)).ToString("N0")

            txtAutorizadoFiscal.Text = lstAutorizacao.QtdeAutorizadaFiscal.ToString("N0")
            txtAutorizadoFisico.Text = lstAutorizacao.QtdeAutorizadaFisica.ToString("N0")

            If lstAutorizacao.Count > 0 Then
                tabContainer.ActiveTabIndex = 1
            End If
        End If
    End Sub

    Private Sub SetarEmpresa(ByVal Empresa As String, ByVal EndEmpresa As String)
        Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa, EndEmpresa)
        Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(objEmpresa, [Lib].Negocio.Cliente))
        txtEmpresa.Text = itemEmpresa.Text
        txtCodigoEmpresa.Value = itemEmpresa.Value
        SessaoRecuperaAutorizacao()
        objAutorizacao.CodigoEmpresa = Empresa
        objAutorizacao.EnderecoEmpresa = EndEmpresa
        SessaoSalvaAutorizacao()
    End Sub

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaAxE" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteAxE" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPedido.Click
        Try
            BuscarPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub BuscarPedido()
        If String.IsNullOrWhiteSpace(txtEmpresa.Text) OrElse String.IsNullOrWhiteSpace(txtCliente.Text) Then
            MsgBox(Me.Page, "Informe a empresa e o cliente antes de selecionar o pedido!")
            Exit Sub
        End If
        Dim empresa() As String = txtCodigoEmpresa.Value.Split("-")
        Dim cliente() As String = txtCodigoCliente.Value.Split("-")

        HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"

        Dim parameters As New Dictionary(Of String, Object)
        If Not String.IsNullOrWhiteSpace(HID.Value) Then parameters.Add("tipo", HID.Value)

        parameters.Add("empresa", empresa(0))
        parameters.Add("enderecoEmpresa", empresa(1))
        parameters.Add("cliente", cliente(0))
        parameters.Add("enderecoCliente", cliente(1))
        parameters.Add("situacao", 1)
        parameters.Add("granel", True)
        Session("ssTipoRetorno") = "objPedidosXItens" & HID.Value
        ucConsultaPedidos.Limpar()
        Dim numberRows As Integer = ucConsultaPedidos.BindGridView(parameters)
        If numberRows > 1 Then
            Popup.ConsultaDePedidos(Me.Page, "objPedidosXItens" & HID.Value)
        End If
    End Sub

    Public Sub Limpar()
        SessaoRecuperaAutorizacao()
        objAutorizacao = New [Lib].Negocio.AutorizacaoDeRetirada
        objAutorizacao.IUD = "I"
        SessaoSalvaAutorizacao()

        txtCodigoEmpresa.Value = ""
        btnEmpresa.Enabled = True
        txtMovimento.Text = DateTime.Now.ToShortDateString()

        lnkNovo.Parent.Visible = True
        lnkConsultar.Parent.Visible = True

        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        btnCliente.Enabled = True
        btnClienteRetirada.Enabled = True
        btnClienteAutorizante.Enabled = True
        txtQuantidadeFiscal.Enabled = True
        txtQuantidadeFisica.Enabled = True
        txtTaxa.Enabled = True

        lblOperacaoPedido.Text = ""
        lblProdutoPedido.Text = ""

        txtQtdeContratada.Text = "0"
        txtQtdeEntregue.Text = "0"
        txtSaldoFiscal.Text = "0"
        txtAutorizadoFiscal.Text = "0"

        txtQtdeContratadaFisica.Text = "0"
        txtQtdeEntregueFisica.Text = "0"
        txtSaldoFisico.Text = "0"
        txtAutorizadoFisico.Text = "0"

        gridAut.DataBind()

        Session.Remove("objEmpresaAxE" & HID.Value)
        Session.Remove("objClienteAxE" & HID.Value)
        Session.Remove("objItensPedidoSelecionadosAxE" & HID.Value)
        Session.Remove("objClienteAxECR" & HID.Value)
        Session.Remove("objClienteAxEA" & HID.Value)
        Session.Remove("objObservacaoAxE" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
        ucConsultaEmpresas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        ucConsultaAutorizacaoDeRetirada.SetarHID(HID.Value)
        ucConsultaObservacoes.SetarHID(HID.Value)

        objAutorizacao.CodigoEmpresa = Session("ssEmpresa")
        objAutorizacao.EnderecoEmpresa = Session("ssEndEmpresa")
        'Dim n As New [Lib].Negocio.Numerador(objAutorizacao.CodigoEmpresa, objAutorizacao.EnderecoEmpresa, 50)
        'objAutorizacao.Autorizacao = n.Sequencia + 1
        'txtAutorizacao.Text = n.Sequencia + 1
        SessaoSalvaAutorizacao()
        AtualizaFormularioComAClasse()
        tabContainer.ActiveTabIndex = 0
    End Sub

    Public Sub Bloquear()
        lnkNovo.Parent.Visible = False
        lnkConsultar.Parent.Visible = False
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkLimpar.Parent.Visible = False
        txtTaxa.Enabled = False
    End Sub

    Public Sub AtualizaFormularioComAClasse()
        SessaoRecuperaAutorizacao()
        SetarEmpresa(objAutorizacao.CodigoEmpresa, objAutorizacao.EnderecoEmpresa)

        If objAutorizacao.Pedido Is Nothing Then
            txtCliente.Text = ""
            txtCodigoCliente.Value = ""
        Else
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objAutorizacao.Pedido.Cliente)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
        End If

        If objAutorizacao.ClienteRetirada Is Nothing Then
            txtClienteRetirada.Text = ""
            txtCodigoClienteRetirada.Value = ""
        Else
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objAutorizacao.ClienteRetirada)
            txtClienteRetirada.Text = itemCliente.Text
            txtCodigoClienteRetirada.Value = itemCliente.Value
        End If

        txtPedido.Text = objAutorizacao.CodigoPedido
        txtAutorizacao.Text = objAutorizacao.Autorizacao

        txtMovimento.Text = objAutorizacao.Movimento.ToShortDateString()
        txtQuantidadeFiscal.Text = objAutorizacao.QuantidadeAutorizadaFiscal.ToString("N0")
        txtQuantidadeFisica.Text = objAutorizacao.QuantidadeAutorizadaFisica.ToString("N0")
        txtTaxa.Text = objAutorizacao.Taxa.ToString("N2")
        txtQtdeEntregue.Text = objAutorizacao.QuantidadeEntregueFiscal.ToString("N0")
        txtQtdeEntregueFisica.Text = objAutorizacao.QuantidadeEntregueFisica.ToString("N0")
        txtSaldoFiscal.Text = objAutorizacao.SaldoFiscal.ToString("N0")
        txtSaldoFisico.Text = objAutorizacao.SaldoFisico.ToString("N0")

        If IsNumeric(txtMaxQuantidadeFiscal.Value) AndAlso IsNumeric(txtMaxQuantidadeFisica.Value) Then
            txtMaxQuantidadeFiscal.Value = CDec(txtMaxQuantidadeFiscal.Value) + objAutorizacao.QuantidadeAutorizadaFiscal
            txtMaxQuantidadeFisica.Value = CDec(txtMaxQuantidadeFisica.Value) + objAutorizacao.QuantidadeAutorizadaFisica
        End If

        If objAutorizacao.IUD <> "I" Then
            txtEmpresa.Enabled = False
            btnEmpresa.Enabled = False
            btnCliente.Enabled = False
            txtPedido.Enabled = False
            btnPedido.Enabled = False

            If objAutorizacao.QuantidadeEntregueFisica > 0 Or objAutorizacao.QuantidadeEntregueFiscal > 0 Then
                btnClienteRetirada.Enabled = False
                btnClienteAutorizante.Enabled = False
            End If

            If objAutorizacao.CodigoPedidoServico > 0 Then
                If objAutorizacao.PedidoServico.Vencimentos(0).Provisao = [Lib].Negocio.eProvisao.Baixa Then
                    txtQuantidadeFiscal.Enabled = False
                    txtQuantidadeFisica.Enabled = False
                    txtTaxa.Enabled = False
                End If
            End If
        End If

        If objAutorizacao.ClienteAutorizante Is Nothing Then
            txtAutorizante.Text = ""
            txtCodigoAutorizante.Value = ""
        Else
            Dim itemAutorizante As ListItem = Funcoes.FormatarListItemCliente(objAutorizacao.ClienteAutorizante)
            txtAutorizante.Text = itemAutorizante.Text
            txtCodigoAutorizante.Value = itemAutorizante.Value
        End If

        txtDataBase.Text = objAutorizacao.DataBaseCalculo.ToShortDateString()
        txtVencimento.Text = objAutorizacao.DataVencimento.ToShortDateString()

        txtPedidoServico.Text = objAutorizacao.CodigoPedidoServico
        txtObservacao.Text = objAutorizacao.Observacao

        txtTitulo.Text = objAutorizacao.CodigoTitulo
        txtValor.Text = objAutorizacao.ValorTitulo.ToString("N2")
    End Sub

    Protected Sub gridAut_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaAutorizacao()

            objAutorizacao = New [Lib].Negocio.AutorizacaoDeRetirada(objAutorizacao.CodigoEmpresa, objAutorizacao.EnderecoEmpresa, objAutorizacao.CodigoPedido, gridAut.Rows(gridAut.SelectedIndex).Cells(1).Text, objAutorizacao.Pedido.SubOperacao.Classe)
            SessaoSalvaAutorizacao()
            AtualizaFormularioComAClasse()
            tabContainer.ActiveTabIndex = 0

            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            lnkNovo.Parent.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtTaxa_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not IsNumeric(txtTaxa.Text) Then Exit Sub
            SessaoRecuperaAutorizacao()
            objAutorizacao.Taxa = txtTaxa.Text
            txtValor.Text = objAutorizacao.ValorTitulo.ToString("N2")
            SessaoSalvaAutorizacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnClienteRetirada_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClienteRetirada.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteAxECR" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtCodigoEmpresa_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If txtCodigoEmpresa.Value.Length > 0 Then
                Dim emp() As String = txtCodigoEmpresa.Value.Split("-")
                SessaoRecuperaAutorizacao()
                objAutorizacao.CodigoEmpresa = emp(0)
                objAutorizacao.EnderecoEmpresa = emp(1)
                SessaoSalvaAutorizacao()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtCodigoClienteRetirada_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SalvarClienteRetirada()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtCodigoAutorizante_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SalvarAutorizante()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub SalvarClienteRetirada()
        If txtCodigoClienteRetirada.Value.Length > 0 Then
            Dim cli() As String = txtCodigoClienteRetirada.Value.Split("-")
            Dim objCliente As New [Lib].Negocio.Cliente(cli(0), cli(1))
            If Not objCliente Is Nothing AndAlso objCliente.Codigo.Length > 0 Then
                SessaoRecuperaAutorizacao()
                objAutorizacao.CodigoClienteRetirada = cli(0)
                objAutorizacao.EnderecoClienteRetirada = cli(1)
                SessaoSalvaAutorizacao()
            End If
        End If
    End Sub

    Private Sub SalvarAutorizante()
        If txtCodigoAutorizante.Value.Length > 0 Then
            Dim aut() As String = txtCodigoAutorizante.Value.Split("-")
            SessaoRecuperaAutorizacao()
            objAutorizacao.CodigoAutorizante = aut(0)
            objAutorizacao.EnderecoAutorizante = aut(1)
            SessaoSalvaAutorizacao()
        End If
    End Sub

    Protected Sub btnClienteAutorizante_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClienteAutorizante.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteAxEA" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function ValidaAutorizacao() As Boolean
        SessaoRecuperaAutorizacao()

        If objAutorizacao.CodigoPedido = 0 Then
            MsgBox(Me.Page, "Selecione Um Pedido para Prosseguir.")
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtQuantidadeFiscal.Text) Then
            txtQuantidadeFiscal.Text = 0
        End If

        If String.IsNullOrWhiteSpace(txtQuantidadeFisica.Text) Then
            txtQuantidadeFisica.Text = 0
        End If

        If Not IsNumeric(txtMaxQuantidadeFiscal.Value) OrElse txtMaxQuantidadeFiscal.Value = 0 Then
            If objAutorizacao.Pedido.SubOperacao.EntradaSaida.ToString.Substring(0, 1) = "S" AndAlso (objAutorizacao.Pedido.SubOperacao.Classe = [Lib].Negocio.eClassesOperacoes.DEPOSITOS Or objAutorizacao.Pedido.SubOperacao.Classe = [Lib].Negocio.eClassesOperacoes.AFIXAR) Then
            ElseIf objAutorizacao.QuantidadeAutorizadaFiscal = objAutorizacao.QuantidadeEntregueFiscal Then
            Else
                MsgBox(Me.Page, "Este Pedido Não tem Saldo Fiscal a ser Autorizado.")
                Return False
            End If
        End If

        If txtMaxQuantidadeFisica.Value = 0 Then
            If objAutorizacao.Pedido.SubOperacao.EntradaSaida.ToString.Substring(0, 1) = "S" AndAlso (objAutorizacao.Pedido.SubOperacao.Classe = [Lib].Negocio.eClassesOperacoes.DEPOSITOS Or objAutorizacao.Pedido.SubOperacao.Classe = [Lib].Negocio.eClassesOperacoes.AFIXAR) Then
            ElseIf objAutorizacao.QuantidadeAutorizadaFisica = objAutorizacao.QuantidadeEntregueFisica Then
            Else
                MsgBox(Me.Page, "Este Pedido Não tem Saldo Físico a ser Autorizado.")
                Return False
            End If
        End If

        If CDbl(txtQuantidadeFiscal.Text) = 0 Then
            MsgBox(Me.Page, "Informe a Quantidade da autorização de retirada.")
            Return False
        End If

        If (String.IsNullOrWhiteSpace(txtCodigoClienteRetirada.Value)) Then
            MsgBox(Me.Page, "Selecione o cliente de retirada!")
            Return False
        End If

        If objAutorizacao.CodigoAutorizante.Length = 0 Then
            MsgBox(Me.Page, "Selecione O Autorizante.")
            Return False
        End If

        If objAutorizacao.CodigoClienteRetirada.Length = 0 Then
            MsgBox(Me.Page, "Selecione o Cliente de Retirada/Destino.")
            Return False
        End If

        If objAutorizacao.QuantidadeEntregueFisica > 0 AndAlso objAutorizacao.QuantidadeEntregueFisica > objAutorizacao.QuantidadeAutorizadaFisica Then
            MsgBox(Me.Page, "A Quantidade Física Autorizada não pode ser Alterada para uma Quantidade Física Menor do que a Quantidade Física Entregue.")
            Return False
        End If

        If objAutorizacao.QuantidadeEntregueFiscal > 0 AndAlso objAutorizacao.QuantidadeEntregueFiscal > objAutorizacao.QuantidadeAutorizadaFiscal Then
            MsgBox(Me.Page, "A Quantidade Fiscal Autorizada não pode ser Alterada para uma Quantidade Fiscal Menor do que a Quantidade Fiscal Entregue.")
            Return False
        End If

        If CDbl(txtQuantidadeFiscal.Text) > CDbl(txtMaxQuantidadeFiscal.Value) Then
            If objAutorizacao.Pedido.SubOperacao.EntradaSaida.ToString.Substring(0, 1) = "S" AndAlso (objAutorizacao.Pedido.SubOperacao.Classe = [Lib].Negocio.eClassesOperacoes.DEPOSITOS Or objAutorizacao.Pedido.SubOperacao.Classe = [Lib].Negocio.eClassesOperacoes.AFIXAR) AndAlso objAutorizacao.Pedido.SubOperacao.Devolucao = False Then
            ElseIf objAutorizacao.QuantidadeAutorizadaFiscal = objAutorizacao.QuantidadeEntregueFiscal Then
            Else
                MsgBox(Me.Page, "O Valor Maximo para a Quantidade Fiscal da Autorização é de: " & Str(CDbl(txtMaxQuantidadeFiscal.Value)))
                Return False
            End If
        End If

        If CDbl(txtQuantidadeFisica.Text) > CDbl(txtMaxQuantidadeFisica.Value) Then
            If objAutorizacao.Pedido.SubOperacao.EntradaSaida.ToString.Substring(0, 1) = "S" AndAlso (objAutorizacao.Pedido.SubOperacao.Classe = [Lib].Negocio.eClassesOperacoes.DEPOSITOS Or objAutorizacao.Pedido.SubOperacao.Classe = [Lib].Negocio.eClassesOperacoes.AFIXAR) AndAlso objAutorizacao.Pedido.SubOperacao.Devolucao = False Then
            ElseIf objAutorizacao.QuantidadeAutorizadaFisica = objAutorizacao.QuantidadeEntregueFisica Then
            Else
                MsgBox(Me.Page, "O Valor Maximo para a Quantidade Física da Autorização é de: " & Str(CDbl(txtMaxQuantidadeFisica.Value)))
                Return False
            End If
        End If

        If objAutorizacao.Taxa > 0 And objAutorizacao.Pedido.SubOperacao.Classe <> [Lib].Negocio.eClassesOperacoes.DEPOSITOS And objAutorizacao.Pedido.SubOperacao.Classe <> [Lib].Negocio.eClassesOperacoes.AFIXAR Then
            MsgBox(Me.Page, "A Cobrança de Serviços de Armazenagem só se dá em Pedidos com a Classe Deposito ou A Fixar.")
            Return False
        End If

        Return True
    End Function

    Protected Sub txtQuantidadeFiscal_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not IsNumeric(txtQuantidadeFiscal.Text) Then Exit Sub
            SessaoRecuperaAutorizacao()
            objAutorizacao.QuantidadeAutorizadaFiscal = txtQuantidadeFiscal.Text
            txtValor.Text = objAutorizacao.ValorTitulo.ToString("N2")
            txtQuantidadeFisica.Focus()
            SessaoSalvaAutorizacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtQuantidadeFisica_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not IsNumeric(txtQuantidadeFisica.Text) Then Exit Sub
            SessaoRecuperaAutorizacao()
            objAutorizacao.QuantidadeAutorizadaFisica = txtQuantidadeFisica.Text
            SessaoSalvaAutorizacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtObservacao_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaAutorizacao()
            objAutorizacao.Observacao = txtObservacao.Text
            SessaoSalvaAutorizacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgNfs_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("AutorizacaoDeRetirada", "RELATORIO") Then
                SessaoRecuperaAutorizacao()
                If objAutorizacao Is Nothing Then
                    MsgBox(Me.Page, "É necessário selecionar uma autorização!")
                    Exit Sub
                End If

                Dim ds As DataSet
                Dim sql As String = ""
                Dim Parametros As String = "Parametros:" & vbCrLf

                Parametros &= "Empresa: " & objAutorizacao.CodigoEmpresa & "-" & objAutorizacao.EnderecoEmpresa & "-" & objAutorizacao.Empresa.Nome & vbCrLf & _
                              "Autorização: " & objAutorizacao.Autorizacao & vbCrLf & _
                              "Pedido: " & objAutorizacao.CodigoPedido & vbCrLf

                sql &= " SELECT NotasFiscais.Empresa_Id,  " & vbCrLf & _
                       " NotasFiscais.EndEmpresa_Id,  " & vbCrLf & _
                       " Empresa.Nome as NomeEmpresa, " & vbCrLf & _
                       " NotasFiscais.Cliente_Id, " & vbCrLf & _
                       " NotasFiscais.EndCliente_Id, " & vbCrLf & _
                       " Clientes.Cidade, " & vbCrLf & _
                       " Clientes.Complemento, " & vbCrLf & _
                       " Clientes.Nome AS NomeCliente, " & vbCrLf & _
                       " NotasFiscais.Pedido, " & vbCrLf & _
                       " NotasFiscais.DataDaNota, " & vbCrLf & _
                       " NotasFiscais.Nota_Id, " & vbCrLf & _
                       " NotasFiscais.Serie_Id, " & vbCrLf & _
                       " NotasFiscais.EntradaSaida_Id, " & vbCrLf & _
                       " NotasFiscaisXItens.QuantidadeFisica, " & vbCrLf & _
                       " NotasFiscaisXItens.QuantidadeFiscal " & vbCrLf & _
                       " FROM NotasFiscais  " & vbCrLf & _
                       " INNER JOIN Clientes as Empresa  " & vbCrLf & _
                       " ON NotasFiscais.Empresa_Id    = Empresa.Cliente_Id  " & vbCrLf & _
                       " AND NotasFiscais.EndEmpresa_Id = Empresa.Endereco_Id  " & vbCrLf & _
                       " INNER JOIN Clientes  " & vbCrLf & _
                       " ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id  " & vbCrLf & _
                       " AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id  " & vbCrLf & _
                       " INNER JOIN NotasFiscaisXItens  " & vbCrLf & _
                       " ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id  " & vbCrLf & _
                       " AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id  " & vbCrLf & _
                       " AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id  " & vbCrLf & _
                       " AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id  " & vbCrLf & _
                       " AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id  " & vbCrLf & _
                       " AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id  " & vbCrLf & _
                       " AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                       " INNER JOIN  AutorizacaoDeRetirada ON NotasFiscais.Empresa_Id = AutorizacaoDeRetirada.Empresa_Id " & vbCrLf & _
                       " AND NotasFiscais.EndEmpresa_Id = AutorizacaoDeRetirada.EndEmpresa_id " & vbCrLf & _
                       " AND NotasFiscais.Pedido = AutorizacaoDeRetirada.Pedido_id " & vbCrLf & _
                       " AND NotasFiscais.Autorizacao = AutorizacaoDeRetirada.Autorizacao_Id " & vbCrLf & _
                       " WHERE NotasFiscais.Autorizacao = " & objAutorizacao.Autorizacao & " " & vbCrLf & _
                       " AND NotasFiscais.Empresa_Id = " & objAutorizacao.CodigoEmpresa & " " & vbCrLf & _
                       " AND NotasFiscais.EndEmpresa_Id = " & objAutorizacao.EnderecoEmpresa & " " & vbCrLf & _
                       " AND NotasFiscais.Pedido = " & objAutorizacao.CodigoPedido & " " & vbCrLf

                ds = Banco.ConsultaDataSet(sql, "NfsAutorizacaoDeRetirada")

                Dim crpt As New ReportDocument()

                Try
                    crpt.FileName = Server.MapPath("~/Reports/Cr_NfsAutorizacaoDeRetirada.rpt")
                    crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                    Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                    Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                    Dim arquivo As String = NomeArquivo

                    crpt.SetDataSource(ds)

                    Dim crparametervalues As ParameterValues
                    Dim crparameterdiscretevalue As ParameterDiscreteValue
                    Dim crparameterfielddefinitions As ParameterFieldDefinitions
                    Dim crparameterfielddefinition As ParameterFieldDefinition

                    crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                    crparameterfielddefinition = crparameterfielddefinitions.Item("Nome")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = objAutorizacao.Empresa.Nome
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("Parametros")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = Parametros
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("Cidade")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = objAutorizacao.Empresa.Cidade & " - " & objAutorizacao.Empresa.CodigoEstado
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("Titulo")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = " Relatório De Notas Fiscais Por Autorizacao De Retirada "
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    crparameterfielddefinition = crparameterfielddefinitions.Item("CNPJ")
                    crparametervalues = crparameterfielddefinition.CurrentValues
                    crparameterdiscretevalue = New ParameterDiscreteValue
                    crparameterdiscretevalue.Value = "CNPJ: " & Funcoes.FormatarCpfCnpj(objAutorizacao.CodigoEmpresa)
                    crparametervalues.Add(crparameterdiscretevalue)
                    crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                    If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                    crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

                    If IO.File.Exists(arquivo) Then
                        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo2 & "');", True)
                    End If
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                Finally
                    crpt.Close()
                    crpt.Dispose()
                End Try
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("AutorizacaoDeRetirada", "GRAVAR") Then
                If Not ValidaAutorizacao() Then Exit Sub

                SessaoRecuperaAutorizacao()
                objAutorizacao.Movimento = CDate(txtMovimento.Text)
                objAutorizacao.DataBaseCalculo = CDate(txtDataBase.Text)
                objAutorizacao.DataVencimento = CDate(txtVencimento.Text)

                objAutorizacao.Observacao = Funcoes.EliminarCaracteresEspeciais(txtObservacao.Text)

                objAutorizacao.UsuarioDeInclusao = Session("ssNomeUsuario")
                objAutorizacao.UsuarioInclusaoData = DateTime.Now
                If objAutorizacao.Salvar() Then
                    MsgBox(Me.Page, "Autorização Gravada com Sucesso.", eTitulo.Sucess)
                    Limpar()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("AutorizacaoDeRetirada", "ALTERAR") Then
                SessaoRecuperaAutorizacao()
                If Not ValidaAutorizacao() Then Exit Sub
                objAutorizacao.Movimento = CDate(txtMovimento.Text)
                objAutorizacao.DataBaseCalculo = CDate(txtDataBase.Text)
                objAutorizacao.DataVencimento = CDate(txtVencimento.Text)

                objAutorizacao.Observacao = Funcoes.EliminarCaracteresEspeciais(txtObservacao.Text)

                objAutorizacao.UsuarioDeAlteracao = Session("ssNomeUsuario")
                objAutorizacao.UsuarioAlteracaoData = DateTime.Now
                objAutorizacao.IUD = "U"
                If objAutorizacao.Salvar() Then
                    MsgBox(Me.Page, "Autorização Alterada com Sucesso.", eTitulo.Sucess)
                    Limpar()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ExisteRomaneio(ByVal autorizacao As [Lib].Negocio.AutorizacaoDeRetirada) As Boolean

        Dim sql As String = "       SELECT TOP 1 Autorizacao 
                                    FROM Romaneios 
                                    WHERE Empresa_Id = '" & autorizacao.CodigoEmpresa & "'
                                        AND Destino = '" & autorizacao.CodigoClienteRetirada & "' 
                                        AND autorizacao = " & autorizacao.CodigoEmpresa & "
                                        AND Pedido = '" & autorizacao.Pedido.Codigo & ";' "

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Romaneio")

        If ds.Tables(0).Rows.Count = 1 Then Return True
        Return False
    End Function

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("AutorizacaoDeRetirada", "EXCLUIR") Then
                SessaoRecuperaAutorizacao()

                If ExisteRomaneio(objAutorizacao) OrElse objAutorizacao.QuantidadeEntregueFiscal > 0 OrElse objAutorizacao.QuantidadeEntregueFisica > 0 Then
                    MsgBox(Me.Page, "A autorizacao já foi utilizada e não pode ser Apagada")
                    Exit Sub
                End If

                If objAutorizacao.CodigoPedidoServico > 0 AndAlso objAutorizacao.PedidoServico.Vencimentos(0).Provisao = [Lib].Negocio.eProvisao.Baixa Then
                    MsgBox(Me.Page, "O Titulo vinculado a esta autorizacao já foi baixado e não pode ser Apagado")
                    Exit Sub
                End If

                objAutorizacao.IUD = "D"
                If objAutorizacao.Salvar() Then
                    MsgBox(Me.Page, "Autorização Apagada com Sucesso.", eTitulo.Sucess)
                    Limpar()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMesage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("AutorizacaoDeRetirada", "LEITURA") Then
                If String.IsNullOrWhiteSpace(txtEmpresa.Text) Then
                    MsgBox(Me.Page, "Informe a empresa.", eTitulo.Info)
                ElseIf String.IsNullOrWhiteSpace(txtCliente.Text) AndAlso (String.IsNullOrWhiteSpace(txtPedido.Text) OrElse txtPedido.Text = 0) Then
                    MsgBox(Me.Page, "Informe o cliente ou pedido.")
                    Exit Sub
                Else
                    SessaoRecuperaAutorizacao()
                    Dim objPedido As [Lib].Negocio.Pedido = New Pedido(txtCodigoEmpresa.Value.Split("-")(0), txtCodigoEmpresa.Value.Split("-")(1), txtPedido.Text)
                    If objPedido IsNot Nothing AndAlso objPedido.Codigo > 0 Then
                        CargaListaAutorizacao(objPedido)
                        tabContainer.ActiveTabIndex = 1
                    Else
                        MsgBox(Me.Page, "Nenhum pedido encontrado.")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnObservacao_Click(ByVal sender As Object, ByVal e As EventArgs) Handles BtnObservacao.Click
        Try
            HttpContext.Current.Session("ssCampo" & HID.Value) = "AutorizacaoDeRetirada"
            HttpContext.Current.Session("Observacoes" & HID.Value) = txtObservacao.Text
            ucConsultaObservacoes.BindGridView()
            Popup.ConsultaDeObservacoes(Me.Page, "objObservacaoAxE" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExtratoPedido_Click(sender As Object, e As ImageClickEventArgs) Handles imgExtratoPedido.Click
        Try
            SessaoRecuperaAutorizacao()
            If Not objAutorizacao.Pedido Is Nothing Then
                Extrato.Emitir(Me.Page, FinanceiroNovo, objAutorizacao.Pedido.CodigoEmpresa, objAutorizacao.Pedido.EnderecoEmpresa, "T", objAutorizacao.CodigoPedido)
                'Dim strQueryString As String = "?fim=" & DateTime.Now.ToString("dd/MM/yyyy")
                'strQueryString &= "&empresa=" & objAutorizacao.Pedido.CodigoEmpresa & "-" & objAutorizacao.EnderecoEmpresa
                'strQueryString &= "&cliente=" & objAutorizacao.Pedido.CodigoCliente & "-" & objAutorizacao.Pedido.EnderecoCliente
                'strQueryString &= "&pedido=" & objAutorizacao.CodigoPedido
                'strQueryString &= "&es=ES"

                'ScriptManager.RegisterClientScriptBlock(Me, Me.GetType, "CarregarHTML", "window.open(""ExtratoPedidoHTML.aspx" & strQueryString & """, ""relatorio"", ""status=no, toolbar=yes, location=no, menubar=yes, resizable=yes, height=600, width=800, scrollbars=yes"");", True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AUTORIZACAODERETIRADA")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class