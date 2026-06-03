Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AgrupamentodePesagem
    Inherits BasePage

    Private objPesagem As [Lib].Negocio.Pesagem

    Private Sub SalvarObjetoSessao()
        Session("objPesagem" & HID.Value) = objPesagem
    End Sub

    Private Sub CarregarObjetoSessao()
        objPesagem = CType(Session("objPesagem" & HID.Value), [Lib].Negocio.Pesagem)
    End Sub

#Region "Events"

    Protected Sub PageLoad(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AgrupamentoDePesagem", "ACESSAR") Then
                If Funcoes.VerificaPermissao("AgrupamentoDePesagem", "LEITURA") Then
                    Limpar("")
                    LiberaEmpresa()
                    CargaUnidadeDeNegocioEmpresa()
                    txtData1.Text = New DateTime(Now.Year, Now.Month, 1).ToString("dd/MM/yyyy")
                    txtData2.Text = Now().ToString("dd/MM/yyyy")
                    txtdata1agr.Text = New DateTime(Now.Year, Now.Month, 1).ToString("dd/MM/yyyy")
                    txtdata2agr.Text = Now().ToString("dd/MM/yyyy")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEmpresa.SelectedIndexChanged
        Try
            CarregarObjetoSessao()
            Dim strEmpresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            objPesagem.CodigoEmpresa = strEmpresa(0)
            objPesagem.EnderecoEmpresa = strEmpresa(1)
            objPesagem.CodigoCliente = ""
            objPesagem.EnderecoCliente = 0
            objPesagem.CodigoPedido = 0

            SalvarObjetoSessao()

            txtCliente.Text = ""
            txtCodigoCliente.Value = 0
            txtPedido.Text = ""
            txtCodigoPedido.Value = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteAXP" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            ElseIf txtCodigoCliente.Value.ToString().Length() = 1 Then
                MsgBox(Me.Page, "Cliente não foi selecionado.")
            Else
                BuscarPedidos()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Validar_Selecao() Then
                CarregarObjetoSessao()

                Dim contemdatas As Boolean = False
                If Not String.IsNullOrWhiteSpace(txtdata1agr.Text) AndAlso Not String.IsNullOrWhiteSpace(txtdata2agr.Text.Trim) Then
                    If IsDate(txtdata1agr.Text) AndAlso IsDate(txtdata2agr.Text) Then
                        contemdatas = True
                    End If
                End If

                Dim ListPesagem As New [Lib].Negocio.ListPesagem(objPesagem, True, IIf(contemdatas, txtdata1agr.Text, ""), IIf(contemdatas, txtdata2agr.Text, ""), "", True)
                If ListPesagem IsNot Nothing AndAlso ListPesagem.Count > 0 Then
                    Session("objListPesagem" & HID.Value) = ListPesagem
                    gridAgrupamentoDePesagem.DataSource = ListPesagem.ToArray()
                    gridAgrupamentoDePesagem.DataBind()

                    objPesagem.Analises = New [Lib].Negocio.ListPesagemXAnalises(objPesagem, True)
                    gridDescontos.DataSource = objPesagem.Analises.ToArray()
                    gridDescontos.DataBind()
                Else
                    gridDescontos.DataSource = Nothing
                    gridDescontos.DataBind()
                    MsgBox(Me.Page, "Sem laudos a ser agrupados para esta consulta.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgIncluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim i As Integer = 0
            hConfirmar.Value = 0

            Dim INTACTAMonsanto As Decimal = 0

            CarregaSelecionados()

            If Session("objListPesagem" & HID.Value) IsNot Nothing Then
                For i = 0 To CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem).Count - 1
                    If CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)(i).SelecionaPesagem Then hConfirmar.Value += 1

                    If INTACTAMonsanto = 0 Then
                        For Each analise In CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)(i).Analises
                            If analise.CodigoAnalise = 12 Then
                                INTACTAMonsanto = analise.Percentual
                            End If
                        Next
                    End If
                Next
            End If

            If hConfirmar.Value > 1 Then
                If Funcoes.VerificaPermissao("AgrupamentoDePesagem", "GRAVAR") Then
                    CarregarObjetoSessao()
                    Dim strHistorico As String = ""
                    Dim strSeparador As String = ""
                    Dim primeiro As Boolean = True
                    Dim MaiorData As DateTime

                    For Each row As [Lib].Negocio.Pesagem In CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)
                        If row.SelecionaPesagem Then
                            strHistorico &= strSeparador & row.Codigo
                            strSeparador = ","

                            If primeiro Then
                                MaiorData = row.Movimento

                                objPesagem.CodigoPedido = row.CodigoPedido
                                objPesagem.CodigoProduto = row.CodigoProduto
                                objPesagem.CodigoTabelaDeClassificacao = row.CodigoTabelaDeClassificacao

                                objPesagem.CodigoDepositante = row.CodigoDepositante
                                objPesagem.EnderecoDepositante = row.EnderecoDepositante
                                objPesagem.CodigoPlaca = row.CodigoPlaca
                                objPesagem.CodigoTransportador = row.CodigoTransportador
                                objPesagem.EnderecoTransportador = row.EnderecoTransportador
                                objPesagem.CodigoMotorista = row.CodigoMotorista
                                objPesagem.EnderecoMotorista = row.EnderecoMotorista
                                objPesagem.CodigoDeposito = row.CodigoDeposito
                                objPesagem.EnderecoDeposito = row.EnderecoDeposito
                                objPesagem.CodigoOperacao = row.CodigoOperacao
                                objPesagem.CodigoSubOperacao = row.CodigoSubOperacao
                                objPesagem.EntradaSaida = row.EntradaSaida

                                objPesagem.PrimeiraPesagem = txtPrimeiraPesagem.Text
                                objPesagem.SegundaPesagem = txtSegundaPesagem.Text
                                objPesagem.BrutoBalanca = txtPesoBruto.Text
                                objPesagem.Desconto = txtDesconto.Text
                                objPesagem.Liquido = txtLiquido.Text

                                objPesagem.CodigoViaDeTransporte = row.CodigoViaDeTransporte
                                objPesagem.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
                                objPesagem.UsuarioAlteracao = ""

                                objPesagem.UsuarioReimpressao = ""
                                objPesagem.UsuarioCancelamento = ""
                                objPesagem.TemRomaneio = row.TemRomaneio                            
                                objPesagem.Processo = "AGRUPAMENTO"
                                objPesagem.CodigoAutorizacao = row.CodigoAutorizacao
                            ElseIf row.Movimento > MaiorData Then
                                MaiorData = row.Movimento
                            End If

                            primeiro = False
                        End If
                    Next

                    objPesagem.EntradaBalanca = MaiorData
                    objPesagem.EntradaPatio = MaiorData
                    objPesagem.SaidaBalanca = MaiorData
                    objPesagem.Movimento = MaiorData
                    objPesagem.DataAlteracao = MaiorData
                    objPesagem.DataReimpressao = MaiorData
                    objPesagem.DataCancelamento = MaiorData
                    objPesagem.CriarRomaneio = True

                    objPesagem.Observacoes &= "AGRUPAMENTO DO(S) LAUDO(S) " & strHistorico

                    If INTACTAMonsanto > 0 Then
                        For Each analise In objPesagem.Analises
                            If analise.CodigoAnalise = 12 Then
                                analise.Percentual = INTACTAMonsanto
                            End If
                        Next
                    End If

                    Dim SqlArray As New ArrayList
                    objPesagem.SalvarSqlAgrupamento(SqlArray, CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem))

                    If Banco.GravaBanco(SqlArray) Then
                        Dim codigoEmpresa As String = objPesagem.CodigoEmpresa
                        Dim endEmpresa As Integer = objPesagem.EnderecoEmpresa
                        Dim pesagem As Integer = objPesagem.Codigo

                        objPesagem = New [Lib].Negocio.Pesagem(codigoEmpresa, endEmpresa, pesagem)
                        Imprimir_Pesagem(objPesagem.CodigoEmpresa, objPesagem.EnderecoEmpresa, objPesagem.Codigo, "", objPesagem.Romaneios(0).Codigo)

                        MsgBox(Me.Page, "Agrupamento realizado com Sucesso.", eTitulo.Sucess)

                        'txtClienteDes.Text = txtCliente.Text
                        'txtCodigoClienteDes.Value = txtCodigoClienteDes.Value
                        'txtData1.Text = objPesagem.Movimento.ToString("dd/MM/yyyy")
                        'txtData2.Text = objPesagem.Movimento.ToString("dd/MM/yyyy")
                        'ConsultarAgrupado("")
                        Limpar("C")
                        'TabContainer1.ActiveTabIndex = 1
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
                End If
            Else
                MsgBox(Me.Page, "Para finalizar o agrupamento, dois ou mais laudos devem ser selecionados.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgPesagem_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgLaudoDePesagem As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgLaudoDePesagem.NamingContainer, GridViewRow)

            Imprimir_Pesagem(CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)(row.RowIndex).CodigoEmpresa, CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)(row.RowIndex).EnderecoEmpresa, CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)(row.RowIndex).Codigo, "REEMISSÃO", CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)(row.RowIndex).Romaneios(0).Codigo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar("L")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AgrupamentoDePesagem")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocioDes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidaDeNegocioDes.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresaDes, CarregarDDL.Tabela.Empresas, ddlUnidaDeNegocioDes.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEmpresaDes_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlEmpresaDes.SelectedIndexChanged
        Try
            CarregarObjetoSessao()
            Dim strEmpresa() As String = ddlEmpresaDes.SelectedValue.ToString.Split("-")
            objPesagem.CodigoEmpresa = strEmpresa(0)
            objPesagem.EnderecoEmpresa = strEmpresa(1)
            objPesagem.CodigoCliente = ""
            objPesagem.EnderecoCliente = 0
            objPesagem.CodigoPedido = 0

            SalvarObjetoSessao()

            txtClienteDes.Text = ""
            txtCodigoClienteDes.Value = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnClienteDes_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClienteDes.Click
        Try
            Limpar("L")
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteDXP" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultarD_Click(sender As Object, e As EventArgs) Handles lnkConsultarD.Click
        Try
            ConsultarAgrupado("")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgDesfazerAgrupamento_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgLaudo As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgLaudo.NamingContainer, GridViewRow)
            Dim SqlArray As New ArrayList
            Dim sql As String = ""

            '1 - Pega laudo agrupado para Desfazer o Agrupamento
            Dim objPesagem As Pesagem = CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)(row.RowIndex)

            '2 - Pega os filhos do Agrupado para Recriar Romaneio e vínculos
            Dim ListPesagem As New [Lib].Negocio.ListPesagem(objPesagem, True, txtData1.Text.Trim(), txtData2.Text.Trim(), "", True, True)

            '3 - Mudar Laudo Agrupado para Cancelado e remover Romaneio
            objPesagem.IUD = "D"
            For Each ro As Romaneio In objPesagem.Romaneios
                ro.IUD = "D"
            Next
            objPesagem.SalvarSql(SqlArray)

            '4 - Criar Romaneio para os Filhos e os Vínculos
            objPesagem.DesfazerAgrupamento(SqlArray, ListPesagem)

            If Banco.GravaBanco(SqlArray) Then
                ConsultarAgrupado("D")
                MsgBox(Me.Page, "Desagrupamento realizado com Sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparD_Click(sender As Object, e As EventArgs) Handles lnkLimparD.Click
        Try
            LimparCamposDes()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub CargaUnidadeDeNegocioEmpresa()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)

        ddl.Carregar(ddlUnidaDeNegocioDes, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidaDeNegocioDes, ddlEmpresaDes)

        CarregarObjetoSessao()

        Dim strEmpresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
        objPesagem.CodigoEmpresa = strEmpresa(0)
        objPesagem.EnderecoEmpresa = strEmpresa(1)

        SalvarObjetoSessao()
    End Sub

    Private Sub Limpar(ByVal tipo As String)
        Session.Remove("objClienteAXP" & HID.Value)
        Session.Remove("objPedidoAxP" & HID.Value)
        Session.Remove("objListPesagem" & HID.Value)
        Session.Remove("objClienteDXP" & HID.Value)
        Session.Remove("objPesagem" & HID.Value)

        If tipo.Length = 0 Then
            ddlUnidadeDeNegocio.SelectedIndex = 0
            ddlEmpresa.Items.Clear()
        End If

        txtCliente.Text = ""
        txtCodigoCliente.Value = 0
        txtPedido.Text = ""
        txtCodigoPedido.Value = 0

        txtPrimeiraPesagem.Text = 0
        txtSegundaPesagem.Text = 0
        txtPesoBruto.Text = 0
        txtDesconto.Text = 0
        txtLiquido.Text = 0
        hConfirmar.Value = 0

        txtdata1agr.Text = New DateTime(Now.Year, Now.Month, 1).ToString("dd/MM/yyyy")
        txtdata2agr.Text = Now().ToString("dd/MM/yyyy")

        gridAgrupamentoDePesagem.DataSource = Nothing
        gridAgrupamentoDePesagem.DataBind()

        gridDescontos.DataSource = Nothing
        gridDescontos.DataBind()

        HID.Value = Guid.NewGuid().ToString()

        objPesagem = New [Lib].Negocio.Pesagem()
        objPesagem.IUD = "I"
        objPesagem.CodigoSituacao = 1
        objPesagem.UsuarioInclusao = Session("ssNomeUsuario")

        If ddlEmpresa.Items.Count > 0 Then
            Dim strEmpresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            objPesagem.CodigoEmpresa = strEmpresa(0)
            objPesagem.EnderecoEmpresa = strEmpresa(1)
        End If

        SalvarObjetoSessao()

        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
            ddlUnidaDeNegocioDes.Enabled = False
            ddlEmpresaDes.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Not Session("objClienteAXP" & HID.Value) Is Nothing Then
                CarregarObjetoSessao()
                objPesagem.CodigoCliente = CType(Session("objClienteAXP" & HID.Value), [Lib].Negocio.Cliente).Codigo
                objPesagem.EnderecoCliente = CType(Session("objClienteAXP" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
                SalvarObjetoSessao()
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteAXP" & HID.Value), [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                txtPedido.Text = ""
                txtCodigoPedido.Value = 0
                Session.Remove("objClienteAXP" & HID.Value)
                BuscarPedidos()
            ElseIf Not Session("objClienteDXP" & HID.Value) Is Nothing Then
                CarregarObjetoSessao()
                objPesagem.CodigoCliente = CType(Session("objClienteDXP" & HID.Value), [Lib].Negocio.Cliente).Codigo
                objPesagem.EnderecoCliente = CType(Session("objClienteDXP" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
                SalvarObjetoSessao()
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteDXP" & HID.Value), [Lib].Negocio.Cliente))
                txtClienteDes.Text = itemCliente.Text
                txtCodigoClienteDes.Value = itemCliente.Value
                Session.Remove("objClienteDXP" & HID.Value)
            ElseIf Not Session("objPedidoAxP" & HID.Value) Is Nothing Then
                CarregarObjetoSessao()
                If CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).CodigoUnidadeNegocio <> ddlUnidadeDeNegocio.SelectedValue Then
                    MsgBox(Me.Page, "Unidade de Negócio da Empresa do Pedido é diferente da Unidade de Negócio da Empresa selecionada.")
                ElseIf CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).CodigoEmpresa <> objPesagem.CodigoEmpresa Or CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).EnderecoEmpresa <> objPesagem.EnderecoEmpresa Then
                    MsgBox(Me.Page, "Empresa do Pedido é diferente da Empresa Fornecedora.")
                ElseIf CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).CodigoCliente <> objPesagem.CodigoCliente Or CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).EnderecoCliente <> objPesagem.EnderecoCliente Then
                    MsgBox(Me.Page, "Fornecedor do Pedido é diferente do Fornecedor informado.")
                ElseIf CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).Itens.Count <> 1 Then
                    MsgBox(Me.Page, "Pedido com mais de um item não pode ser usado para agrupamento.")
                Else
                    objPesagem.CodigoPedido = CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).Codigo
                    objPesagem.CodigoTabelaDeClassificacao = CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).Itens(0).Classificacao.Codigo
                    objPesagem.CodigoProduto = CType(Session("objPedidoAxP" & HID.Value), [Lib].Negocio.Pedido).Itens(0).CodigoProduto
                    SalvarObjetoSessao()
                    txtPedido.Text = objPesagem.CodigoPedido
                    txtCodigoPedido.Value = objPesagem.CodigoPedido
                End If
                Session.Remove("objPedidoAxP" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BuscarPedidos()
        HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
        Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
        Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
        Dim parameters As New Dictionary(Of String, Object)
        HttpContext.Current.Session("ssTipoRetorno") = "objPedidoAxP" & HID.Value
        parameters("unidade") = ddlUnidadeDeNegocio.SelectedValue
        parameters("empresa") = strEmpresa(0)
        parameters("enderecoEmpresa") = strEmpresa(1)
        parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
        parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
        ucConsultaPedidos.Limpar()
        ucConsultaPedidos.BindGridView(parameters)
        Popup.ConsultaDePedidos(Me.Page, "objPedidoAxP" & HID.Value, "txtNome")
    End Sub

    Function Validar_Selecao() As Boolean
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCliente.Text) Then
            MsgBox(Me.Page, "Cliente não foi selecionado.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtPedido.Text) Then
            MsgBox(Me.Page, "Pedido não foi selecionado.")
            Return False
        End If
        Return True
    End Function

    Private Sub CarregaSelecionados()
        For Each row As GridViewRow In gridAgrupamentoDePesagem.Rows
            Dim chkLaudoDePesagem As CheckBox = CType(row.FindControl("chkPesagem"), CheckBox)
            CarregarObjetoSessao()

            If chkLaudoDePesagem.Checked Then
                CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)(row.RowIndex).SelecionaPesagem = True

                For i = 0 To CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)(row.RowIndex).Analises.Count - 1
                    For j = 0 To objPesagem.Analises.Count - 1
                        If objPesagem.Analises(j).CodigoAnalise = CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)(row.RowIndex).Analises(i).CodigoAnalise Then
                            objPesagem.Analises(j).Desconto = objPesagem.Analises(j).Desconto + CType(Session("objListPesagem" & HID.Value), [Lib].Negocio.ListPesagem)(row.RowIndex).Analises(i).Desconto
                        End If
                    Next
                Next
            End If
            SalvarObjetoSessao()
        Next
    End Sub

    Private Sub Imprimir_Pesagem(ByVal Empresa As String, ByVal EndEmpresa As String, ByVal CodigoPesagem As String, ByVal Tipo As String, ByVal Romaneio As String)
        Try
            Dim ds As DataSet = getDataSet(Empresa, EndEmpresa, CodigoPesagem)

            Dim param As New Dictionary(Of String, Object)
            param.Add("Reemissao", Tipo)
            param.Add("Romaneio", Romaneio)

            Funcoes.BindReport(Me.Page, ds, "Cr_LaudoDePesagem", eExportType.PDF, param)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet(ByVal Empresa As String, ByVal EndEmpresa As String, ByVal CodigoPesagem As String) As DataSet

        Dim ds As New DataSet()

        Dim sql As String = "SELECT   Clientes.Numero AS NumeroCliente, " & vbCrLf & _
            "           Pesagem.Pesagem_Id AS Laudo, " & vbCrLf & _
            "           Pesagem.Produto, Pesagem.Placa, " & vbCrLf & _
            "           Pesagem.EntradaSaida, " & vbCrLf & _
            "           Pesagem.PrimeiraPesagem, " & vbCrLf & _
            "           Pesagem.SegundaPesagem, " & vbCrLf & _
            "           Pesagem.BrutoBalanca, " & vbCrLf & _
            "           Pesagem.BrutoBalanca - Pesagem.Liquido AS Descontos, " & vbCrLf & _
            "           Pesagem.Liquido, " & vbCrLf & _
            "           Pesagem.EntradaPatio, " & vbCrLf & _
            "           Pesagem.EntradaBalanca, " & vbCrLf & _
            "           Pesagem.SaidaBalanca, " & vbCrLf & _
            "           Pesagem.Movimento, " & vbCrLf & _
            "           Pesagem.NumeroDaNota AS NotaFiscal, " & vbCrLf & _
            "           Pesagem.SerieDaNota AS SerieNota, " & vbCrLf & _
            "           Pesagem.PesoFiscal, " & vbCrLf & _
            "           Pesagem.Observacoes, " & vbCrLf & _
            "           Produtos.Nome AS NomeProduto, " & vbCrLf & _
            "           Clientes.Cliente_Id AS CodigoCliente, " & vbCrLf & _
            "           Clientes.Endereco_Id AS EndCliente, " & vbCrLf & _
            "           Clientes.Nome AS NomeCliente, " & vbCrLf & _
            "           Clientes.Reduzido AS ReduzidoCliente, " & vbCrLf & _
            "           Clientes.Endereco AS EnderecoCliente, " & vbCrLf & _
            "           Clientes.Cidade AS CidadeCliente, " & vbCrLf & _
            "           Clientes.Estado AS EstadoCliente, " & vbCrLf & _
            "           Transportes.Cliente_Id AS CodigoTransportador, " & vbCrLf & _
            "           Transportes.Endereco_Id AS EndTransportador, " & vbCrLf & _
            "           Transportes.Nome AS NomeTransportador, " & vbCrLf & _
            "           Transportes.Reduzido AS ReduzidoTransportador, " & vbCrLf & _
            "           Transportes.Endereco AS EnderecoTransportador, " & vbCrLf & _
            "           Transportes.Cidade AS CidadeTransportador, " & vbCrLf & _
            "           Transportes.Estado AS EstadoTransportador, " & vbCrLf & _
            "           Depositos.Cliente_Id AS CodigoDeposito, " & vbCrLf & _
            "           Depositos.Endereco_Id AS EndDeposito, " & vbCrLf & _
            "           Depositos.Nome AS NomeDeposito, " & vbCrLf & _
            "           Depositos.Reduzido AS ReduzidoDeposito, " & vbCrLf & _
            "           Depositos.Endereco AS EnderecoDeposito, " & vbCrLf & _
            "           Depositos.Cidade AS CidadeDeposito, " & vbCrLf & _
            "           Depositos.Estado AS EstadoDeposito, " & vbCrLf & _
            "           Depositos.Inscricao AS InscricaoDeposito, " & vbCrLf & _
            "           Placas.Placa01, " & vbCrLf & _
            "           Placas.Placa02, " & vbCrLf & _
            "           Placas.Placa03, " & vbCrLf & _
            "           Placas.CidadePlaca, " & vbCrLf & _
            "           Placas.EstadoPlaca, " & vbCrLf & _
            "           Placas.NomeMotorista, " & vbCrLf & _
            "           Placas.CidadeMotorista, " & vbCrLf & _
            "           Placas.EstadoMotorista, " & vbCrLf & _
            "           Placas.Habilitacao, " & vbCrLf & _
            "           CASE    WHEN Placas.CpfMotorista IS NULL THEN '' " & vbCrLf & _
            "                   WHEN Placas.CpfMotorista = '' THEN '' " & vbCrLf & _
            "           ELSE SUBSTRING(Placas.CpfMotorista, 1, 3) + '.' + SUBSTRING(Placas.CpfMotorista, 4, 3) + '.' + SUBSTRING(Placas.CpfMotorista, 7, 3) + '-' + SUBSTRING(Placas.CpfMotorista, 10, 2) " & vbCrLf & _
            "           END AS CpfMotorista, " & vbCrLf & _
            "                   EstadoPlaca.Descricao AS NomeEstadoPlaca, " & vbCrLf & _
            "                   EstadoMotorista.Descricao AS NomeEstadoMotorista, " & vbCrLf & _
            "                   Pedido, " & vbCrLf & _
            "                   Depositos.Numero AS NumeroDeposito, " & vbCrLf & _
            "                   Depositos.Complemento AS ComplementoDeposito, " & vbCrLf & _
            "                   Depositos.Bairro AS BairroDeposito, " & vbCrLf & _
            "                   Clientes.Numero AS NumeroCliente, " & vbCrLf & _
            "                   Clientes.Complemento AS ComplementoCliente, " & vbCrLf & _
            "                   Clientes.Bairro AS BairroCliente, " & vbCrLf & _
            "                   Clientes.Inscricao AS InscricaoCliente, " & vbCrLf & _
            "                   Pesagem.Empresa_Id as CodigoEmpresa," & vbCrLf & _
            "                   Pesagem.EndEmpresa_id as EndEmpresa," & vbCrLf & _
            "                   Clientes.Nome as NomeEmpresa," & vbCrLf & _
            "                   Clientes.Reduzido as ReduzidoEmpresa," & vbCrLf & _
            "                   Clientes.Endereco as EnderecoEmpresa," & vbCrLf & _
            "                   Clientes.Cidade as CidadeEmpresa," & vbCrLf & _
            "                   Clientes.Estado as EstadoEmpresa," & vbCrLf & _
            "                   Clientes.Inscricao as InscricaoEmpresa," & vbCrLf & _
            "                   Clientes.Numero as NumeroEmpresa," & vbCrLf & _
            "                   Clientes.Complemento as ComplementoEmpresa," & vbCrLf & _
            "                   Clientes.Bairro as BairroEmpresa" & vbCrLf & _
            "   FROM Pesagem " & vbCrLf & _
            "       INNER JOIN Produtos " & vbCrLf & _
            "                   ON Pesagem.Produto COLLATE Latin1_General_CI_AS = Produtos.Produto_Id " & vbCrLf & _
            "       INNER JOIN Clientes " & vbCrLf & _
            "                   ON Pesagem.Cliente COLLATE Latin1_General_CI_AS = Clientes.Cliente_Id AND Pesagem.EndCliente = Clientes.Endereco_Id " & vbCrLf & _
            "       LEFT JOIN Clientes AS Transportes" & vbCrLf & _
            "                   ON Pesagem.Transportador COLLATE Latin1_General_CI_AS = Transportes.Cliente_Id " & vbCrLf & _
            "                       And Pesagem.EndTransportador = Transportes.Endereco_Id " & vbCrLf & _
            "       LEFT JOIN Clientes AS Empresa ON Pesagem.Empresa_Id = Empresa.Cliente_Id And Pesagem.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf & _
            "       INNER JOIN Clientes AS Depositos" & vbCrLf & _
            "                   ON Pesagem.Deposito COLLATE Latin1_General_CI_AS = Depositos.Cliente_Id " & vbCrLf & _
            "                       AND Pesagem.EndDeposito = Depositos.Endereco_Id " & vbCrLf & _
            "       INNER JOIN Placas " & vbCrLf & _
            "                   ON Pesagem.Placa COLLATE Latin1_General_CI_AS = Placas.Placa_Id " & vbCrLf & _
            "       INNER JOIN Estados AS EstadoPlaca " & vbCrLf & _
            "                   ON Placas.EstadoPlaca = EstadoPlaca.Estado_Id " & vbCrLf

        '"       INNER JOIN Estados AS EstadoMotorista" & vbCrLf & _
        '"                   ON Placas.EstadoMotorista = EstadoMotorista.Estado_Id " & vbcrlf & _

        sql &= "        inner join Clientes Motorista                           " & vbCrLf & _
            "               on Motorista.Cliente_Id = placas.CpfMotorista      " & vbCrLf & _
            "              and Motorista.Endereco_Id = placas.EndCpfMotorista  " & vbCrLf & _
            "       INNER JOIN Estados AS EstadoMotorista                      " & vbCrLf & _
            "             ON Motorista.Estado = EstadoMotorista.Estado_Id      " & vbCrLf & _
            "   WHERE (Pesagem.Empresa_Id = '" & Empresa & "') " & vbCrLf & _
            "       AND (Pesagem.EndEmpresa_Id = " & EndEmpresa & ") " & vbCrLf & _
            "       AND (Pesagem.Pesagem_Id = " & CodigoPesagem & ") " & vbCrLf & _
            "       AND (Pesagem.Sequencia_Id = " & 0 & ")" & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "Laudo")

        sql = "SELECT   PesagemXAnalises.Analise_Id AS Analise, " & vbCrLf & _
            "       case when PesagemXAnalises.Analise_Id = 6 then " & vbCrLf & _
            "               Analises.Descricao + ' ' + " & vbCrLf & _
            "               (SELECT Observacoes " & vbCrLf & _
            "                       FROM Classificacoes " & vbCrLf & _
            "                   WHERE (Tabela_Id = Pesagem.TabelaDeClassificacao) " & vbCrLf & _
            "                       AND (Analise_Id = 6) " & vbCrLf & _
            "                       AND (Indice = PesagemXAnalises.Indice) " & vbCrLf & _
            "                       AND (Produto_Id = '" & ds.Tables(0).Rows(0).Item("Produto") & "')) " & vbCrLf & _
            "       else    Analises.Descricao end AS Descricao, " & vbCrLf & _
            "               PesagemXAnalises.Percentual, " & vbCrLf & _
            "               PesagemXAnalises.Indice, " & vbCrLf & _
            "               PesagemXAnalises.Desconto " & vbCrLf & _
            "   FROM PesagemXAnalises " & vbCrLf & _
            "       INNER JOIN Analises" & vbCrLf & _
            "           ON PesagemXAnalises.Analise_Id = Analises.Analise_Id " & vbCrLf & _
            "       INNER JOIN Pesagem" & vbCrLf & _
            "           ON PesagemXAnalises.Empresa_Id = Pesagem.Empresa_Id " & vbCrLf & _
            "               AND PesagemXAnalises.EndEmpresa_Id = Pesagem.EndEmpresa_Id " & vbCrLf & _
            "               And PesagemXAnalises.Pesagem_Id = Pesagem.Pesagem_Id AND PesagemXAnalises.Sequencia_Id = Pesagem.Sequencia_Id " & vbCrLf & _
            "   WHERE (PesagemXAnalises.Empresa_Id = '" & Empresa & "') " & vbCrLf & _
            "       AND (PesagemXAnalises.EndEmpresa_Id = " & EndEmpresa & ") " & vbCrLf & _
            "       AND (PesagemXAnalises.Pesagem_Id = " & CodigoPesagem & ") " & vbCrLf & _
            "       AND (PesagemXAnalises.Sequencia_Id = " & 0 & ") " & vbCrLf & _
            "   ORDER BY PesagemXAnalises.Analise_Id" & vbCrLf

        ds.Merge(Banco.ConsultaDataSet(sql, "Analises"))

        Return ds
    End Function


    Private Sub LimparCamposDes()
        Session.Remove("objClienteDXP" & HID.Value)
        Session.Remove("objPedidoDxP" & HID.Value)
        Session.Remove("objListPesagem" & HID.Value)

        txtClienteDes.Text = ""
        txtCodigoClienteDes.Value = 0
        txtNumeroPesagem.Text = String.Empty

        objPesagem = New [Lib].Negocio.Pesagem()
        objPesagem.IUD = "I"
        objPesagem.CodigoSituacao = 1
        objPesagem.UsuarioInclusao = Session("ssNomeUsuario")
        SalvarObjetoSessao()

        txtData1.Text = New DateTime(Now.Year, Now.Month, 1).ToString("dd/MM/yyyy")
        txtData2.Text = Now().ToString("dd/MM/yyyy")

        grdDesagrupamento.DataSource = Nothing
        grdDesagrupamento.DataBind()

        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
    End Sub

    Function Validar_SelecaoDesag() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresaDes.SelectedValue) Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtClienteDes.Text) Then
            MsgBox(Me.Page, "Cliente não foi selecionado.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtData1.Text) <> String.IsNullOrWhiteSpace(txtData2.Text) Then
            MsgBox(Me.Page, "Se a data Inicial ou final for informada as correspondentes terão que ser preenchidas.")
            Return False
        ElseIf (Not String.IsNullOrWhiteSpace(txtData1.Text) AndAlso Not IsDate(txtData1.Text)) Then
            MsgBox(Me.Page, "Data inicial não é uma data válida.")
            Return False
        ElseIf (Not String.IsNullOrWhiteSpace(txtData2.Text) AndAlso Not IsDate(txtData2.Text)) Then
            MsgBox(Me.Page, "Data final não é uma data válida.")
            Return False
        End If
        Return True
    End Function

    Private Sub ConsultarAgrupado(ByVal tipo As String)
        If Validar_SelecaoDesag() Then
            CarregarObjetoSessao()
            If Not String.IsNullOrWhiteSpace(txtNumeroPesagem.Text) AndAlso CInt(txtNumeroPesagem.Text) > 0 Then
                objPesagem.Codigo = CInt(txtNumeroPesagem.Text)
            Else
                objPesagem.Codigo = 0
            End If

            Dim ListPesagem As New [Lib].Negocio.ListPesagem(objPesagem, True, txtData1.Text.Trim(), txtData2.Text.Trim(), "AGRUPAMENTO", True, False)
            If ListPesagem.ToArray().Count > 0 Then
                Session("objListPesagem" & HID.Value) = ListPesagem
                grdDesagrupamento.DataSource = ListPesagem.ToArray()
                grdDesagrupamento.DataBind()
            Else
                grdDesagrupamento.DataSource = New List(Of Object)
                grdDesagrupamento.DataBind()
                If tipo.Length = 0 Then MsgBox(Me.Page, "Sem Laudos Agrupados para esta pesquisa.")
            End If
        End If
    End Sub

#End Region

End Class