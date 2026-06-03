Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CessaoDeCredito
    Inherits BasePage

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("CessaoDeCredito", "ACESSAR") Then
                    ddl.Carregar(ddlEmpresaC, CarregarDDL.Tabela.Empresas)
                    Funcoes.VerificaEmpresa(ddlEmpresaC)
                    ddl.Carregar(ddlSafraC, CarregarDDL.Tabela.Safra)
                    ddl.Carregar(ddlSituacoes, CarregarDDL.Tabela.Situacao)

                    SetarEmpresa(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
                    txtMovimento.Text = Now.Date

                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaClientes.SetarHID(HID.Value)
                    ucConsultaEmpresas.SetarHID(HID.Value)
                    ucConsultaPedidos.SetarHID(HID.Value)
                    ucConsultaProcuracao.SetarHID(HID.Value)
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo") = "Livre"
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaPXE" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCedente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCedente.Click
        Try
            buscarCliente("PXCed")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedidoCedente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPedidoCedente.Click
        Try
            If Not String.IsNullOrWhiteSpace(txtCedente.Text) Then
                Dim parameters As New Dictionary(Of String, Object)
                parameters("empresa") = txtCodigoEmpresa.Value.Split("-")(0)
                parameters("enderecoEmpresa") = txtCodigoEmpresa.Value.Split("-")(1)

                Dim Cedente() = txtCodigoCedente.Value.ToString.Split("-")
                parameters("cliente") = Cedente(0)
                parameters("enderecoCliente") = Cedente(1)
                parameters("situacao") = 1
                parameters("AnoFinalSafra") = Now.Year
                Session("ssCampo" & HID.Value) = "Pedidos"
                Popup.ConsultaDePedidos(Me.Page, "objPedidoCedente", "txtPedido")
                ucConsultaPedidos.BindGridView(parameters)
            Else
                MsgBox(Me.Page, "Informe o cedente.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCessionario_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCessionario.Click
        Try
            If Not String.IsNullOrWhiteSpace(txtPedidoCedente.Text) Then
                buscarCliente("PXCes")
            Else
                MsgBox(Me.Page, "Informe os dados do cedente.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CessaoDeCredito", "GRAVAR") Then
                If ValidarCampos() Then
                    If Not String.IsNullOrWhiteSpace(txtCessaoDeCredito.Text) Then
                        MsgBox(Me.Page, "Apague o número da Cessão de Crédito.")
                    Else
                        Dim objProcuracao = New [Lib].Negocio.Procuracao()
                        objProcuracao.IUD = "I"
                        objProcuracao.CodigoEmpresa = txtCodigoEmpresa.Value.Split("-")(0)
                        objProcuracao.EnderecoEmpresa = txtCodigoEmpresa.Value.Split("-")(1)
                        objProcuracao.Documento = txtDocumento.Text.ToUpper()
                        objProcuracao.CodigoCedente = txtCodigoCedente.Value.Split("-")(0)
                        objProcuracao.EnderecoCedente = txtCodigoCedente.Value.Split("-")(1)
                        objProcuracao.CodigoPedidoCedente = txtPedidoCedente.Text
                        objProcuracao.CodigoCessionario = txtCodigoCessionario.Value.Split("-")(0)
                        objProcuracao.EnderecoCessionario = txtCodigoCessionario.Value.Split("-")(1)
                        objProcuracao.Movimento = txtMovimento.Text
                        objProcuracao.Quantidade = CStr(txtQuantidade.Text)
                        objProcuracao.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtObservacoes.Text).ToUpper()
                        objProcuracao.UsuarioInclusao = UsuarioServidor.NomeUsuario
                        objProcuracao.UsuarioInclusaoData = Now.Date.ToString("yyyy-MM-dd")

                        If Not VerificaSaldo(objProcuracao) Then Exit Sub
                        If objProcuracao.Salvar() Then LimparCampos(objProcuracao.IUD, objProcuracao.Codigo)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("CessaoDeCredito", "ALTERAR") Then
                If ValidarCampos() Then
                    If String.IsNullOrWhiteSpace(txtCessaoDeCredito.Text) Then
                        MsgBox(Me.Page, "Informe o número da Cessão de Crédito.")
                    Else
                        Dim objProcuracao = New [Lib].Negocio.Procuracao(txtCodigoEmpresa.Value.Split("-")(0), txtCodigoEmpresa.Value.Split("-")(1), txtCessaoDeCredito.Text)
                        If objProcuracao.Codigo > 0 Then
                            'Testes verde 31/01/25
                            'If objProcuracao.Realizado > 0 And objProcuracao.Realizado > Convert.ToDouble(txtQuantidade.Text.Replace(".", "")) Then
                            '    MsgBox(Me.Page, "Quantidade informada não pode ser menor que o valor de notas já vinculadas.")
                            '    'Else
                            objProcuracao.IUD = "U"
                            objProcuracao.Documento = txtDocumento.Text.ToUpper()
                            objProcuracao.Movimento = txtMovimento.Text
                            objProcuracao.Quantidade = Convert.ToDouble(txtQuantidade.Text.Replace(".", ""))
                            objProcuracao.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtObservacoes.Text).ToUpper()
                            objProcuracao.UsuarioAlteracao = HttpContext.Current.Session("ssNomeUsuario")
                            objProcuracao.UsuarioAlteracaoData = Now.Date


                            If objProcuracao.Salvar Then
                                LimparCampos(objProcuracao.IUD, objProcuracao.Codigo)
                            End If
                            'End If
                        Else
                            MsgBox(Me.Page, "Procuracao não localizada.")
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para atualizar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("CessaoDeCredito", "EXCLUIR") Then
                If String.IsNullOrWhiteSpace(txtEmpresa.Text) Then
                    MsgBox(Me.Page, "Informe a Empresa.")
                ElseIf String.IsNullOrWhiteSpace(txtCessaoDeCredito.Text) Then
                    MsgBox(Me.Page, "Informe o número da Cessão de Crédito.")
                Else
                    Dim objProcuracao = New [Lib].Negocio.Procuracao(txtCodigoEmpresa.Value.Split("-")(0), txtCodigoEmpresa.Value.Split("-")(1), txtCessaoDeCredito.Text)

                    If objProcuracao.Codigo > 0 Then
                        If objProcuracao.Realizado > 0 Then
                            MsgBox(Me.Page, "Cessão de Crédito com lançamento(s) realizado(s), não é permitido efetuar exclusão.")
                        Else
                            objProcuracao.IUD = "D"
                            objProcuracao.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtObservacoes.Text).ToUpper()
                            objProcuracao.UsuarioCancelamento = HttpContext.Current.Session("ssNomeUsuario")
                            objProcuracao.UsuarioCancelamentoData = Now.Date
                            If objProcuracao.Salvar Then
                                LimparCampos(objProcuracao.IUD, objProcuracao.Codigo)
                            End If
                        End If
                    Else
                        MsgBox(Me.Page, "Erro na localização da Cessão de Crédito.")
                    End If
                End If

            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos("", "")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("CessaoDeCredito", "LEITURA") Then
                If String.IsNullOrWhiteSpace(ddlEmpresaC.SelectedValue) Then
                    MsgBox(Me.Page, "Empresa não foi selecionada")
                Else
                    Dim parameters As New Dictionary(Of String, Object)
                    parameters("tipo") = "PxProc"
                    parameters("situacao") = ""
                    parameters("emp") = ddlEmpresaC.SelectedValue.Split("-")(0)
                    parameters("ende") = ddlEmpresaC.SelectedValue.Split("-")(1)
                    parameters("proc") = txtCessaoDeCreditoC.Text

                    If Not String.IsNullOrWhiteSpace(HDCedenteC.Value) Then
                        parameters("cliCedente") = HDCedenteC.Value.Split("-")(0)
                        parameters("endCedente") = HDCedenteC.Value.Split("-")(1)
                    Else
                        parameters("cliCedente") = ""
                        parameters("endCedente") = 0
                    End If

                    If Not String.IsNullOrWhiteSpace(HDCessionarioC.Value) Then
                        parameters("cliCessionario") = HDCessionarioC.Value.Split("-")(0)
                        parameters("endCessionario") = HDCessionarioC.Value.Split("-")(1)
                    Else
                        parameters("cliCessionario") = ""
                        parameters("endCessionario") = 0
                    End If

                    parameters("periodo") = IIf(chkUsarPeriodo.Checked AndAlso Not String.IsNullOrWhiteSpace(txtDataInicialC.Text) AndAlso String.IsNullOrWhiteSpace(txtDataFinalC.Text), "'" & txtDataInicialC.Text.ToSqlDate() & "' And '" & txtDataFinalC.Text.ToSqlDate() & "'", "")
                    parameters("safra") = IIf(Not String.IsNullOrWhiteSpace(ddlSafraC.SelectedValue), ddlSafraC.SelectedValue, "")
                    parameters("filtro") = IIf(rdPendentes.Checked, "P", IIf(rdLiquidados.Checked, "L", ""))
                    parameters("pedc") = ""

                    If ucConsultaProcuracao.BindGridView(parameters) Then
                        Popup.ConsultaDeProcuracao(Me, "objProcuracaoPxProc" & HID.Value.ToString)
                    Else
                        MsgBox(Me.Page, "Nenhum resultado encontrado.", eTitulo.Info)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Not String.IsNullOrWhiteSpace(txtEmpresa.Text) Then
                Dim ds As DataSet = getDataSetProcuracao()
                ds.Merge(getDataSetMovimentos())

                Funcoes.BindReport(Me.Page, ds, "Cr_Procuracao", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Informe a Empresa.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Procuracao")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub LimparCampos(ByVal evento As String, ByVal Numero As String)
        txtEmpresa.Text = String.Empty
        txtCodigoEmpresa.Value = String.Empty
        txtCessaoDeCredito.Text = String.Empty
        lblRealizado.Text = "0.0000"
        ddlSituacoes.SelectedValue = String.Empty
        ddlUsuarios.Items.Clear()
        txtCedente.Text = String.Empty
        txtCodigoCedente.Value = String.Empty
        txtPedidoCedente.Text = String.Empty
        txtCessionario.Text = String.Empty
        txtCodigoCessionario.Value = String.Empty

        txtMovimento.Text = Now.Date
        txtDocumento.Text = String.Empty
        txtQuantidade.Text = String.Empty
        txtObservacoes.Text = String.Empty

        btnEmpresa.Enabled = True
        txtCessaoDeCredito.Enabled = True

        grdMovimentos.DataBind()
        HabilitarBotoesEdicao(True)
        txtQuantidade.Enabled = True

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

        SetarEmpresa(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)

        If evento = "I" Then
            MsgBox(Me.Page, "Cessão de Crédito " & Numero & " incluída com Sucesso.", eTitulo.Sucess)
        ElseIf evento = "U" Then
            MsgBox(Me.Page, "Cessão de Crédito " & Numero & " atualizada com Sucesso.", eTitulo.Sucess)
        ElseIf evento = "D" Then
            MsgBox(Me.Page, "Cessão de Crédito " & Numero & " deletada com Sucesso.", eTitulo.Sucess)
        End If
    End Sub

    Private Sub SetarEmpresa(ByVal Empresa As String, ByVal EndEmpresa As String)
        Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa, EndEmpresa)

        Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmpresa)
        txtEmpresa.Text = itemEmpresa.Text
        txtCodigoEmpresa.Value = itemEmpresa.Value
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Not Session("objEmpresaPxE" & HID.Value) Is Nothing Then
                SetarEmpresa(CType(Session("objEmpresaPxE" & HID.Value), [Lib].Negocio.Cliente).Codigo, CType(Session("objEmpresaPxE" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco)
                Session.Remove("objEmpresaPxE" & HID.Value)
            ElseIf Not Session("objCedente" & HID.Value) Is Nothing Then
                Dim itemCedente As ListItem = Funcoes.FormatarListItemCliente(CType(obj, Cliente))
                txtCedente.Text = itemCedente.Text
                txtCodigoCedente.Value = itemCedente.Value
                Session.Remove("objCedente" & HID.Value)
            ElseIf Not Session("objCessionario" & HID.Value) Is Nothing Then
                Dim itemCessionario As ListItem = Funcoes.FormatarListItemCliente(CType(obj, Cliente))
                txtCessionario.Text = itemCessionario.Text
                txtCodigoCessionario.Value = itemCessionario.Value
                Session.Remove("objCessionario" & HID.Value)
            ElseIf Not Session("objCedenteC" & HID.Value) Is Nothing Then
                Dim itemCedenteC As ListItem = Funcoes.FormatarListItemCliente(CType(obj, Cliente))
                txtCedenteC.Text = itemCedenteC.Text
                HDCedenteC.Value = itemCedenteC.Value
                Session.Remove("objCedenteC" & HID.Value)
            ElseIf Not Session("objCessionarioC" & HID.Value) Is Nothing Then
                Dim itemCessionarioC As ListItem = Funcoes.FormatarListItemCliente(CType(obj, Cliente))
                txtCessionarioC.Text = itemCessionarioC.Text
                HDCessionarioC.Value = itemCessionarioC.Value
                Session.Remove("objCessionarioC" & HID.Value)
            ElseIf Not Session("objDeposito" & HID.Value) Is Nothing Then
                Dim itemCessionario As ListItem = Funcoes.FormatarListItemCliente(CType(obj, Cliente))
                txtCessionario.Text = itemCessionario.Text
                txtCodigoCessionario.Value = itemCessionario.Value
                Session.Remove("objDeposito" & HID.Value)
            ElseIf Not Session("objPedidoCedente") Is Nothing Then
                txtPedidoCedente.Text = CType(obj, [Lib].Negocio.Pedido).Codigo
                Session.Remove("objPedidoCedente")
            End If
            Page.Validate()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub buscarCliente(ByVal Tipo As String)
        If Tipo = "PXCed" Then
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCedente" & HID.Value, "txtNome")
        ElseIf Tipo = "PXCes" Then
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCessionario" & HID.Value, "txtNome")
        ElseIf Tipo = "PXCedC" Then
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCedenteC" & HID.Value, "txtNome")
        ElseIf Tipo = "PXCesC" Then
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCessionarioC" & HID.Value, "txtNome")
        Else
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objDeposito" & HID.Value, "txtNome")
        End If
    End Sub

    Private Function ValidarCampos() As Boolean
        If txtCodigoEmpresa.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        ElseIf txtCodigoCedente.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Cedente não foi selecionado.")
            Return False
        ElseIf txtPedidoCedente.Text.Length = 0 Then
            MsgBox(Me.Page, "Pedido do Cedente não foi selecionado.")
            Return False
        ElseIf txtCodigoCessionario.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Cessionário não foi selecionado.")
            Return False
        ElseIf txtDocumento.Text.Length = 0 Then
            MsgBox(Me.Page, "Documento não foi informado.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtMovimento.Text) OrElse Not IsDate(txtMovimento.Text) Then
            MsgBox(Me.Page, "Data do Movimento não foi informada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtQuantidade.Text) OrElse CDec(txtQuantidade.Text) = 0 Then
            MsgBox(Me.Page, "Quantidade não foi informada.")
            Return False
            'Testes verde 31/01/25
            'ElseIf CDec(txtQuantidade.Text) < CDec(lblRealizado.Text) Then
            '    MsgBox(Me.Page, "A quantidade informada não pode ser menor que a quantidade de Realizados.")
            '    Return False
        Else
            Return True
        End If
    End Function

    Private Function VerificaSaldo(ByVal objProcuracao As [Lib].Negocio.Procuracao) As Boolean
        Dim Parametros As New Hashtable
        Parametros.Add("Empresa", objProcuracao.CodigoEmpresa)
        Parametros.Add("EndEmpresa", objProcuracao.EnderecoEmpresa)
        Parametros.Add("Pedido", objProcuracao.CodigoPedidoCedente)
        Dim LevantamentoPedido As New ListSaldoPedido2015(Parametros)

        Dim EntreguePedido As Double = 0
        If LevantamentoPedido.Count > 0 Then
            If LevantamentoPedido(0).Pedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS OrElse LevantamentoPedido(0).Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then
                If LevantamentoPedido(0).Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then
                    EntreguePedido = LevantamentoPedido(0).QtdeEntregueFiscalAFixar
                Else
                    EntreguePedido = LevantamentoPedido(0).QtdeEntregueFiscalDeposito
                End If
            Else
                EntreguePedido = LevantamentoPedido(0).QtdeEntregueFiscalGlobal + LevantamentoPedido(0).QtdeEntregueFiscalDireta
            End If
        End If

        Dim SaldoProcuracoes As Double = New ListProcuracao(objProcuracao.CodigoPedidoCedente).QuantidadeSaldoProcuracoes

        'teste para verde 04/02/2025
        'If LevantamentoPedido(0).Tipo = 1 Then
        '    If objProcuracao.Quantidade > (LevantamentoPedido(0).QtdeContratadoFiscal - EntreguePedido - SaldoProcuracoes) Then
        '        MsgBox(Me.Page, "O Saldo da quantidade entregue do pedido menos as o saldo das procuracoes é igual a " & CStr(Math.Round(LevantamentoPedido(0).QtdeContratadoFiscal - EntreguePedido - SaldoProcuracoes, 4)) & " . E não é suficiente para gerar esta Cessão de Crédito.")
        '        Return False
        '    End If
        'End If
        Return True

    End Function

    Public Sub CarregarProcuracao()
        Try
            If Not Session("objProcuracaoPxProc" & HID.Value) Is Nothing Then
                Dim Empresa() As String = ddlEmpresaC.SelectedValue.Split("-")
                Dim objProcuracao = New [Lib].Negocio.Procuracao(Empresa(0), Empresa(1), Session("objProcuracaoPxProc" & HID.Value))
                Session.Remove("objProcuracaoPxProc" & HID.Value)

                Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objProcuracao.Empresa)
                txtEmpresa.Text = itemEmpresa.Text
                txtCodigoEmpresa.Value = itemEmpresa.Value
                txtCessaoDeCredito.Text = objProcuracao.Codigo

                Dim itemCedente As ListItem = Funcoes.FormatarListItemCliente(objProcuracao.Cedente)
                txtCedente.Text = itemCedente.Text
                txtCodigoCedente.Value = itemCedente.Value
                txtPedidoCedente.Text = objProcuracao.CodigoPedidoCedente

                Dim itemCessionario As ListItem = Funcoes.FormatarListItemCliente(objProcuracao.Cessionario)
                txtCessionario.Text = itemCessionario.Text
                txtCodigoCessionario.Value = itemCessionario.Value

                txtDocumento.Text = objProcuracao.Documento
                txtMovimento.Text = objProcuracao.Movimento
                txtQuantidade.Text = objProcuracao.Quantidade.ToString("N4")
                txtObservacoes.Text = objProcuracao.Observacoes
                ddlSituacoes.SelectedValue = objProcuracao.CodigoSituacao
                lblRealizado.Text = objProcuracao.Realizado.ToString("N4")
                grdMovimentos.DataSource = getDataSetMovimentos()
                grdMovimentos.DataBind()

                ddlUsuarios.Items.Clear()
                If Not String.IsNullOrWhiteSpace(objProcuracao.UsuarioInclusao) Then ddlUsuarios.Items.Add("Inc - " & objProcuracao.UsuarioInclusao)
                If Not String.IsNullOrWhiteSpace(objProcuracao.UsuarioAlteracao) Then ddlUsuarios.Items.Add("Alt - " & objProcuracao.UsuarioAlteracao)
                If Not String.IsNullOrWhiteSpace(objProcuracao.UsuarioCancelamento) Then ddlUsuarios.Items.Add("Canc - " & objProcuracao.UsuarioCancelamento)

                btnEmpresa.Enabled = False
                txtCessaoDeCredito.Enabled = False
                lnkNovo.Parent.Visible = False
                lnkConsultar.Parent.Visible = False

                If objProcuracao.Situacao.Codigo = 1 Then
                    lnkAtualizar.Parent.Visible = True
                    lnkExcluir.Parent.Visible = True
                Else
                    lnkAtualizar.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False
                End If

                If objProcuracao.Saldo > 0 Then
                    If objProcuracao.Realizado > 0 Then
                        MsgBox(Me.Page, "Algumas informações não podem ser alteradas, Cessão de Crédito contém movimentações.")
                        HabilitarBotoesEdicao(False)
                        lnkExcluir.Parent.Visible = False
                    End If
                Else
                    MsgBox(Me.Page, "Cessão de Crédito não pode ser alterada ou excluída - (Saldo = 0.00).")
                    If objProcuracao.Realizado > 0 Then
                        HabilitarBotoesEdicao(False)
                        lnkExcluir.Parent.Visible = False
                        lnkAtualizar.Parent.Visible = False
                        txtQuantidade.Enabled = False
                    End If

                End If
                Page.Validate()
                TabContainer1.ActiveTabIndex = 1
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub HabilitarBotoesEdicao(ByVal enable As Boolean)
        btnCedente.Enabled = enable
        btnPedidoCedente.Enabled = enable
        btnCessionario.Enabled = enable
        txtMovimento.Enabled = enable
        txtDocumento.Enabled = enable
    End Sub

    Private Function getDataSetProcuracao() As DataSet
        Dim sql As String = "Select pr.Empresa_Id as CodEmpresa, pr.EndEmpresa_Id as EndEmpresa, emp.Nome as NomeEmpresa, emp.Cidade as CidadeEmpresa," & vbCrLf &
                            "       emp.Estado as EstadoEmpresa, pr.Procuracao_ID as Procuracao, pr.Documento, pr.Cedente as CodCedente, pr.EndCedente," & vbCrLf &
                            "       ced.Nome as NomeCedente, pr.PedidoCedente, pr.Cessionario as CodCessionario, pr.EndCessionario, ces.Nome as NomeCessionario," & vbCrLf &
                            "       pr.Movimento, pr.Quantidade, mov.Realizado, pr.Quantidade - mov.Realizado as Saldo" & vbCrLf &
                            "  From Procuracoes pr" & vbCrLf &
                            " INNER Join(" & vbCrLf &
                            "          	 Select Empresa_Id, EndEmpresa_Id, Procuracao, SUM(Realizado) as Realizado" & vbCrLf &
                            "              From (" & vbCrLf &
                            "          		 	 SELECT nf.Empresa_Id, nf.EndEmpresa_Id, nf.Procuracao," & vbCrLf &
                            "                   	    SUM(nfi.QuantidadeFisica) AS Realizado" & vbCrLf &
                            "                 	   FROM NotasFiscais NF" & vbCrLf &
                            "                     INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                            "                        ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                            "                       AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                            "                       AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                            "                       AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                            "                       AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                            "                       AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                            "                       AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                            "                     Inner join SubOperacoes SO" & vbCrLf &
                            "          			     on so.Operacao_Id     = NF.Operacao" & vbCrLf &
                            "                   	and so.SubOperacoes_Id = NF.SubOperacao" & vbCrLf &
                            "                     Where so.Classe <> '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'" & vbCrLf &
                            "                       and nf.situacao in (1,4,7)" & vbCrLf &
                            "                     GROUP BY nf.Empresa_Id, nf.EndEmpresa_Id, nf.Procuracao" & vbCrLf &
                            "          		 	  Union All" & vbCrLf &
                            "          		 	 select p.Empresa_Id, p.EndEmpresa_Id, pif.Procuracao," & vbCrLf &
                            "          			        sum(pif.Quantidade) as Realizado" & vbCrLf &
                            "          			   from Pedidos P" & vbCrLf &
                            "          			  Inner join PedidosXItensXFixacoes PIF" & vbCrLf &
                            "          			     on P.Empresa_Id    = PIF.Empresa_Id" & vbCrLf &
                            "                       and P.EndEmpresa_Id = PIF.EndEmpresa_Id" & vbCrLf &
                            "          			    and p.Pedido_Id     = PIF.Pedido_Id" & vbCrLf &
                            "          		      Where P.Situacao      = 1" & vbCrLf &
                            "          			  Group by p.Empresa_Id, p.EndEmpresa_Id, pif.Procuracao" & vbCrLf &
                            "          			 ) as submov" & vbCrLf &
                            "          	    Group by submov.Empresa_Id, submov.EndEmpresa_Id, submov.Procuracao" & vbCrLf &
                            "          	  ) as mov" & vbCrLf &
                            "    on mov.Empresa_Id    = pr.Empresa_Id" & vbCrLf &
                            "   and mov.EndEmpresa_Id = pr.EndEmpresa_Id" & vbCrLf &
                            "   and mov.Procuracao    = pr.Procuracao_ID" & vbCrLf &
                            " Inner Join Clientes emp" & vbCrLf &
                            "    on emp.Cliente_Id  = pr.Empresa_Id" & vbCrLf &
                            "   and emp.Endereco_Id = pr.EndEmpresa_Id" & vbCrLf &
                            " Inner Join Clientes ced" & vbCrLf &
                            "    on ced.Cliente_Id  = pr.Cedente" & vbCrLf &
                            "   and ced.Endereco_Id = pr.EndCedente" & vbCrLf &
                            " Inner Join Clientes ces" & vbCrLf &
                            "    on ces.Cliente_Id  = pr.Cessionario" & vbCrLf &
                            "   and ces.Endereco_Id = pr.EndCessionario" & vbCrLf &
                            " Inner Join Clientes dep" & vbCrLf &
                            "    on dep.Cliente_Id  = pr.Deposito" & vbCrLf &
                            "   and dep.Endereco_Id = pr.EndDeposito" & vbCrLf &
                            " Inner Join Produtos prod" & vbCrLf &
                            "    on prod.produto_id = pr.produto" & vbCrLf &
                            " Where pr.Empresa_Id    ='" & txtCodigoEmpresa.Value.Split("-")(0) & "'" & vbCrLf &
                            "   and pr.EndEmpresa_Id = " & txtCodigoEmpresa.Value.Split("-")(1) & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCessaoDeCredito.Text) Then sql &= "   and pr.Procuracao_ID = " & txtCessaoDeCredito.Text & vbCrLf

        Return Banco.ConsultaDataSet(sql, "Procuracoes")
    End Function

    Private Function getDataSetMovimentos() As DataSet
        Dim sql As String = "select prc.Empresa_Id as CodEmpresa, prc.EndEmpresa_Id as EndEmpresa, prc.Procuracao_ID as Procuracao," & vbCrLf &
                            "       mov.Descricao, mov.Movimento, mov.Realizado" & vbCrLf &
                            "  from Procuracoes prc" & vbCrLf &
                            " INNER Join(" & vbCrLf &
                            "			SELECT nf.Empresa_Id, nf.EndEmpresa_Id, nf.Procuracao," & vbCrLf &
                            "			       ('NotaFiscal: ' + cast(nf.Nota_Id as varchar) + '-' + nf.Serie_Id) as Descricao," & vbCrLf &
                            "                  nf.Movimento, SUM(nfi.QuantidadeFisica) AS Realizado" & vbCrLf &
                            "             FROM NotasFiscais NF" & vbCrLf &
                            "            INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                            "               ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                            "              AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                            "              AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                            "              AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                            "              AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                            "              AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                            "              AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                            "            Inner join SubOperacoes SO" & vbCrLf &
                            "			    on so.Operacao_Id     = NF.Operacao" & vbCrLf &
                            "              and so.SubOperacoes_Id = NF.SubOperacao" & vbCrLf &
                            "            Where so.Classe <> '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'" & vbCrLf &
                            "              and nf.situacao in (1,4,7)" & vbCrLf &
                            "            GROUP BY nf.Empresa_Id, nf.EndEmpresa_Id, nf.Procuracao," & vbCrLf &
                            "                    ('NotaFiscal: ' + cast(nf.Nota_Id as varchar) + '-' + nf.Serie_Id), nf.Movimento" & vbCrLf &
                            "			 Union All" & vbCrLf &
                            "			Select p.Empresa_Id, p.EndEmpresa_Id, pif.Procuracao," & vbCrLf &
                            "                  'Fixacao Número: ' + cast(pif.Fixacao_Id as varchar), pif.Movimento," & vbCrLf &
                            "				   sum(pif.Quantidade) as Realizado" & vbCrLf &
                            "			  from Pedidos P" & vbCrLf &
                            "            Inner join PedidosXItensXFixacoes PIF" & vbCrLf &
                            "			    on P.Empresa_Id    = PIF.Empresa_Id" & vbCrLf &
                            "              and P.EndEmpresa_Id = PIF.EndEmpresa_Id" & vbCrLf &
                            "			   and p.Pedido_Id     = PIF.Pedido_Id" & vbCrLf &
                            "		     Where P.Situacao      = 1" & vbCrLf &
                            "			 Group by p.Empresa_Id, p.EndEmpresa_Id, pif.Procuracao," & vbCrLf &
                            "			          ('Fixacao Número: ' + cast(pif.Fixacao_Id as varchar)), pif.Movimento" & vbCrLf &
                            "			 ) as mov" & vbCrLf &
                            "	 on mov.Empresa_Id    = prc.Empresa_Id" & vbCrLf &
                            "   and mov.EndEmpresa_Id = prc.EndEmpresa_Id" & vbCrLf &
                            "   and mov.Procuracao    = prc.Procuracao_ID" & vbCrLf &
                            " Where mov.Empresa_Id = '" & txtCodigoEmpresa.Value.Split("-")(0) & "'" & vbCrLf &
                            "   and mov.EndEmpresa_Id = " & txtCodigoEmpresa.Value.Split("-")(1) & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCessaoDeCredito.Text) Then sql &= "      and mov.Procuracao = " & txtCessaoDeCredito.Text & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Movimentos")
        Return ds
    End Function

    Protected Sub btnCendenteC_Click(sender As Object, e As EventArgs) Handles btnCendenteC.Click
        buscarCliente("PXCedC")
    End Sub

    Protected Sub btnCessionarioC_Click(sender As Object, e As EventArgs) Handles btnCessionarioC.Click
        buscarCliente("PXCesC")
    End Sub

    Protected Sub chkUsarPeriodo_CheckedChanged(sender As Object, e As EventArgs) Handles chkUsarPeriodo.CheckedChanged
        divPeriodo.Visible = chkUsarPeriodo.Checked
    End Sub

#End Region

End Class