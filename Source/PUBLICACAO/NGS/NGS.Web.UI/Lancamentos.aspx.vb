Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class Lancamentos
    Inherits BasePage

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then
            Try
                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Contabil.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("Lancamentos", "ACESSAR") Then

                    ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)

                    ddl.Carregar(ddlHistorico, CarregarDDL.Tabela.HistoricoContabil)
                    ddl.Carregar(ddlLote, CarregarDDL.Tabela.LoteContabil)
                    ddl.Carregar(ddlCustoDebito, CarregarDDL.Tabela.CentroDeCusto, "len(CentroDeCusto_ID) = 5")
                    ddl.Carregar(ddlCustoCredito, CarregarDDL.Tabela.CentroDeCusto, "len(CentroDeCusto_ID) = 5")
                    ddl.Carregar(DdlGrupo, CarregarDDL.Tabela.GrupoProduto, "", True)
                    Limpar()
                    txtMovimento.Text = Today.ToString("dd/MM/yyyy")
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Contabil.aspx")
                End If
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message)
            End Try
        End If
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            'Limpar()
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEmpresa.SelectedIndexChanged
        Try
            'Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnContasDebito_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnContasDebito.Click
        Dim txtConta As TextBox = ucConsultaPlanoDeContas.FindControl("txtConta")
        Try
            If txtConta IsNot Nothing Then
                txtConta.Text = String.Empty
                txtConta.Focus()
            End If
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me, "objDebitoLancamento" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnClientesDebito_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClientesDebito.Click
        Try
            If String.IsNullOrWhiteSpace(txtCodigoContaDebito.Value) Then
                MsgBox(Me.Page, "Informe a conta de débito.")
            Else

                Dim objContaDebito As PlanoDeConta = New PlanoDeConta("", 0, txtCodigoContaDebito.Value)
                If objContaDebito IsNot Nothing AndAlso Not objContaDebito.TemCliente Then
                    MsgBox(Me.Page, "Conta de débito não permite cliente.")
                Else
                    ucConsultaClientes.Limpar()
                    ucConsultaClientes.SetarTipoCliente("3,4,5,6,7,8,9,10")
                    Popup.ConsultaDeClientes(Me, "objClienteLnctoDEB" & HID.Value, "txtNome")
                End If

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlCustoDebito_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim obj As New PlanoDeConta("", 0, txtCodigoContaDebito.Value)
            If String.IsNullOrWhiteSpace(txtCodigoContaDebito.Value) Then
                MsgBox(Me.Page, "É necessário informar a conta de débito.")
                ddlCustoDebito.SelectedIndex = 0
                Exit Sub
            ElseIf ddlCustoDebito.SelectedValue.Length > 0 AndAlso Not ValidaCentroDeCusto(txtCodigoContaDebito.Value) Then
                ddlCustoDebito.SelectedIndex = 0
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnContasCredito_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnContasCredito.Click
        Dim txtConta As TextBox = ucConsultaPlanoDeContas.FindControl("txtConta")
        Try
            If txtConta IsNot Nothing Then
                txtConta.Text = String.Empty
                txtConta.Focus()
            End If
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me.Page, "objCreditoLancamento" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnClientesCredito_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClientesCredito.Click
        Try
            If String.IsNullOrWhiteSpace(txtCodigoContaCredito.Value) Then
                MsgBox(Me.Page, "Informe a conta de Crédito.")
            Else
                Dim objContaCredito As PlanoDeConta = New PlanoDeConta("", 0, txtCodigoContaCredito.Value)
                If objContaCredito IsNot Nothing AndAlso Not objContaCredito.TemCliente Then
                    MsgBox(Me.Page, "Conta de crédito não permite cliente.")
                Else
                    ucConsultaClientes.Limpar()
                    ucConsultaClientes.SetarTipoCliente("3,4,5,6,7,8,9,10")
                    Popup.ConsultaDeClientes(Me, "objClienteLnctoCRE" & HID.Value, "txtNome")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlCustoCredito_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim obj As New PlanoDeConta("", 0, txtCodigoContaCredito.Value)
            If String.IsNullOrWhiteSpace(txtCodigoContaCredito.Value) Then
                MsgBox(Me.Page, "É necessário informar a conta de crédito.")
                ddlCustoCredito.SelectedIndex = 0
                Exit Sub
            ElseIf ddlCustoCredito.SelectedValue.Length > 0 AndAlso Not ValidaCentroDeCusto(txtCodigoContaCredito.Value) Then
                ddlCustoCredito.SelectedIndex = 0
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlGrupo.SelectedIndexChanged
        Try
            If CType(sender, DropDownList).SelectedIndex > 0 Then
                Dim objContaDebito As New PlanoDeConta("", 0, txtCodigoContaDebito.Value)
                Dim objContaCredito As New PlanoDeConta("", 0, txtCodigoContaCredito.Value)

                If (objContaDebito IsNot Nothing OrElse objContaCredito IsNot Nothing) Then
                    If (objContaDebito IsNot Nothing AndAlso objContaDebito.TemProduto) OrElse (objContaCredito IsNot Nothing AndAlso objContaCredito.TemProduto) Then
                        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, "Grupo = '" & DdlGrupo.SelectedValue & "'", True)
                    Else
                        DdlGrupo.SelectedIndex = 0
                        MsgBox(Me.Page, "Conta(s) selecionada(s) não permite produto.")
                    End If
                Else
                    DdlGrupo.SelectedIndex = 0
                    MsgBox(Me.Page, "Não há nenhuma conta selecionada.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            Dim ContaDebito As PlanoDeConta = New PlanoDeConta("", 0, txtCodigoContaDebito.Value)
            Dim ContaCredito As PlanoDeConta = New PlanoDeConta("", 0, txtCodigoContaCredito.Value)

            If ddlUnidade.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Unidade de negócio não foi selecionada.")
            ElseIf ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            ElseIf ContaDebito IsNot Nothing AndAlso ContaDebito.TemCliente AndAlso String.IsNullOrWhiteSpace(txtCodigoClienteDebito.Value) Then
                MsgBox(Me.Page, "Cliente crédito deve ser selecionado.")
            ElseIf ContaCredito IsNot Nothing AndAlso ContaCredito.TemCliente AndAlso String.IsNullOrWhiteSpace(txtCodigoClienteCredito.Value) Then
                MsgBox(Me.Page, "Cliente crédito deve ser selecionado.")
            ElseIf ddlProduto.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Produto não foi selecionado.")
            Else

                Dim Cliente(1) As String

                If ContaCredito IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(txtCodigoClienteCredito.Value) Then
                    Cliente(0) = txtCodigoClienteCredito.Value.Split("-")(0)
                    Cliente(1) = txtCodigoClienteCredito.Value.Split("-")(1)
                ElseIf ContaDebito IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(txtCodigoClienteDebito.Value) Then
                    Cliente(0) = txtCodigoClienteDebito.Value.Split("-")(0)
                    Cliente(1) = txtCodigoClienteDebito.Value.Split("-")(1)
                End If

                Session("ssCampo" & HID.Value) = "Pedidos"
                Dim parameters As New Dictionary(Of String, Object)
                parameters("unidade") = ddlUnidade.SelectedValue
                parameters("empresa") = Empresa(0)
                parameters("enderecoEmpresa") = Empresa(1)
                parameters("cliente") = IIf(Cliente(0) <> "", Cliente(0), "")
                parameters("enderecoCliente") = IIf(Cliente.Length > 1, Cliente(1), "")
                parameters("CodigoProduto") = ddlProduto.SelectedValue

                Popup.ConsultaDePedidos(Me.Page, "objPedido" & HID.Value, "txtNome")
                ucConsultaPedidos.BindGridView(parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtMovimento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMovimento.TextChanged
        Try
            ddlLote.SelectedIndex = 0
            If ddlSequencia.Text <> "" Then ddlSequencia.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlLote_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlLote.SelectedIndexChanged
        Try
            If ddlSequencia.Items.Count > 0 Then ddlSequencia.SelectedIndex = 0
            TotalizaLote()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlSequencia_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlSequencia.SelectedIndexChanged
        Try
            If Not String.IsNullOrWhiteSpace(ddlSequencia.Text) Then
                Consultar()
            Else
                MsgBox(Me.Page, "Selecione uma sequencia para a consulta.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Lancamentos", "GRAVAR") Then
                If ValidarCampos() Then
                    If ddlSequencia.Text <> "" Then
                        Excluir()
                    End If
                    Gravar()
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
            If Funcoes.VerificaPermissao("Lancamentos", "ALTERAR") Then
                If ValidarCampos() Then
                    Excluir()
                    Gravar()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Lancamentos", "EXCLUIR") Then
                Excluir()
                Limpar(False)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            Consultar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar(True)
            If ddlSequencia.Text <> "" Then
                ddlSequencia.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Lancamentos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Methods/functions"

    Private Sub Limpar(Optional ByVal Tudo As Boolean = True)

        If Tudo Then 'Usado para que tenhamos a possibilidade de não limpar os campos necessários para consulta'
            Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
            txtMovimento.Text = Today.ToString("dd/MM/yyyy")
            ddlLote.SelectedValue = 1
            ddlSequencia.SelectedValue = String.Empty
        End If

        lblUsuario.Text = UsuarioServidor.NomeUsuario

        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkNovo.Parent.Visible = True

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPlanoDeContas.SetarHID(HID.Value)

        ChkAproveitarDados.Checked = False

        'txtTemClienteDebito.Value = False
        txtCodigoContaDebito.Value = String.Empty
        txtContaDebito.Text = String.Empty
        txtCodigoClienteDebito.Value = String.Empty
        txtClienteDebito.Text = String.Empty

        ddlCustoDebito.SelectedValue = String.Empty

        'txtTemClienteCredito.Value = False
        txtCodigoContaCredito.Value = String.Empty
        txtContaCredito.Text = String.Empty
        txtCodigoClienteCredito.Value = String.Empty
        txtClienteCredito.Text = String.Empty

        ddlCustoCredito.SelectedValue = String.Empty
        DdlGrupo.SelectedValue = String.Empty
        ddlProduto.Items.Clear()

        txtPedido.Text = String.Empty
        txtValor.Text = String.Empty
        txtQuantidade.Text = String.Empty

        ddlHistorico.SelectedValue = String.Empty
        txtComplemento.Text = String.Empty

        lblTotalDebitos.Text = "0"
        lblTotalCreditos.Text = "0"
        lblDiferenca.Text = "0"

        btnContasCredito.Enabled = True
        btnClientesCredito.Enabled = True
        ddlCustoCredito.Enabled = True

        btnContasDebito.Enabled = True
        btnClientesDebito.Enabled = True
        ddlCustoDebito.Enabled = True
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objDebitoLancamento" & HID.Value) Is Nothing Then
            Dim objPlanoConta As PlanoDeConta = CType(obj, [Lib].Negocio.PlanoDeConta)
            txtCodigoContaDebito.Value = objPlanoConta.Conta
            txtContaDebito.Text = objPlanoConta.Conta & " - " & objPlanoConta.Titulo
            'txtTemClienteDebito.Value = objPlanoConta.TemCliente

            txtClienteDebito.Text = String.Empty
            txtCodigoClienteDebito.Value = String.Empty

            If objPlanoConta.TemCliente Then
                ucConsultaClientes.Limpar()
                Popup.ConsultaDeClientes(Me, "objClienteLnctoDEB" & HID.Value, "txtNome")
            End If
            Session.Remove("objDebitoLancamento" & HID.Value)

        ElseIf Session("objClienteLnctoDEB" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(obj, [Lib].Negocio.Cliente))
            txtCodigoClienteDebito.Value = itemCliente.Value
            txtClienteDebito.Text = itemCliente.Text
            Session.Remove("objClienteLnctoDEB" & HID.Value)

        ElseIf Not Session("objCreditoLancamento" & HID.Value) Is Nothing Then
            Dim objPlanoConta As PlanoDeConta = CType(obj, [Lib].Negocio.PlanoDeConta)
            txtCodigoContaCredito.Value = objPlanoConta.Conta
            txtContaCredito.Text = objPlanoConta.Conta & " - " & objPlanoConta.Titulo
            'txtTemClienteCredito.Value = objPlanoConta.TemCliente

            txtClienteCredito.Text = String.Empty
            txtCodigoClienteCredito.Value = String.Empty

            If objPlanoConta.TemCliente Then
                ucConsultaClientes.Limpar()
                Popup.ConsultaDeClientes(Me, "objClienteLnctoCRE" & HID.Value, "txtNome")
            End If
            Session.Remove("objCreditoLancamento" & HID.Value)

        ElseIf Session("objClienteLnctoCRE" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(obj, [Lib].Negocio.Cliente))
            txtCodigoClienteCredito.Value = itemCliente.Value
            txtClienteCredito.Text = itemCliente.Text
            Session.Remove("objClienteLnctoCRE" & HID.Value)

        ElseIf Session("objPedido" & HID.Value) IsNot Nothing Then
            txtPedido.Text = CType(obj, Pedido).Codigo
            Session.Remove("objPedido" & HID.Value)
        End If

    End Sub

    Function ValidarCampos() As Boolean
        Dim Empresa = ddlEmpresa.SelectedValue.Split("-")

        If ddlLote.SelectedValue = "" Then
            MsgBox(Me.Page, "Lote é obrigatório.")
            Return False
        ElseIf ddlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            Return False
        ElseIf ddlEmpresa.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf txtMovimento.Text = "" Then
            MsgBox(Me.Page, "Movimento é obrigatório.")
            Return False
        ElseIf Not Funcoes.VerificaAcesso(Empresa(0), Empresa(1), txtMovimento.Text, "CONTABIL") Then
            MsgBox(Me.Page, "Movimento já fechado para esta data.")
            Return False
        ElseIf txtValor.Text = "" OrElse CDbl(txtValor.Text) = 0 Then
            MsgBox(Me.Page, "Valor é Obrigatório.")
            Return False
        ElseIf ddlHistorico.SelectedValue = "" And txtComplemento.Text = "" Then
            MsgBox(Me.Page, "Informe o histórico ou o complemento de histórico.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCodigoContaDebito.Value) And String.IsNullOrWhiteSpace(txtCodigoContaCredito.Value) Then
            MsgBox(Me.Page, "Conta de débito ou crédito é obrigatório.")
            Return False
        End If

        Dim objContaDebito As PlanoDeConta = New PlanoDeConta("", 0, txtCodigoContaDebito.Value)
        Dim objContaCredito As PlanoDeConta = New PlanoDeConta("", 0, txtCodigoContaCredito.Value)

        If objContaDebito IsNot Nothing Then
            If objContaDebito.TemCliente AndAlso String.IsNullOrWhiteSpace(txtCodigoClienteDebito.Value) Then
                MsgBox(Me.Page, "Selecione o cliente do débito.")
                Return False
            ElseIf objContaDebito.TemCentroDeCusto AndAlso String.IsNullOrWhiteSpace(ddlCustoDebito.SelectedValue) Then
                MsgBox(Me.Page, "Selecione o centro de custo do débito.")
                Return False
            ElseIf Not objContaDebito.TemCentroDeCusto AndAlso Not String.IsNullOrWhiteSpace(ddlCustoDebito.SelectedValue) Then
                MsgBox(Me.Page, "Esta conta de débito não tem centro de custo.")
                Return False
            ElseIf objContaDebito.TemProduto AndAlso String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
                MsgBox(Me.Page, "Selecione o produto para a conta de débito.")
                Return False
            End If
        End If

        If objContaCredito IsNot Nothing Then
            If objContaCredito.TemCliente AndAlso String.IsNullOrWhiteSpace(txtClienteCredito.Text.Length) Then
                MsgBox(Me.Page, "Selecione o cliente do crédito.")
                Return False
            ElseIf objContaCredito.TemCentroDeCusto AndAlso String.IsNullOrWhiteSpace(ddlCustoCredito.SelectedValue) Then
                MsgBox(Me.Page, "Selecione o centro de custo do crédito.")
                Return False
            ElseIf Not objContaCredito.TemCentroDeCusto AndAlso Not String.IsNullOrWhiteSpace(ddlCustoCredito.SelectedValue) Then
                MsgBox(Me.Page, "Esta conta de crédito não tem centro de custo.")
                Return False
            ElseIf objContaCredito.TemProduto AndAlso String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
                MsgBox(Me.Page, "Selecione o produto para a conta de crédito.")
                Return False
            End If
        End If

        If objContaDebito IsNot Nothing AndAlso objContaCredito IsNot Nothing AndAlso Not objContaDebito.TemProduto AndAlso Not objContaCredito.TemProduto AndAlso Not String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
            MsgBox(Me.Page, "Nenhuma das Contas permite produto.")
            Return False
        End If

        Return True
    End Function

    Private Function ValidaCentroDeCusto(ByVal conta As Integer) As Boolean
        Dim obj As New PlanoDeConta("", 0, conta)
        If obj IsNot Nothing AndAlso obj.TemCentroDeCusto Then
            Return True
        Else
            MsgBox(Me.Page, "Conta não permite centro de custo.")
            Return False
        End If
    End Function

    Function ValidarLote() As Boolean
        If ddlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Unidade de negocio é obrigatório.")
            Return False
        ElseIf ddlEmpresa.SelectedValue = "" Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        ElseIf txtMovimento.Text = "" Then
            MsgBox(Me.Page, "Data de movimento é obrigatório.")
            Return False
        ElseIf ddlLote.SelectedValue = "" Then
            MsgBox(Me.Page, "Lote é obrigatório.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub TotalizaLote()
        If ValidarLote() Then
            Dim sql As String = String.Empty
            Dim Empresa() = ddlEmpresa.SelectedValue.Split("-")

            sql = " SELECT Sum(DebitoOficial) As Debitos, Sum(CreditoOficial) as Creditos, Sum(DebitoOficial - CreditoOficial) as Diferenca" & vbCrLf & _
                  "   FROM Razao" & vbCrLf & _
                  "  WHERE Empresa_Id    = '" & Empresa(0) & "'" & vbCrLf & _
                  "    AND EndEmpresa_Id =  " & Empresa(1) & vbCrLf & _
                  "    And Movimento_Id  = '" & txtMovimento.Text.ToSqlDate() & "'" & vbCrLf & _
                  "    and Lote_Id       =  " & ddlLote.SelectedValue & vbCrLf

            For Each row As DataRow In Banco.ConsultaDataSet(sql, "Razao").Tables(0).Rows
                If Not IsDBNull(row("Debitos")) Then
                    lblTotalDebitos.Text = CDbl((row("Debitos"))).ToString("N2")
                End If
                If Not IsDBNull(row("Creditos")) Then
                    lblTotalCreditos.Text = CDbl((row("Creditos"))).ToString("N2")
                End If
                If Not IsDBNull(row("Diferenca")) Then
                    lblDiferenca.Text = CDbl((row("Diferenca"))).ToString("N2")
                    If CDbl((row("Diferenca"))) < 0 Then
                        lblDiferenca.ForeColor = Drawing.Color.Red
                    Else
                        lblDiferenca.ForeColor = Drawing.Color.Blue
                    End If
                End If
            Next

            ddlSequencia.Items.Clear()

            sql = " SELECT Sequencia_Id" & vbCrLf & _
                  "   FROM Razao" & vbCrLf & _
                  "  WHERE Empresa_Id    = '" & Empresa(0) & "'" & vbCrLf & _
                  "    AND EndEmpresa_Id =  " & Empresa(1) & vbCrLf & _
                  "    And Movimento_Id  = '" & txtMovimento.Text.ToSqlDate() & "'" & vbCrLf & _
                  "    and Lote_Id       =  " & ddlLote.SelectedValue & vbCrLf & _
                  "  Order by Sequencia_Id" & vbCrLf

            For Each row As DataRow In Banco.ConsultaDataSet(sql, "Clientes").Tables(0).Rows
                ddlSequencia.Items.Add(New ListItem(row("Sequencia_Id"), row("Sequencia_Id")))
            Next

            ddlSequencia.Items.Insert(0, "")
            ddlSequencia.SelectedIndex = 0

            ddlSequencia.Focus()

            If Not Funcoes.VerificaAcesso(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), txtMovimento.Text, "CONTABIL") Then
                MsgBox(Me.Page, "Movimento já fechado para esta data.")
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
                lnkNovo.Parent.Visible = False
            End If
        End If
    End Sub

    Private Function Sequencia(ByVal Empresa As String, ByVal EndEmpresa As Integer)
        Dim Sqll As String
        Dim Seq As Integer = 0

        Sqll = "SELECT isnull(Max(Sequencia_Id),0) + 1 as Sequencia" & vbCrLf & _
               " FROM Razao" & vbCrLf & _
               " WHERE Empresa_Id = '" & Empresa & "' AND EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
               " And Movimento_Id = '" & txtMovimento.Text.ToSqlDate() & "'" & vbCrLf & _
               " and Lote_Id = " & ddlLote.SelectedValue & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sqll, "Razao").Tables(0).Rows
            Seq = Dr("Sequencia")
        Next

        Return Seq
    End Function

    Private Sub Excluir()
        If ValidarLote() Then
            If Not String.IsNullOrWhiteSpace(ddlSequencia.Text) Then
                Dim sql As String
                sql = " DELETE RAZAO " & vbCrLf & _
                      "  WHERE Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                      "    AND EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
                      "    AND Movimento_Id = '" & txtMovimento.Text.ToSqlDate() & "'" & vbCrLf & _
                      "    AND Lote_Id = " & ddlLote.SelectedValue & vbCrLf & _
                      "    AND Sequencia_Id = " & ddlSequencia.SelectedValue & vbCrLf
                Banco.GravaBanco(sql)
            Else
                MsgBox(Me.Page, "Informe a sequencia do lote.")
            End If
        End If
    End Sub

    Private Sub Gravar()
        Dim sql As String
        Dim sqlarray As New ArrayList
        Dim SequenciaLote As Integer

        If txtContaDebito.Text.Length > 0 Then
            sql = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id,                " & vbCrLf & _
                  "                   Sequencia_Id, UnidadeDeNegocio, Titulo, Pedido, PedidoFixacao, Produto, Custo, Indexador,             " & vbCrLf & _
                  "                   DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, PrevistoRealizado,    " & vbCrLf & _
                  "                   Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, ChequeEntregue, PagamentoAutorizado, " & vbCrLf & _
                  "                   Processo, DebitoQuantidade, CreditoQuantidade, UsuarioInclusao, UsuarioInclusaoData)                  " & vbCrLf & _
                  "           VALUES (                                                                                                      " & vbCrLf
            sql &= "'" & ddlEmpresa.SelectedValue.Split("-")(0) & "', " & ddlEmpresa.SelectedValue.Split("-")(1) & ", '" & txtCodigoContaDebito.Value & "', " & vbCrLf

            If String.IsNullOrWhiteSpace(txtCodigoClienteDebito.Value) Then
                sql &= " '' , 0, "
            Else
                sql &= " '" & txtCodigoClienteDebito.Value.Split("-")(0) & "', " & txtCodigoClienteDebito.Value.Split("-")(1) & ", "
            End If

            sql &= "'" & txtMovimento.Text.ToSqlDate() & "', " & ddlLote.SelectedValue & ", "                                                  'Lote

            If ddlSequencia.SelectedValue <> "" Then
                SequenciaLote = ddlSequencia.SelectedValue
            Else
                SequenciaLote = Sequencia(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            End If

            sql &= SequenciaLote & ","                                                          'Sequencia

            sql &= "'" & ddlUnidade.SelectedValue & "', "                                        'Unidade de negocio Debito
            sql &= "0, "                                                                         'Titulo
            If txtClienteDebito.Text.Length > 0 AndAlso txtPedido.Text.Length > 0 AndAlso IsNumeric(txtPedido.Text) AndAlso CInt(txtPedido.Text) > 0 Then
                sql &= txtPedido.Text & ", "                                                     'Pedido
            Else
                sql &= "NULL, "                                                                     'Pedido
            End If

            sql &= "0, "                                                                         'Fixacao

            sql &= IIf(String.IsNullOrWhiteSpace(txtCodigoContaDebito.Value), "''", "'" & ddlProduto.SelectedValue & "'") & ", "

            sql &= IIf(String.IsNullOrWhiteSpace(ddlCustoDebito.SelectedValue), "0", ddlCustoDebito.SelectedValue) & ", "
            sql &= "0,"                                                                         'Indexador
            sql &= "'" & txtMovimento.Text.ToSqlDate() & "',"                                   'Data moeda
            sql &= Str(CDec(txtValor.Text)) & ","                     'Debito Oficial
            sql &= "0,0,0,"                                                                     'Credito Oficial, DebitoMoeda, CreditoMoeda

            If ddlHistorico.SelectedValue <> "" Then
                sql &= "'" & (RTrim(ddlHistorico.SelectedItem.Text) & " " & RTrim(txtComplemento.Text)).ToUpper & "',"  'Historico
            Else
                sql &= "'" & RTrim(txtComplemento.Text).ToUpper & "',"  'Historico
            End If

            sql &= "'P', '', 0, '', '', 0, '', '', 'CONTABIL', "

            Dim objContaDebito As New PlanoDeConta("", 0, txtCodigoContaDebito.Value)

            If Not String.IsNullOrWhiteSpace(txtQuantidade.Text) AndAlso IsNumeric(txtQuantidade.Text) Then
                sql &= Str(CDec(txtQuantidade.Text)) & ","
            Else
                sql &= "0, "         'DebitoQuantidade
            End If

            sql &= "0,"                                               'CreditoQuantidade
            sql &= "'" & UsuarioServidor.NomeUsuario & "',"           'Usuario Que Esta Incluindo
            sql &= "'" & Today.ToSqlDate() & "')"                     'Data da desta Inclusao

            sqlarray.Add(sql)

        End If

        If txtContaCredito.Text.Length > 0 Then
            sql = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id,                " & vbCrLf & _
                 "                   Sequencia_Id, UnidadeDeNegocio, Titulo, Pedido, PedidoFixacao, Produto, Custo, Indexador,             " & vbCrLf & _
                 "                   DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, PrevistoRealizado,    " & vbCrLf & _
                 "                   Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, ChequeEntregue, PagamentoAutorizado, " & vbCrLf & _
                 "                   Processo, DebitoQuantidade, CreditoQuantidade, UsuarioInclusao, UsuarioInclusaoData)                  " & vbCrLf & _
                 "           VALUES (                                                                                                      " & vbCrLf

            sql &= "'" & ddlEmpresa.SelectedValue.Split("-")(0) & "', "                                                      'CGC EmpresaDebito
            sql &= ddlEmpresa.SelectedValue.Split("-")(1) & ", "                                                             'Endereco EmpresaDebito
            sql &= "'" & txtCodigoContaCredito.Value & "',"  'ContaDebito

            If String.IsNullOrWhiteSpace(txtCodigoClienteCredito.Value) Then
                sql &= " '' , 0, "
            Else
                sql &= " '" & txtCodigoClienteCredito.Value.Split("-")(0) & "', " & txtCodigoClienteCredito.Value.Split("-")(1) & ", "
            End If

            sql &= "'" & txtMovimento.Text.ToSqlDate() & "', "                                                            'Data de movimento
            sql &= ddlLote.SelectedValue & ", "                                                  'Lote

            If txtContaDebito.Text.Length > 0 Then
                SequenciaLote += 1
            Else
                If ddlSequencia.SelectedValue <> "" Then
                    SequenciaLote = ddlSequencia.SelectedValue
                Else
                    SequenciaLote = Sequencia(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
                End If
            End If

            sql &= SequenciaLote & ","                                                          'Sequencia

            sql &= "'" & ddlUnidade.SelectedValue & "',"                                        'Unidade de negocio Debito
            sql &= "0,"                                                                         'Titulo

            If txtClienteCredito.Text.Length > 0 AndAlso txtPedido.Text.Length > 0 AndAlso IsNumeric(txtPedido.Text) AndAlso CInt(txtPedido.Text) > 0 Then
                sql &= txtPedido.Text & ","                                                     'Pedido
            Else
                sql &= "NULL,"                                                                     'Pedido
            End If
            sql &= "0,"                                                                         'Fixacao


            sql &= "'" & ddlProduto.SelectedValue & "',"                                    'Produto Debito
            sql &= IIf(String.IsNullOrWhiteSpace(ddlCustoCredito.SelectedValue), "0", ddlCustoCredito.SelectedValue) & ", "                                 'Codigo de custo debito

            sql &= "0,"                                                                         'Indexador
            sql &= "'" & txtMovimento.Text.ToSqlDate() & "',"                                                            'Data moeda
            sql &= "0,"                                                                         'Indexador
            sql &= Str(CDec(txtValor.Text)) & ", "                     'Credito Oficial
            sql &= "0,0,"                                                                       'Credito Oficial, DebitoMoeda, CreditoMoeda

            If ddlHistorico.SelectedValue <> "" Then
                sql &= "'" & (RTrim(ddlHistorico.SelectedItem.Text) & " " & RTrim(txtComplemento.Text)).ToUpper & "',"  'Historico
            Else
                sql &= "'" & RTrim(txtComplemento.Text).ToUpper & "',"  'Historico
            End If

            sql &= "'P', '', 0, '', '', 0, '', '', 'CONTABIL', "
            sql &= "0, "                                                                  'DebitoQuantidade


            If Not String.IsNullOrWhiteSpace(txtQuantidade.Text) AndAlso IsNumeric(txtQuantidade.Text) Then
                sql &= Str(CDec(txtQuantidade.Text)) & ","
            Else
                sql &= "0, "         'DebitoQuantidade
            End If

            sql &= "'" & UsuarioServidor.NomeUsuario & "', "           'Usuario Que Esta Incluindo
            sql &= "'" & Today.ToSqlDate() & "')"                     'Data da desta Inclusao

            sqlarray.Add(sql)
        End If

        If Banco.GravaBanco(sqlarray) Then
            If ddlSequencia.SelectedIndex > -1 Then
                ddlSequencia.SelectedIndex = 0
            End If

            If Not ChkAproveitarDados.Checked Then
                Limpar(False)
            End If
            TotalizaLote()

            MsgBox(Me.Page, "Registro gravado com Sucesso.", eTitulo.Sucess)
        End If
    End Sub

    Private Sub Consultar()
        If ValidarLote() Then
            If String.IsNullOrWhiteSpace(ddlSequencia.Text) Then
                MsgBox(Me.Page, "Sequência do lote é obrigatório.")
            Else
                'Limpar os campos exceto os necessários para consulta
                Limpar(False)

                Dim sql As String
                Dim empresa() As String = ddlEmpresa.SelectedValue.Split("-")
                Dim Codigo As String
                Dim Endereco As Integer

                sql = "SELECT R.Empresa_Id, R.EndEmpresa_Id, R.Conta_Id, R.Cliente_Id, R.EndCliente_Id, Cli.Nome,  R.Movimento_Id, R.Lote_Id, R.Sequencia_Id, R.UnidadeDeNegocio, " & vbCrLf &
                      "       PC.Titulo, isnull(R.Pedido,0) AS Pedido, isnull(R.PedidoFixacao,0) as PedidoFixacao, isnull(R.Produto,'') as Produto, P.Grupo AS GrupoProduto, R.Custo, R.Indexador, " & vbCrLf &
                      "       R.DataMoeda, R.DebitoOficial, R.CreditoOficial, R.DebitoMoeda, R.CreditoMoeda,R. Historico, R.PrevistoRealizado, R.Cliente_Nf, R.EndCliente_Nf, " & vbCrLf &
                      "       R.EntradaSaida_Nf, R.Serie_Nf, R.Numero_Nf, R.ChequeEntregue, R.PagamentoAutorizado, R.DataDaBaixa, R.Conciliacao, R.Deposito, R.EndDeposito, R.Rateado, R.Processo, " & vbCrLf &
                      "       isnull(R.DebitoQuantidade,0) AS DebitoQuantidade, isnull(R.CreditoQuantidade,0) AS CreditoQuantidade, R.Situacao, R.Produto_NF, R.Cfop_NF, R.Sequencia_NF, R.Encargo_NF, " & vbCrLf &
                      "       isnull(R.UsuarioInclusao,'') As UsuarioInclusao" & vbCrLf &
                      "  FROM Razao R" & vbCrLf &
                      "  LEFT JOIN Clientes Cli" & vbCrLf &
                      "    ON R.Cliente_Id = Cli.Cliente_Id" & vbCrLf &
                      "   AND R.EndCliente_Id = Cli.Endereco_Id" & vbCrLf &
                      "  LEFT JOIN Produtos P" & vbCrLf &
                      "    ON R.Produto = P.Produto_Id " & vbCrLf &
                      "  JOIN PlanoDeContas PC " & vbCrLf &
                      "    ON PC.Conta_ID = R.Conta_ID" & vbCrLf &
                      " WHERE R.Empresa_Id    = '" & empresa(0) & "'" & vbCrLf &
                      "   AND R.EndEmpresa_Id =  " & empresa(1) & vbCrLf &
                      "   And R.Movimento_Id  = '" & txtMovimento.Text.ToSqlDate() & "' " & vbCrLf &
                      "   and R.Lote_Id       =  " & ddlLote.SelectedValue & vbCrLf &
                      "   and R.Sequencia_Id  =  " & ddlSequencia.SelectedValue

                For Each row As DataRow In Banco.ConsultaDataSet(sql, "Razao").Tables(0).Rows

                    lblUsuario.Text = row("UsuarioInclusao")

                    If row("DebitoOficial") <> 0 Then
                        'DEBITO
                        txtValor.Text = row("DebitoOficial")
                        txtCodigoContaDebito.Value = row("conta_id")
                        txtContaDebito.Text = row("Conta_Id") & " - " & row("Titulo")

                        If row("Cliente_Id") <> "" Then
                            Codigo = row("Cliente_Id")
                            Endereco = row("EndCliente_Id")
                            txtCodigoClienteDebito.Value = row("cliente_id") & "-" & row("EndCliente_Id")
                            txtClienteDebito.Text = Funcoes.FormatarCpfCnpj(row("Cliente_ID")) & " - " & UCase(row("Nome"))
                        End If

                        If row("Custo") <> 0 Then
                            ddlCustoDebito.SelectedValue = row("Custo")
                        End If

                        If row("Produto") <> "" Then
                            DdlGrupo.SelectedValue = row("GrupoProduto")
                            ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, "Grupo = '" & DdlGrupo.SelectedValue & "'", True)
                            ddlProduto.SelectedValue = row("Produto")
                        End If

                        btnContasCredito.Enabled = False
                        btnClientesCredito.Enabled = False
                        ddlCustoCredito.Enabled = False
                    Else
                        'CREDITO
                        txtValor.Text = row("CreditoOficial")
                        txtCodigoContaCredito.Value = row("conta_id")
                        txtContaCredito.Text = row("Conta_Id") & " - " & row("Titulo")

                        If row("Cliente_Id") <> "" Then
                            txtCodigoClienteCredito.Value = row("cliente_id") & "-" & row("EndCliente_Id")
                            txtClienteCredito.Text = Funcoes.FormatarCpfCnpj(row("Cliente_ID")) & " - " & UCase(row("Nome").ToString())
                        End If

                        If row("Custo") <> 0 Then
                            ddlCustoCredito.SelectedValue = row("Custo")
                        End If

                        btnContasDebito.Enabled = False
                        btnClientesDebito.Enabled = False
                        ddlCustoDebito.Enabled = False
                    End If

                    If row("Produto") <> "" Then
                        DdlGrupo.SelectedValue = row("GrupoProduto")
                        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, "Grupo = '" & DdlGrupo.SelectedValue & "'", True)
                        ddlProduto.SelectedValue = row("Produto")
                    End If

                    txtPedido.Text = row("Pedido")
                    txtComplemento.Text = row("Historico").ToString.ToUpper

                    If row("DebitoQuantidade") > 0 Then
                        txtQuantidade.Text = row("DebitoQuantidade")
                    ElseIf row("CreditoQuantidade") > 0 Then
                        txtQuantidade.Text = row("CreditoQuantidade")
                    End If
                Next


                If Not Funcoes.VerificaAcesso(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), txtMovimento.Text, "CONTABIL") Then
                    MsgBox(Me.Page, "Movimento já fechado para esta data.")
                    lnkAtualizar.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False
                Else
                    lnkAtualizar.Parent.Visible = True
                    lnkExcluir.Parent.Visible = True
                End If

                lnkNovo.Parent.Visible = False
            End If
        End If
    End Sub

#End Region

End Class