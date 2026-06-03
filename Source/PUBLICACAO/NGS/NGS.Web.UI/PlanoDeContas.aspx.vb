Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class PlanoDeContas
    Inherits BasePage

    Private objConta As New [Lib].Negocio.PlanoDeConta
    Private objNovoEncargo As [Lib].Negocio.EncargosPlanoDeContas

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Contabil.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("PlanoDeContas", "ACESSAR") Then
                ddl.Carregar(ddlTipoDaConta, CarregarDDL.Tabela.TipoDaContaContabil, "")
                ddl.Carregar(ddlTiposDeClientes, CarregarDDL.Tabela.TipoDeCliente, "")
                ddl.Carregar(ddlContasOrcamentarias, CarregarDDL.Tabela.ContasOrcamentarias, "")
                ddl.Carregar(ddlResponsabilidadeDeContas, CarregarDDL.Tabela.ResponsabilidadeDaConta, "")
                ddl.Carregar(ddlDacon, CarregarDDL.Tabela.Dacon, "")
                ddl.Carregar(ddlCustoECDBacen, CarregarDDL.Tabela.CCustoECD, "ATIVO = 1")
                carregarGrupoBacen()
                carregarCustoECFBacen()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Contabil.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Session("PL" & HID.Value) IsNot Nothing Then
            Dim objplano As New [Lib].Negocio.PlanoDeConta("", 0, Session("PL" & HID.Value))
            Session.Remove("PL" & HID.Value)
        ElseIf Session("objContaPLxENC" & HID.Value) IsNot Nothing Then
            RecuperarConta()
            Dim objEncargo As [Lib].Negocio.PlanoDeConta = Session("objContaPLxENC" & HID.Value)
            Session.Remove("objContaPLxENC" & HID.Value)

            If Not objEncargo.Encargo Then
                MsgBox(Me.Page, "Contas que não PARTICIPAM NO FINANCEIRO não podem ser usadas.")
                Exit Sub
            End If

            If objConta.EncargosPlanoDeContas.Count = 0 Then
                objNovoEncargo = New [Lib].Negocio.EncargosPlanoDeContas(objConta)
                objNovoEncargo.IUD = "I"
                objNovoEncargo.CodigoContaEncargo = objConta.Conta
                If Not objNovoEncargo.Salvar() Then
                    MsgBox(Me.Page, "Erro ao Salvar o Encargo Principal: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    Exit Sub
                End If
                objConta.IUD = "U"
                objConta.TemEncargo = True
                If Not objConta.Salvar() Then
                    MsgBox(Me.Page, "Erro ao Salvar a Conta Principal: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    Exit Sub
                End If
            End If

            objNovoEncargo = New [Lib].Negocio.EncargosPlanoDeContas(objConta)
            objNovoEncargo.IUD = "I"
            objNovoEncargo.CodigoContaEncargo = objEncargo.Conta

            If Not objNovoEncargo.Salvar() Then
                MsgBox(Me.Page, "Erro ao salvar o encargo: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                Exit Sub
            End If

            objConta.EncargosPlanoDeContas = New [Lib].Negocio.ListEncargosPlanoDeContas(objConta)
            gridEncargos.DataSource = objConta.EncargosPlanoDeContas.ToArray()
            gridEncargos.DataBind()
            SalvarConta()
        End If
    End Sub

    Private Sub SalvarConta()
        Session("objConta" & HID.Value) = objConta
    End Sub

    Private Sub RecuperarConta()
        objConta = CType(Session("objConta" & HID.Value), [Lib].Negocio.PlanoDeConta)
    End Sub

    Private Sub CargaPlanoDeContas(ByVal pConta As String)
        If Funcoes.VerificaPermissao("PlanoDeContas", "LEITURA") Then
            Dim ListaPlano As New [Lib].Negocio.ListPlanoDeConta(pConta)

            gridPlanoDeContas.DataSource = ListaPlano.ToArray()
            gridPlanoDeContas.DataBind()

            Dim i As Integer = 0
            While i < gridPlanoDeContas.Rows.Count
                If CType(gridPlanoDeContas.Rows(i).FindControl("lblTemCliente"), Label).Text = "True" Then
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemCliente"), Label).Text = "SIM"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemCliente"), Label).ForeColor = Drawing.Color.Red
                Else
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemCliente"), Label).Text = "NÃO"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemCliente"), Label).ForeColor = Drawing.Color.Black
                End If

                If CType(gridPlanoDeContas.Rows(i).FindControl("lblTemCCusto"), Label).Text = "True" Then
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemCCusto"), Label).Text = "SIM"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemCCusto"), Label).ForeColor = Drawing.Color.Red
                Else
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemCCusto"), Label).Text = "NÃO"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemCCusto"), Label).ForeColor = Drawing.Color.Black
                End If

                If CType(gridPlanoDeContas.Rows(i).FindControl("lblTemProduto"), Label).Text = "True" Then
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemProduto"), Label).Text = "SIM"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemProduto"), Label).ForeColor = Drawing.Color.Red
                Else
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemProduto"), Label).Text = "NÃO"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemProduto"), Label).ForeColor = Drawing.Color.Black
                End If

                If CType(gridPlanoDeContas.Rows(i).FindControl("lblTemPedido"), Label).Text = "True" Then
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemPedido"), Label).Text = "SIM"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemPedido"), Label).ForeColor = Drawing.Color.Red
                Else
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemPedido"), Label).Text = "NÃO"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemPedido"), Label).ForeColor = Drawing.Color.Black
                End If

                If CType(gridPlanoDeContas.Rows(i).FindControl("lblTemAdto"), Label).Text = "True" Then
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemAdto"), Label).Text = "SIM"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemAdto"), Label).ForeColor = Drawing.Color.Red
                Else
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemAdto"), Label).Text = "NÃO"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemAdto"), Label).ForeColor = Drawing.Color.Black
                End If

                If CType(gridPlanoDeContas.Rows(i).FindControl("lblAdSoContabil"), Label).Text = "True" Then
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblAdSoContabil"), Label).Text = "SIM"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblAdSoContabil"), Label).ForeColor = Drawing.Color.Red
                Else
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblAdSoContabil"), Label).Text = "NÃO"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblAdSoContabil"), Label).ForeColor = Drawing.Color.Black
                End If


                If CType(gridPlanoDeContas.Rows(i).FindControl("lblTemEncargo"), Label).Text = "True" Then
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemEncargo"), Label).Text = "SIM"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemEncargo"), Label).ForeColor = Drawing.Color.Red
                Else
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemEncargo"), Label).Text = "NÃO"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblTemEncargo"), Label).ForeColor = Drawing.Color.Black
                End If

                If CType(gridPlanoDeContas.Rows(i).FindControl("lblAtivo"), Label).Text = "True" Then
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblAtivo"), Label).Text = "SIM"

                Else
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblAtivo"), Label).Text = "NÃO"
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblAtivo"), Label).ForeColor = Drawing.Color.Black
                    CType(gridPlanoDeContas.Rows(i).FindControl("lblAtivo"), Label).ForeColor = Drawing.Color.Red
                End If

                i += 1
            End While
        End If
    End Sub

    Protected Sub gridPlanoDeContas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridPlanoDeContas.SelectedIndexChanged
        Try
            objConta = New [Lib].Negocio.PlanoDeConta("", 0, gridPlanoDeContas.SelectedRow.Cells(1).Text())

            objConta.IUD = "U"
            txtConta.Text = objConta.Conta
            txtTitulo.Text = objConta.Titulo
            ddlTipoDaConta.SelectedIndex = ddlTipoDaConta.Items.IndexOf(ddlTipoDaConta.Items.FindByValue(objConta.TipodeConta))
            ddlTiposDeClientes.SelectedIndex = ddlTiposDeClientes.Items.IndexOf(ddlTiposDeClientes.Items.FindByValue(objConta.CodigoTipoDeCliente))
            ddlContasOrcamentarias.SelectedIndex = ddlContasOrcamentarias.Items.IndexOf(ddlContasOrcamentarias.Items.FindByValue(objConta.ContaOrcamentaria))
            ddlResponsabilidadeDeContas.SelectedIndex = ddlResponsabilidadeDeContas.Items.IndexOf(ddlResponsabilidadeDeContas.Items.FindByValue(objConta.Responsabilidade))
            ddlDacon.SelectedIndex = ddlDacon.Items.IndexOf(ddlDacon.Items.FindByValue(objConta.Dacon))

            If objConta.TemCliente Then
                chkCliente.Checked = True
                ddlTiposDeClientes.Enabled = True
            Else
                chkCliente.Checked = False
                ddlTiposDeClientes.SelectedIndex = 0
                ddlTiposDeClientes.Enabled = False
            End If

            If objConta.TemCentroDeCusto Then
                chkCentroDeCusto.Checked = True
            Else
                chkCentroDeCusto.Checked = False
            End If

            If objConta.TemProduto Then
                chkProduto.Checked = True
            Else
                chkProduto.Checked = False
            End If

            If objConta.TemPedido Then
                chkPedido.Checked = True
            Else
                chkPedido.Checked = False
            End If

            If objConta.ProdutoParaCusto Then
                chkProdutoParaCusto.Checked = True
            Else
                chkProdutoParaCusto.Checked = False
            End If

            If objConta.Adiantamento Then
                chkAdiantamento.Checked = True
                chkAdSoContabil.Enabled = True
            Else
                chkAdiantamento.Checked = False
                chkAdSoContabil.Enabled = False
            End If

            If objConta.AdiantamentoSoContabil Then
                chkAdSoContabil.Checked = True
            Else
                chkAdSoContabil.Checked = False
            End If

            If objConta.Encargo Then
                chkEncargo.Checked = True
                ddlContasAReceber.Enabled = True
                ddlContasAPagar.Enabled = True
                ddlContasAReceber.SelectedValue = objConta.Receber
                ddlContasAPagar.SelectedValue = objConta.Pagar
            Else
                chkEncargo.Checked = False
                ddlContasAReceber.SelectedIndex = 0
                ddlContasAPagar.SelectedIndex = 0
                ddlContasAReceber.Enabled = False
                ddlContasAPagar.Enabled = False
            End If

            If objConta.Ativo Then
                ddlAtivo.SelectedIndex = 0
            Else
                ddlAtivo.SelectedIndex = 1
            End If

            ddlTipoDeCusto.SelectedValue = objConta.TipoDeCusto

            gridEncargos.DataSource = objConta.EncargosPlanoDeContas.ToArray()
            gridEncargos.DataBind()

            txtConta.Enabled = False
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            pnlEncargos.Visible = True

            lnkNovoBacen.Parent.Visible = True
            lnkAtualizarBacen.Parent.Visible = False
            lnkExcluirBacen.Parent.Visible = False

            SalvarConta()

            ddl.Carregar(ddlRefBacen, CarregarDDL.Tabela.ReferencialBacen, "LEFT(CONTA_ID, 1) = " & IIf(Left(objConta.Conta, 1) = 4, 3, Left(objConta.Conta, 1)))

            txtContaBacen.Text = objConta.Conta
            txtTituloBacen.Text = objConta.Titulo

            preencherGridContaBacen(objConta.Conta)

            TabContainer1.ActiveTabIndex = 0

            ValidarConta()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ValidarConta()
        Try
            Dim sql As String = " select top 1 Conta_Id from razao " & vbCrLf &
                                " where Conta_Id = '" & txtConta.Text & "'"
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Razao")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False

                MsgBox(Me.Page, "Conta não pode ser alterado/excluída pois foi contabilizada no Razão.")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Function ValidaCamposPlanoDeContas() As Boolean
        If Trim(txtConta.Text) = "" Then
            MsgBox(Me.Page, "Conta é Obrigatoria.")
            Return False
        ElseIf Trim(txtTitulo.Text) = "" Then
            MsgBox(Me.Page, "Descrição da Conta Obrigatoria.")
            Return False
        ElseIf chkEncargo.Checked AndAlso (ddlContasAReceber.SelectedIndex = 0 Or ddlContasAPagar.SelectedIndex = 0) Then
            If ddlContasAReceber.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Tipo do Encargo no Receber não foi selecionado.")
            Else
                MsgBox(Me.Page, "Tipo do Encargo no Pagar não foi selecionado.")
            End If
            Return False
        ElseIf chkCliente.Checked AndAlso ddlTiposDeClientes.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Tipo do Cliente não foi selecionado.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub Limpar()
        Session.Remove("PL" & HID.Value)
        Session.Remove("PLxENC" & HID.Value)
        Session.Remove("objConta" & HID.Value)
        HID.Value = Guid.NewGuid().ToString
        txtConta.Text = String.Empty
        txtTitulo.Text = String.Empty
        ddlTipoDaConta.SelectedIndex = 0
        ddlTiposDeClientes.SelectedIndex = 0
        ddlContasOrcamentarias.SelectedIndex = 0
        ddlResponsabilidadeDeContas.SelectedIndex = 0
        ddlDacon.SelectedIndex = 0
        ddlTipoDeCusto.SelectedIndex = 0

        ddlCustoECDBacen.SelectedIndex = 0

        ddlContasAReceber.SelectedIndex = 0
        ddlContasAPagar.SelectedIndex = 0
        ddlAtivo.SelectedIndex = 0
        ddlContasAReceber.Enabled = False
        ddlContasAPagar.Enabled = False

        chkCliente.Checked = False
        chkCentroDeCusto.Checked = False
        chkProduto.Checked = False
        chkPedido.Checked = False
        chkAdiantamento.Checked = False
        chkAdSoContabil.Enabled = False
        chkAdSoContabil.Checked = False
        chkEncargo.Checked = False
        chkProdutoParaCusto.Checked = False

        gridEncargos.DataSource = Nothing
        gridEncargos.DataBind()

        txtConta.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        pnlEncargos.Visible = False

        objConta = New [Lib].Negocio.PlanoDeConta()
        objConta.IUD = "I"

        SalvarConta()
        ucConsultaPlanoDeContas.SetarHID(HID.Value)

        LimparCamposBacen(True)
    End Sub

    Protected Sub cmdAjuda_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim NomeArquivo As String = "Manual/PlanoDeContas.mht"
        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
    End Sub

    Protected Sub imgExcluirEncargo_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim imgEncargo As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(imgEncargo.NamingContainer, GridViewRow)

        RecuperarConta()
        If objConta.EncargosPlanoDeContas(row.RowIndex).CodigoContaEncargo = objConta.Conta And objConta.EncargosPlanoDeContas.Count > 1 Then
            MsgBox(Me.Page, "Encargo principal não pode ser Excluido.")
            Exit Sub
        End If

        objConta.EncargosPlanoDeContas(row.RowIndex).IUD = "D"
        If Not objConta.EncargosPlanoDeContas(row.RowIndex).Salvar Then
            MsgBox(Me.Page, "Erro ao Excluir o Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
            Exit Sub
        End If

        'If objConta.EncargosPlanoDeContas.Count = 2 Then
        '    objNovoEncargo = New Negocio.EncargosPlanoDeContas(objConta)
        '    objNovoEncargo.IUD = "D"
        '    objNovoEncargo.CodigoContaEncargo = objConta.Conta
        '    If Not objNovoEncargo.Salvar() Then
        '        MsgBox(Me.Page, "Erro ao Excluir o Encargo Principal: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
        '        Exit Sub
        '    End If

        '    objConta.IUD = "U"
        '    objConta.TemEncargo = False
        '    If Not objConta.Salvar() Then
        '        MsgBox(Me.Page, "Erro ao Salvar a Conta Principal: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
        '        Exit Sub
        '    End If
        'End If

        objConta.EncargosPlanoDeContas = New [Lib].Negocio.ListEncargosPlanoDeContas(objConta)
        gridEncargos.DataSource = objConta.EncargosPlanoDeContas.ToArray()
        gridEncargos.DataBind()
        SalvarConta()
    End Sub

    Protected Sub chkEncargo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If chkEncargo.Checked Then
            ddlContasAReceber.Enabled = True
            ddlContasAPagar.Enabled = True
        Else
            RecuperarConta()
            If objConta.EncargosPlanoDeContas.Count > 0 Then
                chkEncargo.Checked = True
                MsgBox(Me.Page, "A Participacao no Financeiro só pode ser desmarcada se nao houver nenhuma conta relacionada.")
                Exit Sub
            End If

            ddlContasAReceber.SelectedIndex = 0
            ddlContasAPagar.SelectedIndex = 0
            ddlContasAReceber.Enabled = False
            ddlContasAPagar.Enabled = False
        End If
    End Sub

    Protected Sub chkCliente_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If chkCliente.Checked Then
            ddlTiposDeClientes.Enabled = True
        Else
            ddlTiposDeClientes.SelectedIndex = 0
            ddlTiposDeClientes.Enabled = False
        End If
    End Sub

    Protected Sub chkAdiantamento_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If CType(sender, CheckBox).Checked Then
            chkAdSoContabil.Enabled = True
        Else
            chkAdSoContabil.Enabled = False
            chkAdSoContabil.Checked = False
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("PlanoDeContas", "GRAVAR") Then
                If txtConta.Text.Length = 0 Then
                    MsgBox(Me.Page, "Conta não foi informada")
                    Exit Sub
                ElseIf chkEncargo.Checked AndAlso (ddlContasAReceber.SelectedIndex = 0 AndAlso ddlContasAPagar.SelectedIndex = 0) Then
                    MsgBox(Me.Page, "Selecione como esta conta se comportará no contas a RECEBER/PAGAR")
                    Exit Sub
                End If

                RecuperarConta()

                objConta.CodigoEmpresa = "99999999999999"
                objConta.EnderecoEmpresa = 0
                objConta.IUD = "I"
                objConta.Conta = txtConta.Text
                objConta.Titulo = txtTitulo.Text

                If ddlTipoDaConta.SelectedIndex > 0 Then
                    objConta.TipodeConta = ddlTipoDaConta.SelectedValue
                Else
                    objConta.TipodeConta = ""
                End If

                If ddlTiposDeClientes.SelectedIndex > 0 Then
                    objConta.CodigoTipoDeCliente = ddlTiposDeClientes.SelectedValue
                Else
                    objConta.CodigoTipoDeCliente = 0
                End If

                If ddlContasOrcamentarias.SelectedIndex > 0 Then
                    objConta.ContaOrcamentaria = ddlContasOrcamentarias.SelectedValue
                Else
                    objConta.ContaOrcamentaria = ""
                End If

                If ddlResponsabilidadeDeContas.SelectedIndex > 0 Then
                    objConta.Responsabilidade = ddlResponsabilidadeDeContas.SelectedValue
                Else
                    objConta.Responsabilidade = ""
                End If

                If ddlDacon.SelectedIndex > 0 Then
                    objConta.Dacon = ddlDacon.SelectedValue
                Else
                    objConta.Dacon = 0
                End If

                objConta.TemCliente = chkCliente.Checked
                objConta.TemCentroDeCusto = chkCentroDeCusto.Checked
                objConta.TemProduto = chkProduto.Checked
                objConta.TemPedido = chkPedido.Checked
                objConta.ProdutoParaCusto = chkProdutoParaCusto.Checked
                objConta.Adiantamento = chkAdiantamento.Checked
                objConta.AdiantamentoSoContabil = chkAdSoContabil.Checked

                If chkEncargo.Checked Then
                    objConta.Encargo = True

                    If ddlContasAReceber.SelectedIndex > 0 Then
                        objConta.Receber = ddlContasAReceber.SelectedValue
                    Else
                        objConta.Receber = ""
                    End If

                    If ddlContasAPagar.SelectedIndex > 0 Then
                        objConta.Pagar = ddlContasAPagar.SelectedValue
                    Else
                        objConta.Pagar = ""
                    End If
                Else
                    objConta.Encargo = False
                    objConta.Receber = ""
                    objConta.Pagar = ""
                End If

                objConta.TipoDeCusto = ddlTipoDeCusto.SelectedValue
                objConta.Ativo = ddlAtivo.SelectedValue

                objConta.UsuarioInclusao = UsuarioServidor.Usuario.Usuario_Id
                objConta.UsuarioInclusaoData = DateTime.Now()

                If objConta.Salvar Then
                    CargaPlanoDeContas("Conta_Id LIKE '" & txtConta.Text & "%'")
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

    Protected Sub lnkpdf_Click(sender As Object, e As EventArgs) Handles lnkPdf.Click
        Try
            Emitir(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkWord_Click(sender As Object, e As EventArgs) Handles lnkWord.Click
        Try
            Emitir(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            Emitir(False, False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("PlanoDeContas", "ALTERAR") Then
                If chkEncargo.Checked AndAlso (ddlContasAReceber.SelectedIndex = 0 And ddlContasAPagar.SelectedIndex = 0) Then
                    MsgBox(Me.Page, "Selecione como esta conta se comportará no contas a RECEBER/PAGAR")
                    Exit Sub
                End If

                RecuperarConta()

                objConta.IUD = "U"
                objConta.Titulo = txtTitulo.Text

                If ddlTipoDaConta.SelectedIndex > 0 Then
                    objConta.TipodeConta = objConta.TipodeConta = ddlTipoDaConta.SelectedValue
                Else
                    objConta.TipodeConta = ""
                End If

                If ddlTiposDeClientes.SelectedIndex > 0 Then
                    objConta.CodigoTipoDeCliente = ddlTiposDeClientes.SelectedValue
                Else
                    objConta.CodigoTipoDeCliente = 0
                End If

                If ddlContasOrcamentarias.SelectedIndex > 0 Then
                    objConta.ContaOrcamentaria = ddlContasOrcamentarias.SelectedValue
                Else
                    objConta.ContaOrcamentaria = ""
                End If

                If ddlResponsabilidadeDeContas.SelectedIndex > 0 Then
                    objConta.Responsabilidade = ddlResponsabilidadeDeContas.SelectedValue
                Else
                    objConta.Responsabilidade = 0
                End If

                If ddlDacon.SelectedIndex > 0 Then
                    objConta.Dacon = ddlDacon.SelectedValue
                Else
                    objConta.Dacon = 0
                End If

                objConta.TemCliente = chkCliente.Checked
                objConta.TemCentroDeCusto = chkCentroDeCusto.Checked
                objConta.TemProduto = chkProduto.Checked
                objConta.TemPedido = chkPedido.Checked
                objConta.ProdutoParaCusto = chkProdutoParaCusto.Checked
                objConta.Adiantamento = chkAdiantamento.Checked
                objConta.AdiantamentoSoContabil = chkAdSoContabil.Checked

                If chkEncargo.Checked Then
                    objConta.Encargo = True

                    If ddlContasAReceber.SelectedIndex > 0 Then
                        objConta.Receber = ddlContasAReceber.SelectedValue
                    Else
                        objConta.Receber = ""
                    End If

                    If ddlContasAPagar.SelectedIndex > 0 Then
                        objConta.Pagar = ddlContasAPagar.SelectedValue
                    Else
                        objConta.Pagar = ""
                    End If
                Else
                    objConta.Encargo = False
                    objConta.Receber = ""
                    objConta.Pagar = ""
                End If

                objConta.TipoDeCusto = ddlTipoDeCusto.SelectedValue
                objConta.Ativo = ddlAtivo.SelectedValue

                objConta.UsuarioAlteracao = UsuarioServidor.Usuario.Usuario_Id
                objConta.UsuarioAlteracaoData = DateTime.Now()

                If objConta.Salvar Then
                    CargaPlanoDeContas("Conta_Id LIKE '" & txtConta.Text & "%'")
                    Limpar()
                Else
                    MsgBox(Me.Page, "Erro ao Alterar o Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
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
            If Funcoes.VerificaPermissao("PlanoDeContas", "EXCLUIR") Then
                RecuperarConta()
                If objConta.VerificaRazao Then
                    MsgBox(Me.Page, "Conta com Lançamentos no Razão não pode ser Excluída")
                ElseIf objConta.VerificaOperacao Then
                    MsgBox(Me.Page, "Conta vinculada na SubOperações não pode ser Excluída")
                ElseIf objConta.VerificaOperacoesXEncargos Then
                    MsgBox(Me.Page, "Conta vinculada na OperacoesXEncargos não pode ser Excluída")
                Else
                    objConta.IUD = "D"
                    If objConta.Salvar Then
                        CargaPlanoDeContas("Conta_Id LIKE '" & txtConta.Text & "%'")
                        Limpar()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("PlanoDeContas", "LEITURA") Then
                If txtTitulo.Text.Length > 0 Then
                    CargaPlanoDeContas("Titulo LIKE '" & txtTitulo.Text & "%' And isnull(Ativo, 1) = " & ddlAtivo.SelectedValue)
                ElseIf txtConta.Text.Length > 0 Then
                    CargaPlanoDeContas("Conta_Id LIKE '" & txtConta.Text & "%' And isnull(Ativo, 1) = " & ddlAtivo.SelectedValue)
                Else
                    MsgBox(Me.Page, "Informe Parte da Conta ou Parte do Nome da Conta para Consulta.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro(s).")
            End If
            TabContainer1.ActiveTabIndex = 1
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
        CargaPlanoDeContas("")
    End Sub

    Private Function getDataSet() As DataSet
        Dim objBanco As New AcessaBanco()
        Dim sql As String

        sql = "SELECT Conta_Id as Conta, Titulo, isnull(Cliente,'N') AS TemCliente, isnull(Produto,'N') AS TemProduto, isnull(CentroDeCusto,'N') AS TemCentroDeCusto, " & vbCrLf & _
              "       ISNULL(TipoDeCliente, 0) AS CodigoTipoDeCliente, Responsabilidade, ContaOrcamentaria, isnull(Dacon,0) as Dacon" & vbCrLf & _
              "  FROM PlanoDeContas " & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtConta.Text) Then
            sql &= "Where Conta_Id LIKE '" & txtConta.Text & "%'" & vbCrLf
        End If

        sql &= " ORDER BY Conta_Id "

        Return objBanco.ConsultaDataSet(sql, "PlanoDeContas")
    End Function

    Private Sub Emitir(ByVal pdf As Boolean, Optional ByVal excel As Boolean = True)
        If Funcoes.VerificaPermissao("PlanoDeContas", "RELATORIO") Then
            Dim ds As DataSet = getDataSet()
            Dim param As New Dictionary(Of String, Object)
            param.Add("Detalhado", chkDetalhado.Checked)

            If pdf Then
                Funcoes.BindReport(Me.Page, ds, "Cr_PlanoDeContas", eExportType.PDF, param)
            Else
                Funcoes.BindReport(Me.Page, ds, "Cr_PlanoDeContas", IIf(excel, eExportType.DOC, eExportType.ExcelCrystal), param)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para tirar relatório.")
        End If
    End Sub

    Protected Sub lnkAdicionarContas_Click(sender As Object, e As EventArgs) Handles lnkAdicionarContas.Click
        Try
            If Not chkEncargo.Checked Then
                MsgBox(Me.Page, "O Relacionamento com outras contas só é permitido se a conta principal Participar no Financeiro")
                Exit Sub
            End If
            RecuperarConta()
            If Not objConta.IUD = "U" Then
                MsgBox(Me.Page, "Conta principal não foi selecionada.")
            Else
                ucConsultaPlanoDeContas.Limpar()
                ucConsultaPlanoDeContas.BindGridView(True)
                Popup.ConsultaDePlanoDeContas(Me.Page, "objContaPLxENC" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PlanoDeContas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Private Function validaCamposBacen() As Boolean
        If String.IsNullOrWhiteSpace(txtContaBacen.Text) Then
            MsgBox(Me.Page, "Informe a conta.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlRefBacen.SelectedValue) Then
            MsgBox(Me.Page, "Informe a referencia.")
        End If

        Return True
    End Function

    Private Function PreencheObjBacen(ByVal IUD As String) As [Lib].Negocio.PlanoDeContaXReferencialBacen
        Dim objPCXRefBacen As [Lib].Negocio.PlanoDeContaXReferencialBacen = New [Lib].Negocio.PlanoDeContaXReferencialBacen()

        objPCXRefBacen.IUD = IIf(IUD = "I", "I", "U")
        objPCXRefBacen.Codigo = ViewState("Codigo")
        objPCXRefBacen.Conta = txtConta.Text
        objPCXRefBacen.Referencial = ddlRefBacen.SelectedValue
        objPCXRefBacen.CodigoGrupo = IIf(Not String.IsNullOrWhiteSpace(ddlGrupoBacen.SelectedValue), ddlGrupoBacen.SelectedValue, "NULL")
        objPCXRefBacen.CodigoProduto = IIf(Not String.IsNullOrWhiteSpace(ddlProdutoBacen.SelectedValue), ddlProdutoBacen.SelectedValue, "NULL")
        objPCXRefBacen.CodigodeCustoECf = IIf(Not String.IsNullOrWhiteSpace(ddlCustoECFBacen.SelectedValue), ddlCustoECFBacen.SelectedValue, 0)
        objPCXRefBacen.CodigodeCustoECD = IIf(Not String.IsNullOrWhiteSpace(ddlCustoECDBacen.SelectedValue), ddlCustoECDBacen.SelectedValue, 0)

        Return objPCXRefBacen
    End Function

    Protected Sub lnkNovoBacen_Click(sender As Object, e As EventArgs) Handles lnkNovoBacen.Click
        Try
            If Funcoes.VerificaPermissao("PlanoDeContas", "GRAVAR") Then
                If validaCamposBacen() Then
                    Dim objPCXRefBacen As [Lib].Negocio.PlanoDeContaXReferencialBacen = PreencheObjBacen("I")

                    If objPCXRefBacen.Salvar() Then
                        MsgBox(Me.Page, "Dados inseridos com Sucesso.", eTitulo.Sucess)
                    End If
                    LimparCamposBacen(False)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizarBacen_Click(sender As Object, e As EventArgs) Handles lnkAtualizarBacen.Click
        Try
            If Funcoes.VerificaPermissao("PlanoDeContas", "ALTERAR") Then
                If validaCamposBacen() Then
                    Dim objPCXRefBacen As [Lib].Negocio.PlanoDeContaXReferencialBacen = PreencheObjBacen("U")

                    If objPCXRefBacen.Salvar() Then
                        MsgBox(Me.Page, "Dados atualizados com Sucesso.", eTitulo.Sucess)
                    End If
                    LimparCamposBacen(False)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirBacen_Click(sender As Object, e As EventArgs) Handles lnkExcluirBacen.Click
        Try
            If Funcoes.VerificaPermissao("PlanoDeContas", "EXCLUIR") Then
                If validaCamposBacen() Then
                    Dim objPCXRefBacen As New [Lib].Negocio.PlanoDeContaXReferencialBacen()
                    objPCXRefBacen.IUD = "D"
                    objPCXRefBacen.Codigo = ViewState("Codigo")

                    If objPCXRefBacen.Salvar() Then
                        MsgBox(Me.Page, "Dados excluidos com Sucesso.", eTitulo.Sucess)
                    End If
                    LimparCamposBacen(False)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LimparCamposBacen(ByVal geral As Boolean)
        ddlGrupoBacen.SelectedValue = String.Empty
        ddlProdutoBacen.Items.Clear()
        ddlRefBacen.SelectedValue = String.Empty
        ddlCustoECFBacen.SelectedValue = String.Empty
        ddlCustoECDBacen.SelectedValue = String.Empty
        ViewState("Codigo") = 0

        ddlGrupoBacen.Enabled = True
        ddlProdutoBacen.Enabled = True
        ddlRefBacen.Enabled = True

        lnkNovoBacen.Parent.Visible = True
        lnkAtualizarBacen.Parent.Visible = False
        lnkExcluirBacen.Parent.Visible = False

        If geral Then
            lnkNovoBacen.Parent.Visible = False
            txtContaBacen.Text = String.Empty
            txtTituloBacen.Text = String.Empty
            gridContaBacen.DataSource = Nothing
            gridContaBacen.DataBind()
        Else
            preencherGridContaBacen(txtConta.Text)
        End If
    End Sub

    Protected Sub lnkLimparBacen_Click(sender As Object, e As EventArgs) Handles lnkLimparBacen.Click
        Try
            LimparCamposBacen(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub carregarGrupoBacen()
        ddl.Carregar(ddlGrupoBacen, CarregarDDL.Tabela.GrupoProduto, "")
    End Sub

    Private Sub carregarCustoECFBacen()
        ddl.Carregar(ddlCustoECFBacen, CarregarDDL.Tabela.CCustoECF, "")
    End Sub

    Protected Sub ddlGrupoBacen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoBacen.SelectedIndexChanged
        Try
            ddl.Carregar(ddlProdutoBacen, CarregarDDL.Tabela.Produto, "Grupo = '" & ddlGrupoBacen.SelectedValue & "'")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub preencherGridContaBacen(ByVal Conta As String)
        Dim lst As New [Lib].Negocio.ListPlanoDeContaXReferencialBacen(Conta)

        gridContaBacen.DataSource = lst
        gridContaBacen.DataBind()
    End Sub

    Protected Sub gridContaBacen_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridContaBacen.SelectedIndexChanged
        Try
            ViewState("Codigo") = gridContaBacen.SelectedRow.Cells(1).Text
            ddlRefBacen.SelectedValue = Server.HtmlDecode(gridContaBacen.SelectedRow.Cells(2).Text)
            ddlGrupoBacen.SelectedValue = Server.HtmlDecode(gridContaBacen.SelectedRow.Cells(3).Text).Split("-")(0).Trim()
            ddlProdutoBacen.SelectedValue = Server.HtmlDecode(gridContaBacen.SelectedRow.Cells(4).Text).Split("-")(0).Trim()
            ddlCustoECFBacen.SelectedValue = IIf(Server.HtmlDecode(gridContaBacen.SelectedRow.Cells(5).Text).Equals("0"), "", gridContaBacen.SelectedRow.Cells(5).Text)

            ddlCustoECDBacen.SelectedIndex = ddlCustoECDBacen.Items.IndexOf(ddlCustoECDBacen.Items.FindByValue(gridContaBacen.SelectedRow.Cells(6).Text))

            ddlGrupoBacen.Enabled = False
            ddlProdutoBacen.Enabled = False
            ddlRefBacen.Enabled = False

            lnkNovoBacen.Parent.Visible = False
            lnkAtualizarBacen.Parent.Visible = True
            lnkExcluirBacen.Parent.Visible = True

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class