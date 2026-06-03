Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CPR
    Inherits BasePage

    Private ObjCpr As [Lib].Negocio.CPR
    Private ObjListCpr As [Lib].Negocio.ListCPR

#Region "Sessão"
    Private Sub SessionRecuperaCPR()
        ObjCpr = CType(Session("objCpr"), [Lib].Negocio.CPR)
        If ObjCpr Is Nothing Then ObjCpr = New [Lib].Negocio.CPR
    End Sub

    Private Sub SessionSalvaCPR()
        Session("objCpr") = ObjCpr
    End Sub

    Private Sub SessionRecuperaListCPR()
        ObjListCpr = CType(Session("objListCpr"), [Lib].Negocio.ListCPR)
        If ObjListCpr Is Nothing Then ObjListCpr = New [Lib].Negocio.ListCPR
    End Sub

    Private Sub SessionSalvaListCPR()
        Session("objListCpr") = ObjListCpr
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CPR", "ACESSAR") Then

                ddl.Carregar(ddlGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "")
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")
                ddl.Carregar(ddlGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "")

                WebHelpers.BindDropDownWithEnum(ddlTipoRelacao, GetType(eTipoDeRelacaoCPR), True, "-- [SELECIONE] --")


                ddl.Carregar(ddlConsultaSafra, CarregarDDL.Tabela.Safra, "")
                Dim strSQL As String
                strSQL = " SELECT TOP 1 Safra_Id " &
                         " FROM Safras  " &
                         " WHERE '" & Date.Now.Date.ToString("MM-dd-yyyy") & "' BETWEEN InicioDeSafra AND Vencimento AND Safra_Id <> 'NENHUMA' " &
                         " ORDER BY Safra_Id"
                Dim dsSafra As DataSet = Banco.ConsultaDataSet(strSQL, "Safra")
                If dsSafra IsNot Nothing AndAlso dsSafra.Tables.Count > 0 Then
                    Dim drSafra As DataRow = dsSafra.Tables(0).Rows(0)
                    ddlConsultaSafra.SelectedValue = drSafra("Safra_Id")
                Else
                End If

                tcCPR.ActiveTabIndex = 0
                tcAuxiliar.ActiveTabIndex = 0
            End If

            If Not Session("objEmpresaEmp") Is Nothing Then
                SessionRecuperaCPR()
                ObjCpr.Empresa = CType(Session("objEmpresaEmp"), [Lib].Negocio.Cliente)
                ObjCpr.CodigoEmpresa = ObjCpr.Empresa.Codigo
                ObjCpr.EndEmpresa = ObjCpr.Empresa.CodigoEndereco

                Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(ObjCpr.Empresa)
                txtEmpresa.Text = itemEmpresa.Text
                txtCodigoEmpresa.Value = itemEmpresa.Value

                Session.Remove("objEmpresaEmp")
                Session.Remove("ObjUnNegocioEmp")
            End If

            If Not Session("CPR") Is Nothing Then
                SessionRecuperaCPR()
                gridFazenda.DataSource = ObjCpr.Fazendas.ToArray
                gridFazenda.DataBind()
                gridMatricula.DataSource = Nothing
                gridMatricula.DataBind()
                gridProprietario.DataSource = Nothing
                gridProprietario.DataBind()
                Session.Remove("CPR")
                If ObjCpr.Fazendas.Count > 0 Then
                    btnCartorio.Enabled = False
                    txtCPR.Enabled = False
                    txtDataEmissao.Enabled = False
                    txtDataVencimento.Enabled = False
                    txtAreaCPR.Text = ObjCpr.Area
                End If
            End If

            If Not Session("ObjClienteCar") Is Nothing Then
                Dim car As [Lib].Negocio.Cliente = Session("ObjClienteCar")
                Dim itemCartorio As ListItem = Funcoes.FormatarListItemCliente(car)
                txtCartorio.Text = itemCartorio.Text
                txtCodigoCartorio.Value = itemCartorio.Value
                SessionRecuperaCPR()
                ObjCpr.CodigoCartorio = car.Codigo
                ObjCpr.EndCartorio = car.CodigoEndereco
                ObjCpr.Cartorio = car
                SessionSalvaCPR()
                Session.Remove("ObjClienteCar")
            End If

            Limpar()
            SetarEmpresa(Session("ssEmpresa"), Session("ssEndEmpresa"))
        Else
            MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
            Exit Sub
        End If
    End Sub

    Private Sub SetarEmpresa(ByVal Empresa As String, ByVal EndEmpresa As String)
        Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa, EndEmpresa)
        Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmpresa)

        txtEmpresa.Text = itemEmpresa.Text
        txtCodigoEmpresa.Value = itemEmpresa.Value
    End Sub

    '****************** CPR ***************************************
    Protected Sub lnkConsultarC_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultarC.Click
        Consultar()
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("CPR", "GRAVAR") Then
            If ValidaDados() Then
                NovoCPR()
                ObjCpr.IUD = "I"
                If ObjCpr.Salvar() Then
                    MsgBox(Me.Page, "CPR Salva com Sucesso.", eTitulo.Sucess)
                    Limpar()
                Else
                    MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub NovoCPR()
        SessionRecuperaCPR()
        ObjCpr.IUD = "I"
        If Not String.IsNullOrWhiteSpace(txtCodigoEmpresa.Value) Then
            Dim Emp As String() = txtCodigoEmpresa.Value.Split("-")
            ObjCpr.CodigoEmpresa = Emp(0)
            ObjCpr.EndEmpresa = Emp(1)
        Else
            ObjCpr.CodigoEmpresa = 0
            ObjCpr.EndEmpresa = ""
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoCartorio.Value) Then
            Dim Car As String() = txtCodigoCartorio.Value.Split("-")
            ObjCpr.CodigoCartorio = Car(0)
            ObjCpr.EndCartorio = Car(1)
        Else
            ObjCpr.CodigoCartorio = 0
            ObjCpr.EndCartorio = ""
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            Dim cli As String() = txtCodigoCliente.Value.Split("-")
            ObjCpr.CodigoCliente = cli(0)
            ObjCpr.EndCliente = cli(1)
        Else
            ObjCpr.CodigoCliente = 0
            ObjCpr.EndCliente = ""
        End If

        If txtCodigoEndosso.Value = "" Then
            ObjCpr.CodigoEndossado = ""
            ObjCpr.EndEndossado = 0
        Else
            Dim Endo As String() = txtCodigoEndosso.Value.Split("-")
            ObjCpr.CodigoEndossado = Endo(0)
            ObjCpr.EndEndossado = Endo(1)
        End If

        ObjCpr.CodigoCPR = txtCPR.Text
        ObjCpr.CodigoProduto = ddlProduto.SelectedValue
        ObjCpr.Observacao = txtObservacao.Text
        ObjCpr.CodigoSituacao = ddlSituacao.SelectedValue
        ObjCpr.Registro = txtRegistro.Text
        ObjCpr.DataRegistro = CDate(txtDataRegistro.Text)
        ObjCpr.Quantidade = txtQtde.Text
        ObjCpr.Produtividade = txtProdutividade.Text
        ObjCpr.CodigoSafra = ddlSafra.SelectedValue

        ObjCpr.DataVencimento = CDate(txtDataVencimento.Text)


        SessionSalvaCPR()
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        If Funcoes.VerificaPermissao("CPR", "ALTERAR") Then
            If ValidaDados() Then
                SessionRecuperaCPR()

                Dim Car As String() = txtCodigoCartorio.Value.Split("-")
                ObjCpr.CodigoCartorio = Car(0)
                ObjCpr.EndCartorio = Car(1)

                Dim Cli As String() = txtCodigoCliente.Value.Split("-")
                ObjCpr.CodigoCliente = Cli(0)
                ObjCpr.EndCliente = Cli(1)


                If String.IsNullOrWhiteSpace(txtCodigoEndosso.Value) Then
                    ObjCpr.CodigoEndossado = ""
                    ObjCpr.EndEndossado = 0
                Else
                    Dim Endo As String() = txtCodigoEndosso.Value.Split("-")
                    ObjCpr.CodigoEndossado = Endo(0)
                    ObjCpr.EndEndossado = Endo(1)
                End If

                ObjCpr.CodigoCPR = txtCPR.Text
                ObjCpr.CodigoProduto = ddlProduto.SelectedValue
                ObjCpr.Observacao = txtObservacao.Text
                ObjCpr.CodigoSituacao = ddlSituacao.SelectedValue
                ObjCpr.Registro = txtRegistro.Text
                ObjCpr.DataRegistro = CDate(txtDataRegistro.Text)
                ObjCpr.Quantidade = txtQtde.Text
                ObjCpr.Produtividade = txtProdutividade.Text
                ObjCpr.CodigoSafra = ddlSafra.SelectedValue
                ObjCpr.IUD = "U"
                If ObjCpr.Salvar() Then
                    MsgBox(Me.Page, "CPR Alterada com Sucesso.", eTitulo.Sucess)
                    Limpar()
                Else
                    MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar Registro")
        End If

    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        If Funcoes.VerificaPermissao("CPR", "EXCLUIR") Then
            SessionRecuperaCPR()
            Dim Car As String() = txtCodigoCartorio.Value.Split("-")
            ObjCpr.IUD = "D"
            If ObjCpr.Salvar() Then
                MsgBox(Me.Page, "CPR Excluída com Sucesso.", eTitulo.Sucess)
                Limpar()
            Else
                MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click

    End Sub

    Protected Sub gridConsulta_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles gridConsulta.SelectedIndexChanged
        SessionRecuperaListCPR()
        Dim CPR As String = gridConsulta.SelectedRow.Cells(4).Text
        Dim CodigoCartorio As String = gridConsulta.SelectedRow.Cells(1).Text
        Dim EndCartorio As Integer = gridConsulta.SelectedRow.Cells(2).Text

        If Session("VincularCPR") = False Then
            ObjCpr = (From c In ObjListCpr Where c.CodigoCPR = CPR And c.CodigoCartorio = CodigoCartorio And c.EndCartorio = EndCartorio Select c).Single
            txtDataLiquidacao.Text = Date.Now.ToString("dd/MM/yyyy")

            ObjCpr.IUD = "U"

            SessionSalvaCPR()
            AtualizarComAClasse()
            tbLiquidacao.Visible = String.IsNullOrWhiteSpace(ObjCpr.CautelaCodigoCPR)
            tcCPR.ActiveTabIndex = 0
        Else
            If Not String.IsNullOrWhiteSpace(txtCPR.Text) Then
                SessionRecuperaCPR()
                If Not CPRJaEstaVinculada(CPR, CodigoCartorio, EndCartorio) Then
                    ObjCpr.DevedorSolidarioCodigoCPR = CPR
                    ObjCpr.DevedorSolidarioCodigoCartorio = CodigoCartorio
                    ObjCpr.DevedorSolidarioEndCartorio = EndCartorio
                    lblDevedorSolidarioCPR.Text = ObjCpr.DevedorSolidarioCodigoCPR
                    lblDescDevedorSolidario.Text = ObjCpr.DevedorSolidarioCPRCautela.Cliente.Nome & " - " & ObjCpr.DevedorSolidarioCPRCautela.EndCliente
                    SessionSalvaCPR()
                    tcCPR.ActiveTabIndex = 0
                Else
                    MsgBox(Me.Page, "Esta CPR já está vinculada a outra!")
                End If
            End If

        End If
    End Sub

    Protected Sub ddlGrupoProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, " Grupo = '" & ddlGrupoProduto.SelectedValue & "'")
    End Sub

    Protected Sub BtnConsultaPrd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnConsultaPrd.Click
        ucConsultaProduto.Limpar()
        Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
        Popup.ConsultaDeProduto(Me.Page, "objProdutoCPR" & HID.Value, txtNome.ClientID, True)
    End Sub

    Protected Sub ddlProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim p As New [Lib].Negocio.Produto(ddlProduto.SelectedValue)
        lblProduto.Text = p.Unidade
    End Sub

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmpresa.Click
        ucConsultaEmpresas.Limpar()
        Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaCPR" & HID.Value)
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objEmpresaCPR" & HID.Value) Is Nothing Then
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objEmpresaCPR" & HID.Value), [Lib].Negocio.Cliente))
            txtEmpresa.Text = itemEmpresa.Text
            txtCodigoEmpresa.Value = itemEmpresa.Value
            Session.Remove("objEmpresaCPR" & HID.Value)
        ElseIf Not Session("objFazenda" & HID.Value) Is Nothing Then
            SessionRecuperaCPR()
            If ObjCpr.Fazendas.Count > 0 Then
                txtAreaCPR.Text = ObjCpr.Area
                gridFazenda.DataSource = ObjCpr.Fazendas
                gridFazenda.DataBind()
                If Not String.IsNullOrWhiteSpace(txtQtde.Text) And Not String.IsNullOrWhiteSpace(txtAreaCPR.Text) Then
                    SessionRecuperaCPR()
                    ObjCpr.Quantidade = txtQtde.Text
                    txtProdutividade.Text = ObjCpr.ProdutividadeCalculada
                    SessionSalvaCPR()
                End If
            End If
            Session.Remove("objFazenda" & HID.Value.ToString)
        ElseIf Not Session("objCPRxFazenda" & HID.Value) Is Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objCPRxFazenda" & HID.Value.ToString), [Lib].Negocio.Cliente)
            txtFazenda.Text = pCliente.Complemento
            txtCodigoFazenda.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
            Session.Remove("objCPRxFazenda" & HID.Value.ToString)
        ElseIf Not Session("objCPRxProprietario" & HID.Value) Is Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objCPRxProprietario" & HID.Value.ToString), [Lib].Negocio.Cliente)
            txtProprietario.Text = pCliente.Nome
            txtcodigoProprietario.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
            Session.Remove("objCPRxProprietario" & HID.Value.ToString)
        ElseIf Not Session("objCPRxCartorio" & HID.Value) Is Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objCPRxCartorio" & HID.Value.ToString), [Lib].Negocio.Cliente)
            txtCartorio.Text = pCliente.Nome
            txtCodigoCartorio.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
            Session.Remove("objCPRxCartorio" & HID.Value.ToString)
        ElseIf Not Session("objCPRxCliente" & HID.Value) Is Nothing Then
            SessionRecuperaCPR()
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objCPRxCliente" & HID.Value.ToString), [Lib].Negocio.Cliente)
            If Session("ConsultaCliente") = "S" Then
                txtConsultaCliente.Text = pCliente.Nome
                txtConsultaCodigoCliente.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
            Else
                txtCliente.Text = pCliente.Nome
                txtCodigoCliente.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
                ObjCpr.CodigoCliente = pCliente.Codigo
                ObjCpr.EndCliente = pCliente.CodigoEndereco
            End If
            Session.Remove("objCPRxCliente" & HID.Value.ToString)
        ElseIf Not Session("objCPRxEndosso" & HID.Value) Is Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objCPRxEndosso" & HID.Value.ToString), [Lib].Negocio.Cliente)
            txtEndosso.Text = pCliente.Nome
            txtCodigoEndosso.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
            Session.Remove("objCPRxEndosso" & HID.Value.ToString)
        ElseIf Not Session("objCPRxAvalista" & HID.Value) Is Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objCPRxAvalista" & HID.Value.ToString), [Lib].Negocio.Cliente)
            txtAvalista.Text = pCliente.Nome
            txtCodigoAvalista.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
            Session.Remove("objCPRxAvalista" & HID.Value.ToString)
        ElseIf Not Session("objCPRxGrau" & HID.Value) Is Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objCPRxGrau" & HID.Value.ToString), [Lib].Negocio.Cliente)
            txtGrau.Text = pCliente.Nome
            txtCodigoGrau.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
            Session.Remove("objCPRxGrau" & HID.Value.ToString)
        ElseIf Session("objCPRxProduto" & HID.Value) IsNot Nothing Then
            Dim pProduto As [Lib].Negocio.Produto = CType(Session("objCPRxProduto" & HID.Value.ToString), [Lib].Negocio.Produto)
            ddlProduto.SelectedValue = pProduto.Descricao
            ddlProduto.SelectedIndex = pProduto.Codigo
            Session.Remove("objCPRxProduto" & HID.Value.ToString)
        ElseIf Not Session("objProdutoCPR" & HID.Value) Is Nothing Then
            SessionRecuperaCPR()
            Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoCPR" & HID.Value)
            ddlGrupoProduto.SelectedValue = objProduto.CodigoGrupo
            ddlGrupoProduto_SelectedIndexChanged(Nothing, Nothing)
            ObjCpr.CodigoProduto = objProduto.Codigo
            ddlGrupoProduto.SelectedValue = objProduto.CodigoGrupo
            ddlProduto.SelectedValue = objProduto.Codigo
            lblProduto.Text = objProduto.Unidade
            SessionSalvaCPR()
            Session.Remove("objProdutoCPR" & HID.Value)
        End If
    End Sub

    Protected Sub btnCartorio_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCartorio.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objCPRxCartorio" & HID.Value.ToString, "txtCartorio")
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCliente.Click
        ucConsultaClientes.Limpar()
        Session("ConsultaCliente") = "N"
        Popup.ConsultaDeClientes(Me, "objCPRxCliente" & HID.Value.ToString, "txtCliente")
    End Sub

    Protected Sub btnConsultaCliente_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultaCliente.Click
        Session("ConsultaCliente") = "S"
        Popup.ConsultaDeClientes(Me, "objCPRxCliente" & HID.Value.ToString, "txtConsultaCliente")
    End Sub

    Protected Sub txtDataEmissao_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessionRecuperaCPR()
        If IsDate(txtDataEmissao.Text) Then
            ObjCpr.DataEmissao = CDate(txtDataEmissao.Text)
        Else
            ObjCpr.DataEmissao = Now.Date
        End If
        SessionSalvaCPR()
    End Sub

    Private Sub AtualizarComAClasse()
        SessionRecuperaCPR()
        If Not ObjCpr.Empresa Is Nothing Then
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(ObjCpr.Empresa)
            txtEmpresa.Text = itemEmpresa.Text
            txtCodigoEmpresa.Value = itemEmpresa.Value
        Else
            txtEmpresa.Text = ""
            txtCodigoEmpresa.Value = ""
        End If

        If Not ObjCpr.Cartorio Is Nothing Then
            Dim itemCartorio As ListItem = Funcoes.FormatarListItemCliente(ObjCpr.Cartorio)
            txtCartorio.Text = itemCartorio.Text
            txtCodigoCartorio.Value = itemCartorio.Value
        Else
            txtCartorio.Text = ""
            txtCodigoCartorio.Value = ""
        End If

        If Not ObjCpr.Cliente Is Nothing Then
            txtCliente.Text = ObjCpr.Cliente.Nome
            txtCodigoCliente.Value = ObjCpr.CodigoCliente & "-" & ObjCpr.EndCliente
        Else
            txtCliente.Text = ""
            txtCodigoCliente.Value = ""
        End If

        If Not ObjCpr.Endossado Is Nothing Then
            Dim itemEndosso As ListItem = Funcoes.FormatarListItemCliente(ObjCpr.Endossado)
            txtEndosso.Text = itemEndosso.Text
            txtCodigoEndosso.Value = itemEndosso.Value
        Else
            txtEndosso.Text = ""
            txtCodigoEndosso.Value = ""
        End If

        txtCPR.Text = ObjCpr.CodigoCPR

        ddlSafra.SelectedValue = ObjCpr.CodigoSafra

        If ObjCpr.CodigoProduto.Length > 0 Then
            ddlGrupoProduto.SelectedValue = ObjCpr.Produto.CodigoGrupo
            ddlGrupoProduto_SelectedIndexChanged(Nothing, Nothing)
            ddlProduto.SelectedValue = ObjCpr.CodigoProduto
            lblProduto.Text = ObjCpr.Produto.Unidade
        Else
            ddlGrupoProduto.SelectedIndex = 0
            ddlGrupoProduto_SelectedIndexChanged(Nothing, Nothing)
            ddlProduto.SelectedIndex = 0
            lblProduto.Text = ""
        End If

        txtObservacao.Text = ObjCpr.Observacao
        ddlSituacao.SelectedValue = ObjCpr.CodigoSituacao
        txtDataEmissao.Text = ObjCpr.DataEmissao.ToString("dd/MM/yyyy")
        txtRegistro.Text = ObjCpr.Registro
        txtDataRegistro.Text = ObjCpr.DataRegistro.ToString("dd/MM/yyyy")
        txtDataVencimento.Text = ObjCpr.DataVencimento.ToString("dd/MM/yyyy")
        txtQtde.Text = ObjCpr.Quantidade
        txtProdutividade.Text = ObjCpr.Produtividade
        txtAreaCPR.Text = ObjCpr.Area

        If Not String.IsNullOrWhiteSpace(ObjCpr.DevedorSolidarioCodigoCPR) Then
            lblDevedorSolidarioCPR.Text = ObjCpr.DevedorSolidarioCodigoCPR
            lblDescDevedorSolidario.Text = ObjCpr.DevedorSolidarioCPRCautela.Cliente.Nome & " - " & ObjCpr.DevedorSolidarioCPRCautela.EndCliente
        Else
            lblDevedorSolidarioCPR.Text = String.Empty
            lblDescDevedorSolidario.Text = String.Empty
        End If

        If ObjCpr.Fazendas.Count > 0 Then
            If gridFazenda.SelectedIndex - 1 <= ObjCpr.Fazendas.Count Then
                Dim i As Integer = gridFazenda.SelectedIndex
                gridFazenda.DataSource = ObjCpr.Fazendas.ToArray
                gridFazenda.DataBind()
                gridFazenda.SelectedIndex = i
            Else
                gridFazenda.SelectedIndex = 0
            End If
            gridFazenda_SelectedIndexChanged(gridFazenda, Nothing)
        Else
            gridFazenda.DataSource = ObjCpr.Fazendas.ToArray
            gridFazenda.DataBind()
        End If

        gridAvalista.DataSource = ObjCpr.Avalistas.ToArray
        gridAvalista.DataBind()

        gridGrau.DataSource = ObjCpr.Graus.ToArray
        gridGrau.DataBind()

        GrdLiquidacao.DataSource = ObjCpr.Liquidacoes
        GrdLiquidacao.DataBind()

        SessionSalvaCPR()
    End Sub

    Private Function ValidaDados() As Boolean
        SessionRecuperaCPR()

        If ddlSituacao.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Situação")
            Return False
        End If

        If txtEmpresa.Text.Length = 0 Then
            MsgBox(Me.Page, "Informe a Empresa")
            Return False
        End If

        If txtCartorio.Text.Length = 0 Then
            MsgBox(Me.Page, "Informe o Cartório")
            Return False
        End If

        If txtCliente.Text.Length = 0 Then
            MsgBox(Me.Page, "Informe o Cliente")
            Return False
        End If

        If txtCPR.Text.Length = 0 Then
            MsgBox(Me.Page, "Informe o Código da CPR")
            Return False
        End If

        If ddlSafra.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Safra")
            Return False
        End If

        If ddlProduto.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe o Produto")
            Return False
        End If

        If Not IsDate(txtDataEmissao.Text) Then
            MsgBox(Me.Page, "Informe a Data de Emissão")
            Return False
        End If

        If Not IsNumeric(txtRegistro.Text) Then
            MsgBox(Me.Page, "Informe o código do Registro")
            Return False
        End If

        If Not IsDate(txtDataRegistro.Text) Then
            MsgBox(Me.Page, "Informe a Data do Registro")
            Return False
        End If

        If Not IsDate(txtDataVencimento.Text) Then
            MsgBox(Me.Page, "Informe a Data do Vencimento")
            Return False
        End If

        If txtQtde.Text.Length = 0 OrElse CDec(txtQtde.Text) = 0 Then
            MsgBox(Me.Page, "Informe a Quantidade da CPR")
            Return False
        End If

        If txtProdutividade.Text.Length = 0 OrElse CDec(txtProdutividade.Text) = 0 Then
            MsgBox(Me.Page, "Informe a Produtividade")
            Return False
        End If

        If ObjCpr.Area = 0 Then
            MsgBox(Me.Page, "Informe a(s) Área(s) da CPR")
            Return False
        End If

        'If ObjCpr.Avalistas.Count = 0 Then
        '    MsgBox(Me.Page, "Informe pelo menos 1 Avalista / Devedor!")
        '    Return False
        'End If

        If ObjCpr.Graus.Count = 0 Then
            MsgBox(Me.Page, "Informe o Grau da CPR")
            Return False
        End If

        Return True
    End Function

    Private Sub Limpar()
        ObjCpr = New [Lib].Negocio.CPR
        ObjCpr.IUD = ""
        ObjCpr.CodigoSituacao = 1
        ObjCpr.DataEmissao = Now.Date
        ObjCpr.DataRegistro = Now.Date
        ObjCpr.DataVencimento = Now.Date.AddYears(1)

        btnCartorio.Enabled = True
        txtCPR.Enabled = True
        txtDataEmissao.Enabled = True
        txtDataVencimento.Enabled = True

        txtAvalista.Text = ""
        txtCodigoAvalista.Value = ""

        txtGrau.Text = ""
        txtCodigoGrau.Value = ""

        gridMatricula.DataSource = Nothing
        gridMatricula.DataBind()

        gridProprietario.DataSource = Nothing
        gridProprietario.DataBind()

        gridGrau.DataSource = Nothing
        gridGrau.DataBind()

        GrdLiquidacao.DataSource = Nothing
        GrdLiquidacao.DataBind()

        SessionSalvaCPR()
        AtualizarComAClasse()

        Session("VincularCPR") = False

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaEmpresas.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)
        ucConsultaFazendaCPR.SetarHID(HID.Value)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            btnEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub btnConsultaFazenda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConsultaFazenda.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objCPRxFazenda" & HID.Value.ToString, "txtFazenda")
    End Sub

    Protected Sub btnConsultaProprietario_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objCPRxProprietario" & HID.Value.ToString, "txtProprietario")
    End Sub

    Private Sub Consultar()
        ' If Funcoes.VerificaPermissao("CPR", "LEITURA") Then
        'SessionRecuperaCPR()

        Dim Parametros As New Hashtable

        ObjCpr = New [Lib].Negocio.CPR()

        If Not String.IsNullOrWhiteSpace(txtConsultaRegistro.Text) Then
            ObjCpr.Registro = txtConsultaRegistro.Text
        Else
            ObjCpr.Registro = 0
        End If

        If txtConsultaCPR.Text.Trim.Length > 0 Then
            ObjCpr.CodigoCPR = txtConsultaCPR.Text
        Else
            ObjCpr.CodigoCPR = String.Empty
        End If

        If txtConsultaCliente.Text.Trim.Length > 0 Then
            ObjCpr.CodigoCliente = txtConsultaCodigoCliente.Value.Split("-")(0)
            ObjCpr.EndCliente = txtConsultaCodigoCliente.Value.Split("-")(1)

        Else
            ObjCpr.CodigoCliente = String.Empty
        End If


        'If ddlProduto.SelectedIndex > 0 Then
        '    ObjCpr.CodigoProduto = ddlProduto.SelectedValue
        'Else
        '    ObjCpr.CodigoProduto = ""
        '    If ddlGrupoProduto.SelectedIndex > 0 Then Parametros.Add("GrupoProduto", ddlGrupoProduto.SelectedValue)
        'End If

        'If txtObservacao.Text.Trim.Length > 0 Then
        '    ObjCpr.Observacao = txtObservacao.Text
        'Else
        '    ObjCpr.Observacao = ""
        'End If

        'If ddlSituacao.SelectedIndex > 0 Then
        '    ObjCpr.CodigoSituacao = ddlSituacao.SelectedValue
        'Else
        '    ObjCpr.CodigoSituacao = -1
        'End If

        'If txtRegistro.Text.Trim.Length = 0 Then
        '    ObjCpr.Registro = txtaRegistro.Text
        'Else
        '    ObjCpr.Registro = 0
        'End If

        'If IsNumeric(txtQtde.Text) Then
        '    ObjCpr.Quantidade = txtQtde.Text
        'Else
        '    ObjCpr.Quantidade = 0
        'End If

        'If IsNumeric(txtProdutividade.Text) Then
        '    ObjCpr.Produtividade = txtProdutividade.Text
        'Else
        '    ObjCpr.Produtividade = 0
        'End If

        If ddlConsultaSafra.SelectedIndex > 0 Then
            ObjCpr.CodigoSafra = ddlConsultaSafra.SelectedValue
        End If


        If txtcodigoProprietario.Value.Length > 0 Then
            Parametros.Add("Proprietario", txtcodigoProprietario.Value)
            Parametros.Add("Consolidar", chkConsolidar.Checked)
        End If

        If txtCodigoFazenda.Value.Length > 0 Then Parametros.Add("Fazenda", txtCodigoFazenda.Value)
        If txtConsultaMatricula.Text.Trim.Length > 0 Then Parametros.Add("Matricula", txtConsultaMatricula.Text)
        If txtConsultaGrau.Text.Trim.Length > 0 Then Parametros.Add("Grau", txtConsultaGrau.Text)
        If IsDate(txtDataEmissaoDe.Text) And IsDate(txtDataEmissaoAte.Text) Then Parametros.Add("Emissao", CDate(txtDataEmissaoDe.Text).ToString("yyyy-MM-dd") & "|" & CDate(txtDataEmissaoAte.Text).ToString("yyyy-MM-dd"))
        If IsDate(txtDataVencimentoDe.Text) And IsDate(txtDataVencimentoAte.Text) Then Parametros.Add("Vencimento", CDate(txtDataVencimentoDe.Text).ToString("yyyy-MM-dd") & "|" & CDate(txtDataVencimentoAte.Text).ToString("yyyy-MM-dd"))

        If Session("VincularCPR") Then
            Parametros.Add("ExcessaoCodigoCPR", txtCPR.Text)
            Parametros.Add("ExcessaoCodigoCartorio", txtCodigoCartorio.Value.Split("-")(0))
            Parametros.Add("ExcessaoEndCartorio", txtCodigoCartorio.Value.Split("-")(1))
            ObjCpr.CodigoSafra = ddlSafra.SelectedValue
            ObjCpr.CodigoSituacao = 1 'Abertas
        End If

        ObjListCpr = New [Lib].Negocio.ListCPR(ObjCpr, Parametros)
        gridConsulta.DataSource = ObjListCpr.ToArray
        gridConsulta.DataBind()
        SessionSalvaListCPR()

        'If ObjListCpr.Count = 1 Then
        '    ObjCpr = ObjListCpr(0)
        '    SessionSalvaCPR()
        '    AtualizarComAClasse()
        '    'TC_Principal.ActiveTabIndex = 0
        'End If
        'End If

    End Sub

    Private Sub LimparCamposConsulta()
        txtFazenda.Text = ""
        txtCodigoFazenda.Value = ""
        txtProprietario.Text = ""
        txtcodigoProprietario.Value = ""
        txtConsultaCliente.Text = ""
        txtConsultaCodigoCliente.Value = ""
        txtDataEmissaoDe.Text = ""
        txtDataEmissaoAte.Text = ""
        txtConsultaMatricula.Text = ""
        chkConsolidar.Checked = False
        txtConsultaGrau.Text = ""
        txtDataVencimentoDe.Text = ""
        txtDataVencimentoAte.Text = ""
        txtConsultaCPR.Text = String.Empty
        txtConsultaRegistro.Text = String.Empty
        Session("VincularCPR") = False
        'Limpar()
    End Sub

    Protected Sub lnkLimparC_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparC.Click
        LimparCamposConsulta()
    End Sub

    '****************** FAZENDA ***************************************
    Protected Sub btnAdicionarFazenda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAdicionarFazenda.Click
        Try
            If txtCartorio.Text.Length <> 0 Or txtCPR.Text.Length <> 0 Then
                SessionRecuperaCPR()
                If ObjCpr Is Nothing Then
                    NovoCPR()
                End If

                ucConsultaFazendaCPR.Limpar()
                If ucConsultaFazendaCPR.Consultar() > 0 Then
                    Popup.ConsultaDeFazendaCPR(Me, "objFazenda" & HID.Value.ToString)

                Else
                    MsgBox(Me.Page, "Não foram encontradas matrículas para o cliente da CPR!")
                End If
            Else
                MsgBox(Me.Page, "Informe o Cartório, o número da CPR e as datas de início e vencimento da CPR")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub gridFazenda_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessionRecuperaCPR()
        If gridFazenda.SelectedIndex = -1 Then Exit Sub
        gridMatricula.DataSource = ObjCpr.Fazendas(gridFazenda.SelectedIndex).Matriculas
        gridMatricula.DataBind()
        gridProprietario.DataSource = ObjCpr.Fazendas(gridFazenda.SelectedIndex).Clientes.ToArray
        gridProprietario.DataBind()
    End Sub

    Protected Sub gridFazenda_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridFazenda.RowDataBound
        If e.Row.RowType = DataControlRowType.Footer Then
            SessionRecuperaCPR()
            If Not IsNothing(ObjCpr.Fazendas) Then
                e.Row.Cells(4).Text = String.Format("{0}", ObjCpr.Fazendas.AreaTotal)
                e.Row.Cells(4).HorizontalAlign = HorizontalAlign.Right
            End If
        End If
    End Sub

    Protected Sub gridFazenda_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs)
        SessionRecuperaCPR()
        If ObjCpr.IUD = "U" Then
            ObjCpr.Fazendas(e.RowIndex).IUD = "D"
            If Not ObjCpr.Fazendas(e.RowIndex).Salvar() Then
                ObjCpr.Fazendas(e.RowIndex).IUD = ""
                MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                Exit Sub
            End If
        End If

        ObjCpr.Fazendas.RemoveAt(e.RowIndex)

        If ObjCpr.Fazendas.Count = 0 Then
            btnCartorio.Enabled = True
            txtCPR.Enabled = True
            txtDataEmissao.Enabled = True
            txtDataVencimento.Enabled = True
        Else
            gridFazenda.SelectedIndex = 0
            gridFazenda_SelectedIndexChanged(gridFazenda, Nothing)
        End If
        gridFazenda.DataSource = ObjCpr.Fazendas
        gridFazenda.DataBind()
        gridMatricula.DataSource = Nothing
        gridMatricula.DataBind()
        gridProprietario.DataSource = Nothing
        gridProprietario.DataBind()
        txtAreaCPR.Text = ObjCpr.Area

        SessionSalvaCPR()
    End Sub

    Protected Sub gridMatricula_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs)
        SessionRecuperaCPR()

        If ObjCpr.IUD = "U" Then
            If ObjCpr.Fazendas(gridFazenda.SelectedIndex).Matriculas.Count = 1 Then
                ObjCpr.Fazendas(gridFazenda.SelectedIndex).IUD = "D"
                If Not ObjCpr.Fazendas(gridFazenda.SelectedIndex).Salvar() Then
                    ObjCpr.Fazendas(gridFazenda.SelectedIndex).IUD = ""
                    MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    Exit Sub
                End If
            Else
                ObjCpr.Fazendas(gridFazenda.SelectedIndex).Matriculas(e.RowIndex).IUD = "D"
                If Not ObjCpr.Fazendas(gridFazenda.SelectedIndex).Matriculas(e.RowIndex).Salvar() Then
                    ObjCpr.Fazendas(gridFazenda.SelectedIndex).Matriculas(e.RowIndex).IUD = ""
                    MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    Exit Sub
                End If
            End If
        End If

        If ObjCpr.Fazendas(gridFazenda.SelectedIndex).Matriculas.Count = 1 Then
            gridFazenda.DataSource = ObjCpr.Fazendas.ToArray
            gridFazenda.DataBind()
            gridFazenda.SelectedIndex = 0
            gridFazenda_SelectedIndexChanged(gridFazenda, Nothing)
        Else
            ObjCpr.Fazendas(gridFazenda.SelectedIndex).Matriculas.RemoveAt(e.RowIndex)
            gridMatricula.DataSource = ObjCpr.Fazendas(gridFazenda.SelectedIndex).Matriculas.ToArray
            gridMatricula.DataBind()
        End If

        gridFazenda.DataSource = ObjCpr.Fazendas.ToArray
        gridFazenda.DataBind()
        txtAreaCPR.Text = ObjCpr.Area
        SessionSalvaCPR()
    End Sub

    '************* AVALISTA ****************************************************
    Protected Sub btnAvalista_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAvalista.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objCPRxAvalista" & HID.Value.ToString, "txtAvalista")
    End Sub

    Protected Sub btnAdicionarAvalista_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdicionarAvalista.Click
        If String.IsNullOrWhiteSpace(txtCodigoAvalista.Value) Then
            MsgBox(Me.Page, "Indique um Avalista / Devedor!")
        ElseIf ddlTipoRelacao.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Indique um tipo de Relação!")
            ddlTipoRelacao.Focus()
        Else
            SessionRecuperaCPR()

            Dim Ava As String() = txtCodigoAvalista.Value.Split("-")
            Dim CxA As New [Lib].Negocio.CPRxAvalista(ObjCpr)

            CxA.CodigoAvalista = Ava(0)
            CxA.EndAvalista = Ava(1)
            CxA.TipoDeRelacao = CType(CInt(ddlTipoRelacao.SelectedIndex), eTipoDeRelacaoCPR)

            If Not ObjCpr.Avalistas.JaExiste(CxA) Then
                CxA.IUD = "I"
                ObjCpr.Avalistas.Add(CxA)
                gridAvalista.DataSource = ObjCpr.Avalistas
                gridAvalista.DataBind()
                SessionSalvaCPR()
                txtAvalista.Text = ""
                txtCodigoAvalista.Value = 0
            Else
                MsgBox(Me.Page, "Avalista Já Cadastrado!")
            End If
        End If
    End Sub

    Protected Sub gridAvalista_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs)
        SessionRecuperaCPR()
        If ObjCpr.IUD = "U" Then
            ObjCpr.Avalistas(e.RowIndex).IUD = "D"
            If Not ObjCpr.Avalistas(e.RowIndex).Salvar() Then
                ObjCpr.Avalistas(e.RowIndex).IUD = ""
                MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                Exit Sub
            End If
        End If
        ObjCpr.Avalistas.RemoveAt(e.RowIndex)

        gridAvalista.DataSource = ObjCpr.Avalistas.ToArray
        gridAvalista.DataBind()

        SessionSalvaCPR()
    End Sub

    '************* GRAU ****************************************************
    Protected Sub btnGrau_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGrau.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objCPRxGrau" & HID.Value.ToString, "txtGrau")
    End Sub

    Protected Sub btnAdicionarGrau_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdicionarGrau.Click
        SessionRecuperaCPR()

        Dim Gra As String() = txtCodigoGrau.Value.Split("-")
        Dim CxG As New [Lib].Negocio.CPRxGrau(ObjCpr)

        CxG.CodigoCliente = Gra(0)
        CxG.EndCliente = Gra(1)
        CxG.Grau = gridGrau.Rows.Count + 1

        If Not ObjCpr.Graus.JaExiste(CxG) Then
            CxG.IUD = "I"
            ObjCpr.Graus.Add(CxG)
            gridGrau.DataSource = ObjCpr.Graus
            gridGrau.DataBind()
            SessionSalvaCPR()
        Else
            MsgBox(Me.Page, "Cliente Já Cadastrado")
        End If
    End Sub

    Protected Sub gridGrau_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles gridGrau.RowDeleting
        SessionRecuperaCPR()
        If ObjCpr.IUD = "U" Then
            ObjCpr.Graus(e.RowIndex).IUD = "D"
            If Not ObjCpr.Graus(e.RowIndex).Salvar() Then
                ObjCpr.Graus(e.RowIndex).IUD = ""
                MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                Exit Sub
            End If
        End If
        ObjCpr.Graus.RemoveAt(e.RowIndex)

        gridGrau.DataSource = ObjCpr.Graus.ToArray
        gridGrau.DataBind()

        SessionSalvaCPR()
    End Sub

    '************* ENDOSSO ****************************************************
    Protected Sub btnEndosso_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnEndosso.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objCPRxEndosso" & HID.Value.ToString, "txtEndosso")
    End Sub

    Protected Sub ImdVincularOutraCPR_Click(ByVal sender As Object, ByVal e As EventArgs)
        If Not String.IsNullOrWhiteSpace(txtCPR.Text) Then
            SessionRecuperaCPR()
            Dim CPR As String = gridConsulta.SelectedRow.Cells(4).Text
            Dim CodigoCartorio As String = gridConsulta.SelectedRow.Cells(1).Text
            Dim EndCartorio As Integer = gridConsulta.SelectedRow.Cells(2).Text

            If Not CPRJaEstaVinculada(CPR, CodigoCartorio, EndCartorio) Then
                ObjCpr.DevedorSolidarioCodigoCPR = CPR
                ObjCpr.DevedorSolidarioCodigoCartorio = CodigoCartorio
                ObjCpr.DevedorSolidarioEndCartorio = EndCartorio
                lblDevedorSolidarioCPR.Text = ObjCpr.DevedorSolidarioCodigoCPR
                lblDescDevedorSolidario.Text = ObjCpr.DevedorSolidarioCPRCautela.Cliente.Nome & " - " & ObjCpr.DevedorSolidarioCPRCautela.EndCliente
                SessionSalvaCPR()
                tcCPR.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Esta CPR já está vinculada a outra!")
            End If
        End If
    End Sub

    Private Function CPRJaEstaVinculada(ByVal CPR As String, ByVal CodigoCartorio As String, ByVal EndCartorio As Integer) As Boolean
        SessionRecuperaListCPR()
        For Each ObjCPRTemp As [Lib].Negocio.CPR In ObjListCpr
            If ObjCPRTemp.DevedorSolidarioCodigoCPR = CPR And ObjCPRTemp.DevedorSolidarioCodigoCartorio = CodigoCartorio And ObjCPRTemp.DevedorSolidarioEndCartorio = EndCartorio Then
                Return True
                Exit Function
            End If
        Next
        Return False
    End Function

    Protected Sub lnkDevedorSolidarioCPR_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkDevedorSolidarioCPR.Click
        SessionRecuperaCPR()
        If Not ObjCpr.CodigoSituacao = 0 Then
            Session("VincularCPR") = True
            Consultar()
            tcCPR.ActiveTabIndex = 1
        Else
            MsgBox(Me.Page, "Não é possível colocar outra CPR como cautela em uma CPR já Liquidada!")
        End If

    End Sub

    Protected Sub btnAdicionarLiquidacao_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAdicionarLiquidacao.Click
        If CType(txtQtdLiquidacao.Text, Decimal) > 0 Then

            SessionRecuperaCPR()
            Dim objLiquidacao As New [Lib].Negocio.CPRxLiquidacao(ObjCpr)
            objLiquidacao.Data = txtDataLiquidacao.Text
            objLiquidacao.Quantidade = txtQtdLiquidacao.Text
            objLiquidacao.Liquidacao = GrdLiquidacao.Rows.Count + 1
            objLiquidacao.IUD = "I"
            ObjCpr.Liquidacoes.Add(objLiquidacao)

            If ObjCpr.Liquidacoes.QuantidadeTotalLiquidada <= CType(txtQtde.Text, Decimal) Then
                GrdLiquidacao.DataSource = ObjCpr.Liquidacoes
                GrdLiquidacao.DataBind()
                txtQtdLiquidacao.Text = String.Empty
                SessionSalvaCPR()
            Else
                ObjCpr.Liquidacoes.Remove(objLiquidacao)
                MsgBox(Me.Page, "A quantidade a ser liquidada não pode ultrapassar a quantidade da CPR!")
            End If
        Else
            MsgBox(Me.Page, "Digite a quantidade a ser Liquidada!")
        End If
    End Sub

    Protected Sub grdLiquidacao_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GrdLiquidacao.RowDeleting
        SessionRecuperaCPR()
        If ObjCpr.IUD = "U" Then
            ObjCpr.Liquidacoes(e.RowIndex).IUD = "D"
            If Not ObjCpr.Liquidacoes(e.RowIndex).Salvar() Then
                ObjCpr.Liquidacoes(e.RowIndex).IUD = ""
                MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                Exit Sub
            End If
        End If
        ObjCpr.Liquidacoes.RemoveAt(e.RowIndex)

        GrdLiquidacao.DataSource = ObjCpr.Liquidacoes
        GrdLiquidacao.DataBind()

        SessionSalvaCPR()
    End Sub

    Protected Sub GrdLiquidacao_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GrdLiquidacao.RowDataBound
        If e.Row.RowType = DataControlRowType.Footer Then
            SessionRecuperaCPR()
            If Not IsNothing(ObjCpr.Liquidacoes) Then
                e.Row.Cells(1).Text = String.Format("{0}", ObjCpr.Liquidacoes.QuantidadeTotalLiquidada)
                e.Row.Cells(1).HorizontalAlign = HorizontalAlign.Right
            End If
        End If
    End Sub

    Protected Sub txtQtde_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtQtde.TextChanged
        If Not String.IsNullOrWhiteSpace(txtQtde.Text) And Not String.IsNullOrWhiteSpace(txtAreaCPR.Text) Then
            SessionRecuperaCPR()
            ObjCpr.Quantidade = txtQtde.Text
            txtProdutividade.Text = ObjCpr.ProdutividadeCalculada
            SessionSalvaCPR()
        End If
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CPR")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class