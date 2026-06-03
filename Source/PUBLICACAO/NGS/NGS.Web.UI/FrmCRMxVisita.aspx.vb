Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Reflection
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared


Public Class frmCRMxVisita
    Inherits BasePage
    Private ds As DataSet
    Private CRMListaConsulta As ListCRM
    Private ListParametrosVisitas As ListCRMxParametroVisita
    Private objParametroVisita As CRMxParametroVisita
    Private objVisita As Visita
    Private objVisitaMotivo As VisitaMotivo
    Private objVisitaAmeaca As VisitaAmeaca
    Private objVisitaProdutividade As VisitaProdutividade
    Private ListVisitas As ListVisita
    Private DataInicial As Date = Nothing
    Private DataInicioSafra As Date = Nothing
    Private DataVencimentoSafra As Date

#Region "Session"
    Private Sub SessaoSalvarListaCrmXParametroVisita()
        Session("ListaParametroVisita" + HID.Value.ToString) = ListParametrosVisitas
    End Sub

    Private Sub SessaoRecuperarListaCrmXParametroVisita()
        ListParametrosVisitas = Session("ListaParametroVisita" + HID.Value.ToString)
    End Sub


    Private Sub SessaoSalvarParametroXVisita()
        Session("ParametroVisita" + HID.Value.ToString) = objParametroVisita
    End Sub

    Private Sub SessaoRecuperarParametroXVisita()
        objParametroVisita = Session("ParametroVisita" + HID.Value.ToString)
    End Sub


    Private Sub SessaoSalvarVisita()
        Session("Visita" + HID.Value.ToString) = objVisita
    End Sub

    Private Sub SessaoRecuperarVisita()
        objVisita = Session("Visita" + HID.Value.ToString)
    End Sub

    Private Sub SessaoSalvarMotivoAmeacaProdutividade()
        Session("VisitaMotivo" + HID.Value.ToString) = objVisitaMotivo
        Session("VisitaAmeaca" + HID.Value.ToString) = objVisitaAmeaca
        Session("VisitaProdutividade" + HID.Value.ToString) = objVisitaProdutividade
    End Sub

    Private Sub SessaoRecuperarMotivoAmeacaProdutividade()
        objVisitaMotivo = Session("VisitaMotivo" + HID.Value.ToString)
        objVisitaAmeaca = Session("VisitaAmeaca" + HID.Value.ToString)
        objVisitaProdutividade = Session("VisitaProdutividade" + HID.Value.ToString)
    End Sub

    Private Sub SessaoSalvarListaVisita()
        Session("ListaVisita" + HID.Value.ToString) = ListVisitas
    End Sub

    Private Sub SessaoRecuperarListaVisita()
        If Session("ListaVisita" + HID.Value.ToString) Is Nothing Then
            ListVisitas = New ListVisita()
        Else
            ListVisitas = Session("ListaVisita" + HID.Value.ToString)
        End If
    End Sub

    Private Sub SessaoSalvaDataInicioVencimentoSafra()
        Session("DataInicioSafra" + HID.Value.ToString) = DataInicioSafra
        Session("DataVencimentoSafra" + HID.Value.ToString) = DataVencimentoSafra
    End Sub

    Private Sub SessaoRecuperaDataInicioVencimentoSafra()
        DataInicioSafra = Session("DataInicioSafra" + HID.Value.ToString)
        DataVencimentoSafra = Session("DataVencimentoSafra" + HID.Value.ToString)
    End Sub

#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("frmCRMxVisita", "ACESSAR") Then
                HID.Value = Guid.NewGuid.ToString
                ucConsultaClientes.SetarHID(HID.Value)
                'ddl.Carregar(ddlEmpresaVisita, CarregarDDL.Tabela.Empresas, "", True)
                'ddl.Carregar(ddlAnoVisita, CarregarDDL.Tabela.Ano, "2012;6;C", True)

                ddl.Carregar(ddlParametroAnoConsulta, CarregarDDL.Tabela.Ano, "2016;10;C", True)
                ddl.Carregar(ddlRepresentante, CarregarDDL.Tabela.ClientesXTipos, "6")
                ddl.Carregar(ddlMotivo, CarregarDDL.Tabela.Motivo, "")
                ddl.Carregar(ddlAmeaca, CarregarDDL.Tabela.ClientesXTipos, "17")
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")
                ddl.Carregar(ddlCultura, CarregarDDL.Tabela.Cultura, "")
                ddl.Carregar(ddlConsultaAnoCRM, CarregarDDL.Tabela.Ano, "2016;10;C", True)
                ddl.Carregar(ddlRegiao, CarregarDDL.Tabela.Regiao, "")
                ddl.Carregar(ddlMicroRegiao, CarregarDDL.Tabela.MicroRegiao, "")
                rbVisitaCRM_CheckedChanged(Me, New EventArgs())

                txtConsultaDataInicial.Text = Now.Date.AddMonths(-1)
                txtConsultaDataFinal.Text = Now.Date


                tcCRMxVisitas.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    '***************** ABA PARÂMETROS *************************************************************

    '*****************Escolha do Ano do CRM *******************************************************
    Protected Sub ddlAnoConsulta_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlParametroAnoConsulta.SelectedIndexChanged
        CRMListaConsulta = New ListCRM(IIf(String.IsNullOrWhiteSpace(ddlParametroAnoConsulta.SelectedValue), 0, ddlParametroAnoConsulta.SelectedValue))
        gridParametroConsulta.DataSource = (From p In CRMListaConsulta Where p.Consolidado = False Select p.Ano, p.CodigoEmpresa, p.EndEmpresa, p.NomeEmpresa)
        gridParametroConsulta.DataBind()

    End Sub

    '*****************Lista de CRM's **************************************************************
    Protected Sub gridParametroConsulta_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles gridParametroConsulta.SelectedIndexChanged
        'hidCodigoEmpresa.Value = gridParametroConsulta.SelectedRow.Cells(1).Text
        'hidEndEmpresa.Value = gridParametroConsulta.SelectedRow.Cells(2).Text
        'lblEmpresa.Text = gridParametroConsulta.SelectedRow.Cells(3).Text
        'lblAno.Text = ddlParametroAnoConsulta.SelectedValue

        lblIndiceEmpresa.Text = gridParametroConsulta.SelectedRow.Cells(3).Text
        lblIndiceAno.Text = ddlParametroAnoConsulta.SelectedValue

        'CarregaGridRepresentante()

        ListParametrosVisitas = New ListCRMxParametroVisita(ddlParametroAnoConsulta.SelectedValue)
        grdParametroRepresentante.DataSource = ListParametrosVisitas.Representantes
        grdParametroRepresentante.DataBind()

        SessaoSalvarListaCrmXParametroVisita()
    End Sub

    '*****************Lista dos Representantes para Clientes CRM **********************************
    Protected Sub grdParametroRepresentante_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles grdParametroRepresentante.SelectedIndexChanged
        SessaoRecuperarListaCrmXParametroVisita()
        grdParametroVisita.DataSource = ListParametrosVisitas.ClientesPorRepresentante(grdParametroRepresentante.SelectedRow.Cells(1).Text)
        grdParametroVisita.DataBind()
    End Sub

    Protected Sub grdParametroCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles grdParametroVisita.SelectedIndexChanged
        'SessaoRecuperarListaVisita()
        'SessaoRecuperarListaCrmXParametroVisita()

        'CarregaListaDeVisitasNoGrid()

        'Dim cli As String() = grdParametroCliente.SelectedRow.Cells(1).Text.Split("-")
        'objParametroVisita = ListParametrosVisitas.Find(Function(s) s.CodigoCliente = cli(0) And s.EndCliente = cli(1))

        'If rbClientesCRM.Checked Then
        '    ListVisitas = ListParametrosVisitas.Find(Function(s) s.CodigoCliente = cli(0) And s.EndCliente = cli(1)).Visitas
        'Else
        '    ListVisitas.ListVisitasRealizadasClienteAvulso(cli(0), cli(1))
        'End If

        'grdVisualizaVisita.DataSource = ListVisitas.ToArray
        'grdVisualizaVisita.DataBind()

        'Dim NomeCliente As String = grdConsultaCliente.SelectedRow.Cells(2).Text
        'Dim Fazenda As String = grdConsultaCliente.SelectedRow.Cells(3).Text
        'Dim CodigoRepresentante As String = grdConsultaRepresentante.SelectedRow.Cells(1).Text 'ListVisitas.Find(Function(s) s.CodigoCliente = cli(0) And s.EndCliente = cli(1)).CodigoRepresentante
        'Dim EndRepresentante As Integer = grdConsultaRepresentante.SelectedRow.Cells(2).Text 'ListVisitas.Find(Function(s) s.CodigoCliente = cli(0) And s.EndCliente = cli(1)).EndRepresentante
        'lblCliente.Text = NomeCliente.Substring(0, IIf(Len(NomeCliente) > 31, 31, Len(NomeCliente)))
        'lblFazenda.Text = Fazenda.Substring(0, IIf(Len(Fazenda) > 33, 33, Len(Fazenda)))

        'If CodigoRepresentante <> "NENHUM" Then
        '    ddlRepresentante.SelectedValue = CodigoRepresentante & "-" & EndRepresentante
        'End If

        ''SessaoSalvarListaVisita()
        ''SessaoSalvarParametroXVisita()
        'tcVisitas.ActiveTabIndex = 1
    End Sub

    '*****************Gravação dos Parâmetros de Visitas para Clientes CRM *************************
    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("CRMxVisita", "GRAVAR") Then
            SessaoRecuperarListaCrmXParametroVisita()
            If ListParametrosVisitas.Salvar() Then
                MsgBox(Me.Page, "Registros gravados com Sucesso.", eTitulo.Sucess)
            End If
            SessaoSalvarListaCrmXParametroVisita()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para Gravar Parâmetro.")
        End If

    End Sub

    Protected Sub txtNumVisita_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        SessaoRecuperarListaCrmXParametroVisita()
        Dim txt As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
        Dim cli As String() = grdParametroVisita.Rows(row.RowIndex).Cells(1).Text.Split("-")
        Dim c As CRMxParametroVisita
        c = ListParametrosVisitas.Find(Function(s) s.CodigoCliente = cli(0) And s.EndCliente = cli(1))
        c.MinimoVisita = txt.Text
        c.IUD = "U"
        SessaoSalvarListaCrmXParametroVisita()
    End Sub
    '***********************************************************************************************


    '***************** ABA CONSULTA ****************************************************************

    Protected Sub rbVisitaCRM_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rbVisitaCRM.CheckedChanged
        LimpaVisitas()
        pnlConsultaCRM.Visible = True
        pnlConsultaAvulsaTodas.Visible = False
        pnlRepresentante.Visible = True
        pnlClientesVisitas.Width = "690"
        lblCliente.Text = ""
        lblFazenda.Text = ""
    End Sub

    Protected Sub rbVisitaAvulsa_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rbVisitaAvulsa.CheckedChanged
        LimpaVisitas()
        pnlConsultaCRM.Visible = False
        pnlConsultaAvulsaTodas.Visible = True
        pnlRepresentante.Visible = False
        pnlClientesVisitas.Width = "920"
        lblCliente.Text = ""
        lblFazenda.Text = ""

    End Sub

    Protected Sub rbTodasVisitas_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rbTodasVisitas.CheckedChanged
        LimpaVisitas()
        pnlConsultaCRM.Visible = False
        pnlConsultaAvulsaTodas.Visible = True
        pnlRepresentante.Visible = False
        pnlClientesVisitas.Width = "920"
        lblCliente.Text = ""
        lblFazenda.Text = ""
    End Sub

    Private Sub LimpaVisitas()
        grdConsultaRepresentante.DataSource = Nothing
        grdConsultaRepresentante.DataBind()

        grdConsultaCliente.DataSource = Nothing
        grdConsultaCliente.DataBind()

        grdVisualizaVisita.DataSource = Nothing
        grdVisualizaVisita.DataBind()

    End Sub

    Protected Sub ddlConsultaAnoCRM_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlConsultaAnoCRM.SelectedIndexChanged
        CRMListaConsulta = New ListCRM(IIf(String.IsNullOrWhiteSpace(ddlConsultaAnoCRM.SelectedValue), 0, ddlConsultaAnoCRM.SelectedValue))
        grdConsultaCRM.DataSource = (From p In CRMListaConsulta Where p.Consolidado = False Select p.Ano, p.CodigoEmpresa, p.EndEmpresa, p.NomeEmpresa)
        grdConsultaCRM.DataBind()

        Dim strSQL As String = " SELECT MIN(InicioDeSafra) InicioDeSafra, MAX(Vencimento) Vencimento  " & _
                               "   FROM Safras " & _
                               "  WHERE YEAR(Vencimento) = " & ddlConsultaAnoCRM.SelectedValue & ""
        Dim dsSafras As DataSet = Banco.ConsultaDataSet(strSQL, "Safras")
        DataInicioSafra = dsSafras.Tables(0).Rows(0)("InicioDeSafra")
        DataVencimentoSafra = dsSafras.Tables(0).Rows(0)("Vencimento")

        SessaoSalvaDataInicioVencimentoSafra()

    End Sub

    Protected Sub grdConsultaCRM_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles grdConsultaCRM.SelectedIndexChanged
        'ListParametrosVisitas = New ListCRMxParametroVisita(ddlConsultaAnoCRM.SelectedValue)
        'grdConsultaRepresentante.DataSource = ListParametrosVisitas.Representantes
        'grdConsultaRepresentante.DataBind()
        'SessaoSalvarListaCrmXParametroVisita()

        'If String.IsNullOrWhiteSpace(txtConsultaDataInicial.Text) Or String.IsNullOrWhiteSpace(txtConsultaDataFinal.Text) Or Not IsDate(txtConsultaDataInicial.Text) Or Not IsDate(txtConsultaDataFinal.Text) Then
        '    MsgBox(Me.Page, "Indique a Data Inicial e Final para a Consulta!")
        'Else

        'hidCodigoEmpresa.Value = grdConsultaCRM.SelectedRow.Cells(1).Text
        'hidEndEmpresa.Value = grdConsultaCRM.SelectedRow.Cells(2).Text
        'lblEmpresa.Text = grdConsultaCRM.SelectedRow.Cells(3).Text

        'ddlEmpresaVisita.SelectedValue = grdConsultaCRM.SelectedRow.Cells(1).Text & "-" & grdConsultaCRM.SelectedRow.Cells(2).Text
        'ddlEmpresaVisita.Enabled = False

        'lblAno.Text = ddlConsultaAnoCRM.SelectedValue

        'ddlAnoVisita.SelectedValue = ddlConsultaAnoCRM.SelectedValue
        'ddlAnoVisita.Enabled = False

        If rbVisitaCRM.Checked Then
            Dim CodigoEmpresa As String = grdConsultaCRM.SelectedRow.Cells(1).Text
            Dim EndEmpresa As Integer = grdConsultaCRM.SelectedRow.Cells(2).Text
            Dim objCRM As [Lib].Negocio.CRM = New [Lib].Negocio.CRM(CodigoEmpresa, EndEmpresa, ddlConsultaAnoCRM.SelectedValue, False)

            ListVisitas = New ListVisita(objCRM)

            grdConsultaRepresentante.DataSource = ListVisitas.ListRepresentantes '(From c In ListVisitas Select c.CodigoRepresentante, c.EndRepresentante, c.NomeRepresentante).Distinct
            grdConsultaRepresentante.DataBind()
            'grdConsultaCliente.DataSource = ListVisitas.ListClientesVisitadosAvulsamente()
            'grdConsultaCliente.DataBind()
            tcVisitas.ActiveTabIndex = 0
            SessaoSalvarListaVisita()
            'End If
        End If
    End Sub

    Protected Sub ddlRegiao_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlRegiao.SelectedIndexChanged
        If Not String.IsNullOrWhiteSpace(ddlRegiao.SelectedValue) Then
            ddl.Carregar(ddlMicroRegiao, CarregarDDL.Tabela.MicroRegiao, " Regiao_Id= " & ddlRegiao.SelectedValue)
        Else
            ddl.Carregar(ddlMicroRegiao, CarregarDDL.Tabela.MicroRegiao, "")
        End If
        'CarregaGridRepresentante()
    End Sub

    Protected Sub lnkConsultaVisitas_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultaVisitas.Click

        If Not rbVisitaCRM.Checked Then

            If String.IsNullOrWhiteSpace(txtConsultaDataInicial.Text) Or String.IsNullOrWhiteSpace(txtConsultaDataFinal.Text) Or Not IsDate(txtConsultaDataInicial.Text) Or Not IsDate(txtConsultaDataFinal.Text) Then
                MsgBox(Me.Page, "Indique a Data Inicial e Final para a Consulta!")
            Else
                ListVisitas = New ListVisita(txtConsultaDataInicial.Text, txtConsultaDataFinal.Text, ddlRegiao.SelectedValue, ddlMicroRegiao.SelectedIndex)
                grdConsultaCliente.DataSource = ListVisitas.ListClientes("", IIf(rbVisitaAvulsa.Checked, "N", ""))

                grdConsultaCliente.DataBind()
                tcVisitas.ActiveTabIndex = 0
                SessaoSalvarListaVisita()
            End If

        End If

    End Sub

    Protected Sub grdConsultaRepresentante_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles grdConsultaRepresentante.SelectedIndexChanged
        SessaoRecuperarListaCrmXParametroVisita()
        'grdConsultaCliente.DataSource = ListParametrosVisitas.ClientesPorRepresentante(grdConsultaRepresentante.SelectedRow.Cells(1).Text)
        'grdConsultaCliente.DataBind()

        SessaoRecuperarListaVisita()
        grdConsultaCliente.DataSource = ListVisitas.ListClientes(grdConsultaRepresentante.SelectedRow.Cells(1).Text)
        grdConsultaCliente.DataBind()
        tcVisitas.ActiveTabIndex = 0
    End Sub

    Protected Sub grdConsultaCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles grdConsultaCliente.SelectedIndexChanged
        SessaoRecuperarListaVisita()
        SessaoRecuperarListaCrmXParametroVisita()
        Dim cli As String() = grdConsultaCliente.SelectedRow.Cells(1).Text.Split("-")


        Dim NomeCliente As String = grdConsultaCliente.SelectedRow.Cells(2).Text
        Dim Fazenda As String = grdConsultaCliente.SelectedRow.Cells(3).Text

        'Dim CodigoRepresentante As String = grdConsultaRepresentante.SelectedRow.Cells(1).Text
        'Dim EndRepresentante As Integer = grdConsultaRepresentante.SelectedRow.Cells(2).Text

        lblCliente.Text = NomeCliente.Substring(0, IIf(Len(NomeCliente) > 31, 31, Len(NomeCliente)))
        lblFazenda.Text = Fazenda.Substring(0, IIf(Len(Fazenda) > 33, 33, Len(Fazenda)))

        'If Not rbVisitaCRM.Checked Then
        'grdVisualizaVisita.DataSource = ListVisitas.ListVisitasRealizadasClienteAvulso(cli(0), cli(1))
        grdVisualizaVisita.DataSource = ListVisitas.ListVisitasRealizadas(cli(0), cli(1))
        grdVisualizaVisita.DataBind()
        'Else
        'grdVisualizaVisita.DataSource = ListParametrosVisitas.Find(Function(s) s.CodigoCliente = cli(0) And s.EndCliente = cli(1)).Visitas
        'grdVisualizaVisita.DataBind()
        'End If

        tcVisitas.ActiveTabIndex = 1
    End Sub


    '***************** MANUTENÇÃO DA VISITA ********************************************************
    Protected Sub lnkNovoCadVisita_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovoCadVisita.Click
        If Funcoes.VerificaPermissao("CRMxVisita", "GRAVAR") Then

            SessaoRecuperarParametroXVisita()
            SessaoRecuperarVisita()
            SessaoRecuperarListaCrmXParametroVisita()
            SessaoRecuperarListaVisita()
            SessaoRecuperaDataInicioVencimentoSafra()

            If ValidaCamposVisita() Then

                'If (IsDate(DataInicioSafra) And IsDate(DataVencimentoSafra)) AndAlso (CType(txtData.Text, Date) >= DataInicioSafra And CType(txtData.Text, Date) <= DataVencimentoSafra) Then

                objParametroVisita = New CRMxParametroVisita(ListParametrosVisitas)

                objVisita = New Visita(objParametroVisita)
                objVisita.IUD = "I"
                'objVisita.CodigoEmpresa = ddlEmpresaVisita.SelectedValue.Split("-")(0)
                'objVisita.EndEmpresa = ddlEmpresaVisita.SelectedValue.Split("-")(1)
                'objVisita.Ano = ddlAnoVisita.SelectedValue
                objVisita.Data = CDate(txtData.Text)
                'objVisita.Ano = objVisita.Data.Year

                'objVisita.Consolidado = 0
                objVisita.CodigoCliente = txtCodigoCliente.Value
                objVisita.EndCliente = txtEndCliente.Value

                Dim Rep As String() = ddlRepresentante.SelectedValue.Split("-")
                objVisita.CodigoRepresentante = Rep(0)
                objVisita.EndRepresentante = Rep(1)


                If Not String.IsNullOrWhiteSpace(txtKmInicial.Text) Then
                    objVisita.KmInicial = txtKmInicial.Text
                End If

                If Not String.IsNullOrWhiteSpace(txtKmFinal.Text) Then
                    objVisita.KmFinal = txtKmFinal.Text
                End If

                If objVisita.IUD = "I" Then
                    ListVisitas.Add(objVisita)
                End If

                If objVisita.Salvar Then
                    MsgBox(Me.Page, "Visita Gravada com Sucesso.", eTitulo.Sucess)
                    SessaoSalvarVisita()
                    SessaoSalvarListaVisita()
                Else
                    MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            End If

        Else
            MsgBox(Me.Page, "Usuário sem permissão para Gravar Visita.")
        End If
    End Sub

    Private Function ValidaCamposVisita() As Boolean
        Dim retorno As Boolean = True
        If ddlRepresentante.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Indique o Representante!")
            ddlRepresentante.Focus()
            retorno = False
        ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            MsgBox(Me.Page, "Indique Cliente!")
            btnCliente.Focus()
            retorno = False
        ElseIf String.IsNullOrWhiteSpace(txtData.Text) Then
            MsgBox(Me.Page, "Indique a Data da Visita!")
            txtData.Focus()
            retorno = False
        ElseIf (CDate(txtData.Text).Date < Now.Date.AddDays(-15)) Then
            MsgBox(Me.Page, "A data da Visita não pode ser menor do que a data atual menos 15 dias!")
            txtData.Focus()
            retorno = False
        End If
        Return retorno
    End Function

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        SessaoRecuperarParametroXVisita()
        SessaoRecuperarVisita()
        SessaoRecuperarListaCrmXParametroVisita()
        SessaoRecuperarListaVisita()

        'objParametroVisita = New CRMxParametroVisita(ListParametrosVisitas)

        objVisita.IUD = "U"

        'objVisita.CodigoEmpresa = hidCodigoEmpresa.Value
        'objVisita.EndEmpresa = hidEndEmpresa.Value
        'objVisita.Ano = lblAno.Text

        'objVisita.CodigoEmpresa = ddlEmpresaVisita.SelectedValue.Split("-")(0)
        'objVisita.EndEmpresa = ddlEmpresaVisita.SelectedValue.Split("-")(1)
        'objVisita.Ano = ddlAnoVisita.SelectedValue

        'objVisita.Consolidado = 0

        objVisita.CodigoCliente = txtCodigoCliente.Value
        objVisita.EndCliente = txtEndCliente.Value

        Dim Rep As String() = ddlRepresentante.SelectedValue.Split("-")
        objVisita.CodigoRepresentante = Rep(0)
        objVisita.EndRepresentante = Rep(1)

        objVisita.Data = CDate(txtData.Text)
        objVisita.KmInicial = txtKmInicial.Text
        objVisita.KmFinal = txtKmFinal.Text

        If objVisita.Salvar Then
            MsgBox(Me.Page, "Visita Alterada com Sucesso.", eTitulo.Sucess)
        End If

        SessaoSalvarVisita()
        SessaoSalvarListaVisita()
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClientexVisita" & HID.Value) Is Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objClientexVisita" & HID.Value.ToString), [Lib].Negocio.Cliente)
            txtNomeDoCliente.Text = pCliente.Nome
            txtCodigoCliente.Value = pCliente.Codigo
            txtEndCliente.Value = pCliente.CodigoEndereco
            txtFazendaCliente.Text = pCliente.Complemento
            Session.Remove("objClientexVisita" & HID.Value.ToString)
        End If

    End Sub


    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCliente.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClientexVisita" & HID.Value.ToString, "txtNome")
    End Sub

    Protected Sub lnkLimparCadVisita_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparCadVisita.Click
        LimparCamposVisita()
    End Sub

    Private Sub LimparCamposVisita()
        'ddlEmpresaVisita.Enabled = True
        'ddlAnoVisita.Enabled = True

        txtCodigoCliente.Value = ""
        txtNomeDoCliente.Text = ""
        txtEndCliente.Value = ""
        txtFazendaCliente.Text = ""

        txtData.Text = ""
        txtKmInicial.Text = ""
        txtKmFinal.Text = ""

        LimparVisitaMotivo()
        LimparVisitaAmeaca()
        limparVisitaProdutividade()

        SessaoRecuperarVisita()

        If Not objVisita Is Nothing Then
            objVisita.ListVisitaMotivo.Clear()
            objVisita.ListVisitaAmeaca.Clear()
            objVisita.ListVisitaProdutividade.Clear()
            objVisita = Nothing
        End If

        'CarregaMotivoNoGrid()
        'CarregaAmeacaNoGrid()
        'CarregaProdutividadeNoGrid()
        grdVisitaMotivos.DataSource = Nothing
        grdVisitaMotivos.DataBind()
        grdVisitaAmeaca.DataSource = Nothing
        grdVisitaAmeaca.DataBind()
        grdVisitaProdutividade.DataSource = Nothing
        grdVisitaProdutividade.DataBind()


        SessaoSalvarVisita()
    End Sub

    Protected Sub ImdAlterar_Click(ByVal sender As Object, ByVal e As EventArgs)
        If Funcoes.VerificaPermissao("CRMxVisita", "ALTERAR") Then

            SessaoRecuperarParametroXVisita()
            SessaoRecuperarListaVisita()
            'SessaoRecuperarVisita()

            Dim lnkAlterar As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnkAlterar.NamingContainer, GridViewRow)
            Dim NumVisita As Integer = grdVisualizaVisita.Rows(row.RowIndex).Cells(0).Text

            Dim cli As String() = grdConsultaCliente.SelectedRow.Cells(1).Text.Split("-")
            objVisita = ListVisitas.Find(Function(S) S.CodigoCliente = cli(0) And S.EndCliente = cli(1) And S.NumeroVisita = NumVisita)

            objVisita.IUD = "U"
            ddlRepresentante.SelectedValue = objVisita.CodigoRepresentante & "-" & objVisita.EndRepresentante
            txtCodigoCliente.Value = objVisita.CodigoCliente
            txtNomeDoCliente.Text = objVisita.Cliente.Nome & " - " & objVisita.EndCliente
            txtEndCliente.Value = objVisita.EndCliente
            txtFazendaCliente.Text = objVisita.Cliente.Complemento
            txtData.Text = objVisita.Data
            txtKmInicial.Text = objVisita.KmInicial
            txtKmFinal.Text = objVisita.KmFinal

            SessaoSalvarVisita()

            CarregaMotivoNoGrid()
            CarregaAmeacaNoGrid()
            CarregaProdutividadeNoGrid()

            tcCRMxVisitas.ActiveTabIndex = 2
            tcOutros.ActiveTabIndex = 0

        Else
            MsgBox(Me.Page, "Usuário sem permissão para Alterar Visita.")
        End If

    End Sub

    Protected Sub ImdVisualizar_Click(ByVal sender As Object, ByVal e As EventArgs)

        'If Funcoes.VerificaPermissao("CRMxVisita", "RELATORIO") Then

        '    Using rpt As New ReportDocument()
        '        Try
        '            'Sql = "  SELECT Estado_Id as UF, Regiao, Descricao" & _
        '            '    " FROM Estados " & _
        '            '    " ORDER BY Estado_Id"

        '            'ds = Banco.ConsultaDataSet(Sql, "Estados")


        '            'DS = CRMxVisitas.ToArray

        '            rpt.FileName = Server.MapPath("~/Reports/Cr_CRMVisualizaVisita.rpt")
        '            rpt.Load(rpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

        '            Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
        '            Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
        '            Dim arquivo As String = NomeArquivo

        '            rpt.SetDataSource(ds)


        '            Dim parameters = New Dictionary(Of String, Object)()
        '            'parameters.Add("ObsVisivel", ObsVisivel)

        '            BindReport(rpt, parameters)

        '            If Dir(NomeArquivo).Length > 0 Then
        '                Kill(NomeArquivo)
        '            End If



        '            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        '            rpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
        '            ScriptManager.RegisterClientScriptBlock(Me, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo2 & "');", True)
        '        Catch ex As Exception
        '            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        '        Finally
        '            rpt.Close()
        '            rpt.Dispose()
        '        End Try
        '    End Using
        'Else
        '    MsgBox(Me.Page, "Usuario sem permissao para tirar Relatório")
        'End If



    End Sub

    '*****************Manutenção do Motivo da Visita***********************************************
    Protected Sub lnkGravarMotivo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGravarMotivo.Click
        If ddlMotivo.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Indique um Motivo!")
            ddlMotivo.Focus()
        ElseIf String.IsNullOrWhiteSpace(txtObsMotivo.Text) Then
            MsgBox(Me.Page, "Preencha o campo Observação!")
            txtObsMotivo.Focus()
        Else
            SessaoRecuperarVisita()
            SessaoRecuperarMotivoAmeacaProdutividade()
            If objVisitaMotivo Is Nothing Then
                objVisitaMotivo = New VisitaMotivo(objVisita.NumeroVisita)
                objVisitaMotivo.IUD = "I"
            End If

            objVisitaMotivo.NumeroVisita = objVisita.NumeroVisita
            objVisitaMotivo.CodigoMotivo = ddlMotivo.SelectedValue
            objVisitaMotivo.Observacao = txtObsMotivo.Text

            If objVisitaMotivo.IUD = "I" Then
                objVisita.ListVisitaMotivo.Add(objVisitaMotivo)
            End If

            If objVisitaMotivo.Salvar() Then
                SessaoSalvarMotivoAmeacaProdutividade()
                SessaoSalvarVisita()
                CarregaMotivoNoGrid()
                LimparVisitaMotivo()
            Else
                objVisita.ListVisitaMotivo.Remove(objVisitaMotivo)
                MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
            End If
        End If
    End Sub

    Public Sub CarregaMotivoNoGrid()
        SessaoRecuperarVisita()
        grdVisitaMotivos.DataSource = objVisita.ListVisitaMotivo
        grdVisitaMotivos.DataBind()
    End Sub

    Protected Sub grdVisitaMotivos_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles grdVisitaMotivos.SelectedIndexChanged
        SessaoRecuperarVisita()
        Dim codMotivo As Integer = grdVisitaMotivos.SelectedRow.Cells(1).Text
        objVisitaMotivo = objVisita.ListVisitaMotivo.Find(Function(S) S.NumeroVisita = objVisita.NumeroVisita.ToString And S.CodigoMotivo = codMotivo)
        objVisitaMotivo.IUD = "U"
        ddlMotivo.SelectedValue = objVisitaMotivo.CodigoMotivo
        txtObsMotivo.Text = objVisitaMotivo.Observacao
        SessaoSalvarMotivoAmeacaProdutividade()
    End Sub

    Protected Sub lnkExcluirMotivo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluirMotivo.Click
        If grdVisitaMotivos.SelectedIndex > -1 Then
            SessaoRecuperarMotivoAmeacaProdutividade()
            SessaoRecuperarVisita()
            objVisitaMotivo.IUD = "D"
            objVisitaMotivo.Salvar()
            objVisita.ListVisitaMotivo.Remove(objVisitaMotivo)
            CarregaMotivoNoGrid()
            LimparVisitaMotivo()
        Else
            MsgBox(Me.Page, "Selecione um registro para Exluir!")
        End If
    End Sub

    Public Sub LimparVisitaMotivo()
        ddlMotivo.SelectedValue = ""
        txtObsMotivo.Text = ""
        objVisitaMotivo = Nothing
        SessaoSalvarMotivoAmeacaProdutividade()
    End Sub

    '*****************Manutenção da Ameaça que o visitado sofreu **********************************
    Protected Sub lnkGravarAmeaca_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGravarAmeaca.Click
        If ddlAmeaca.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Indique uma Ameaça!")
            ddlAmeaca.Focus()
        ElseIf String.IsNullOrWhiteSpace(txtObsAmeaca.Text) Then
            MsgBox(Me.Page, "Preencha o campo Observação!")
            txtObsAmeaca.Focus()
        Else
            SessaoRecuperarVisita()
            SessaoRecuperarMotivoAmeacaProdutividade()
            If objVisitaAmeaca Is Nothing Then
                objVisitaAmeaca = New VisitaAmeaca(objVisita.NumeroVisita)
                objVisitaAmeaca.IUD = "I"
            End If

            Dim Ameaca As String() = ddlAmeaca.SelectedValue.Split("-")
            objVisitaAmeaca.CodigoAmeaca = Ameaca(0)
            objVisitaAmeaca.EndAmeaca = Ameaca(1)
            objVisitaAmeaca.Observacao = txtObsAmeaca.Text

            If objVisitaAmeaca.IUD = "I" Then
                objVisita.ListVisitaAmeaca.Add(objVisitaAmeaca)
            End If

            If objVisitaAmeaca.Salvar() Then
                SessaoSalvarMotivoAmeacaProdutividade()
                SessaoSalvarVisita()
                CarregaAmeacaNoGrid()
                LimparVisitaAmeaca()
            Else
                objVisita.ListVisitaAmeaca.Remove(objVisitaAmeaca)
                MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
            End If
        End If
    End Sub

    Public Sub CarregaAmeacaNoGrid()
        SessaoRecuperarVisita()
        grdVisitaAmeaca.DataSource = objVisita.ListVisitaAmeaca.ToArray
        grdVisitaAmeaca.DataBind()
    End Sub

    Protected Sub grdVisitaAmeaca_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles grdVisitaAmeaca.SelectedIndexChanged
        SessaoRecuperarVisita()
        SessaoRecuperarMotivoAmeacaProdutividade()
        Dim CodigoAmeaca As String = grdVisitaAmeaca.SelectedRow.Cells(1).Text
        Dim EndAmeaca As Integer = grdVisitaAmeaca.SelectedRow.Cells(2).Text
        objVisitaAmeaca = objVisita.ListVisitaAmeaca.Find(Function(S) S.NumeroVisita = objVisita.NumeroVisita.ToString And S.CodigoAmeaca = CodigoAmeaca And S.EndAmeaca = EndAmeaca)
        objVisitaAmeaca.IUD = "U"
        ddlAmeaca.SelectedValue = objVisitaAmeaca.CodigoAmeaca & "-" & objVisitaAmeaca.EndAmeaca
        txtObsAmeaca.Text = objVisitaAmeaca.Observacao
        SessaoSalvarMotivoAmeacaProdutividade()
    End Sub

    Protected Sub lnkExcluirAmeaca_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluirAmeaca.Click
        If grdVisitaAmeaca.SelectedIndex > -1 Then
            SessaoRecuperarMotivoAmeacaProdutividade()
            SessaoRecuperarVisita()
            objVisitaAmeaca.IUD = "D"
            objVisitaAmeaca.Salvar()
            objVisita.ListVisitaAmeaca.Remove(objVisitaAmeaca)
            objVisitaAmeaca = Nothing
            SessaoSalvarMotivoAmeacaProdutividade()
            CarregaAmeacaNoGrid()
            LimparVisitaAmeaca()
        Else
            MsgBox(Me.Page, "Selecione um Registro para Excluir!")
        End If

    End Sub

    Public Sub LimparVisitaAmeaca()
        ddlAmeaca.SelectedValue = ""
        txtObsAmeaca.Text = ""
        objVisitaAmeaca = Nothing
        SessaoSalvarMotivoAmeacaProdutividade()
    End Sub

    '*****************Manutenção da Produtividade do Visitado***************************************
    Protected Sub lnkGravarProdutividade_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkGravarProdutividade.Click
        If ddlSafra.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Indique a Safra!")
            ddlSafra.Focus()
        ElseIf ddlCultura.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Indique a Cultura!")
            ddlCultura.Focus()
        ElseIf CDec(TxtProdutividade.Text) = 0 Then
            MsgBox(Me.Page, "Indique a Produtividade!")
            TxtProdutividade.Focus()
        Else
            SessaoRecuperarVisita()
            SessaoRecuperarMotivoAmeacaProdutividade()

            If objVisitaProdutividade Is Nothing Then
                objVisitaProdutividade = New VisitaProdutividade(objVisita.NumeroVisita)
                objVisitaProdutividade.IUD = "I"
            End If

            objVisitaProdutividade.CodigoSafra = ddlSafra.SelectedValue
            objVisitaProdutividade.CodigoCultura = ddlCultura.SelectedValue
            objVisitaProdutividade.Produtividade = TxtProdutividade.Text


            If objVisitaProdutividade.IUD = "I" Then
                objVisita.ListVisitaProdutividade.Add(objVisitaProdutividade)
            End If

            If objVisitaProdutividade.Salvar() Then
                SessaoSalvarVisita()
                SessaoSalvarMotivoAmeacaProdutividade()
                CarregaProdutividadeNoGrid()
                limparVisitaProdutividade()
            Else
                objVisita.ListVisitaProdutividade.Remove(objVisitaProdutividade)
                MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
            End If
        End If
    End Sub

    Public Sub CarregaProdutividadeNoGrid()
        SessaoRecuperarVisita()
        grdVisitaProdutividade.DataSource = objVisita.ListVisitaProdutividade
        grdVisitaProdutividade.DataBind()
    End Sub

    Protected Sub grdVisitaProdutividade_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles grdVisitaProdutividade.SelectedIndexChanged
        SessaoRecuperarVisita()
        SessaoRecuperarMotivoAmeacaProdutividade()
        Dim CodigoSafra As String = grdVisitaProdutividade.SelectedRow.Cells(1).Text
        Dim CodigoCultura As Integer = grdVisitaProdutividade.SelectedRow.Cells(2).Text
        objVisitaProdutividade = objVisita.ListVisitaProdutividade.Find(Function(S) S.NumeroVisita = objVisita.NumeroVisita.ToString And S.CodigoSafra = CodigoSafra And S.CodigoCultura = CodigoCultura)
        objVisitaProdutividade.IUD = "U"
        ddlSafra.SelectedValue = objVisitaProdutividade.CodigoSafra
        ddlCultura.SelectedValue = objVisitaProdutividade.CodigoCultura
        TxtProdutividade.Text = objVisitaProdutividade.Produtividade
        SessaoSalvarMotivoAmeacaProdutividade()
    End Sub

    Protected Sub lnkExcluirProdutividade_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluirProdutividade.Click
        If grdVisitaProdutividade.SelectedIndex > -1 Then
            SessaoRecuperarMotivoAmeacaProdutividade()
            SessaoRecuperarVisita()
            objVisitaProdutividade.IUD = "D"
            objVisitaProdutividade.Salvar()
            objVisita.ListVisitaProdutividade.Remove(objVisitaProdutividade)
            objVisitaProdutividade = Nothing
            SessaoSalvarMotivoAmeacaProdutividade()
            CarregaProdutividadeNoGrid()
            limparVisitaProdutividade()
        Else
            MsgBox(Me.Page, "Selecione um Registro para Excluir!")
        End If
    End Sub

    Public Sub limparVisitaProdutividade()
        ddlSafra.SelectedValue = ""
        ddlCultura.SelectedValue = ""
        TxtProdutividade.Text = ""
        objVisitaProdutividade = Nothing
        SessaoSalvarMotivoAmeacaProdutividade()
    End Sub

    '************************************************************************************************

    '*****************Relatórios**********************************************************************
    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        'If Funcoes.VerificaPermissao("CRMxVisita", "RELATORIO") Then

        '    SessaoSalvarListaCrmXParametroVisita()

        '    Using crpt As New ReportDocument()
        '        Try
        '            'Sql = "  SELECT Estado_Id as UF, Regiao, Descricao" & _
        '            '    " FROM Estados " & _
        '            '    " ORDER BY Estado_Id"

        '            'ds = Banco.ConsultaDataSet(Sql, "Estados")


        '            'DS = CRMxVisitas.ToArray

        '            crpt.FileName = Server.MapPath("~/Reports/Cr_CRMxVisitas.rpt")
        '            crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

        '            Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
        '            Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
        '            Dim arquivo As String = NomeArquivo

        '            crpt.SetDataSource(ds)

        '            Dim crparametervalues As ParameterValues
        '            Dim crparameterdiscretevalue As ParameterDiscreteValue
        '            Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
        '            Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

        '            crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

        '            crparameterfielddefinition = crparameterfielddefinitions.Item("Nome")
        '            crparametervalues = crparameterfielddefinition.CurrentValues
        '            crparameterdiscretevalue = New ParameterDiscreteValue
        '            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssNomeEmpresa")
        '            crparametervalues.Add(crparameterdiscretevalue)
        '            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

        '            crparameterfielddefinition = crparameterfielddefinitions.Item("Cidade")
        '            crparametervalues = crparameterfielddefinition.CurrentValues
        '            crparameterdiscretevalue = New ParameterDiscreteValue
        '            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa")
        '            crparametervalues.Add(crparameterdiscretevalue)
        '            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

        '            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        '            crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
        '            ScriptManager.RegisterClientScriptBlock(Me, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo2 & "');", True)
        '        Catch ex As Exception
        '            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        '        Finally
        '            crpt.Close()
        '            crpt.Dispose()
        '        End Try
        '    End Using
        'Else
        '    MsgBox(Me.Page, "Usuario sem permissao para tirar Relatório")
        'End If
    End Sub

    Protected Sub lnkConsulta_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsulta.Click
        CarregaGerenciamentoDeVisitas()
    End Sub

    Private Sub CarregaGerenciamentoDeVisitas()
        SessaoRecuperarParametroXVisita()
        SessaoRecuperarListaVisita()

        If gridParametroConsulta.SelectedIndex < 0 Then
            MsgBox(Me.Page, "Selecione um CRM para Consultar!")
        Else
            Dim CodEmpresa As String = gridParametroConsulta.SelectedRow.Cells(1).Text
            Dim EndEmpresa As Integer = gridParametroConsulta.SelectedRow.Cells(2).Text
            Dim ano As Integer = ddlParametroAnoConsulta.SelectedValue

            Dim sql As New StringBuilder
            sql.Append(
                "select MIN(InicioDeSafra) as datasafra, " & vbCrLf & _
                "       year(MIN(InicioDeSafra)) as ano," & vbCrLf & _
                "       month(MIN(InicioDeSafra)) as mes," & vbCrLf & _
                "       case " & vbCrLf & _
                "         when DATEDIFF(MONTH,MIN(InicioDeSafra),CONVERT(date,GETDATE())) >12" & vbCrLf & _
                "           then  12" & vbCrLf & _
                "           else" & vbCrLf & _
                "                 DateDiff(Month, MIN(InicioDeSafra), GETDATE())" & vbCrLf & _
                "       END as MesAtual" & vbCrLf & _
                "  into #Temp" & vbCrLf & _
                "  from safras " & vbCrLf & _
                " where YEAR(vencimento) = " & ano & vbCrLf & _
                " /*                                         */ " & vbCrLf & _
                "select V.Cliente as CodigoCliente, " & vbCrLf & _
                "       V.EndCliente," & vbCrLf & _
                "       C.Nome as NomeCliente," & vbCrLf & _
                "       t.datasafra as InicioCRM, " & vbCrLf & _
                "       t.MesAtual," & vbCrLf & _
                "       cxpv.MinimoVisitas," & vbCrLf & _
                "       SUM(case when Year(data) = t.ano and MONTH(DATA) = t.mes  then 1  else 0 end) Mes01," & vbCrLf & _
                " SUM(case                                                   " & vbCrLf & _
                "      when Year(data)  = case                               " & vbCrLf & _
                "                           when t.mes + 1 > 12              " & vbCrLf & _
                "                             then t.ano + 1                 " & vbCrLf & _
                "                             else t.ano                     " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "       and MONTH(DATA) = case                               " & vbCrLf & _
                "                           when t.mes + 1 > 12              " & vbCrLf & _
                "                             then t.mes + 1 - 12            " & vbCrLf & _
                "                             else t.Mes + 1                 " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "         then 1                                             " & vbCrLf & _
                "         else 0                                             " & vbCrLf & _
                "    end) Mes02,                                             " & vbCrLf & _
                "  SUM(case                                                  " & vbCrLf & _
                "      when Year(data)  = case                               " & vbCrLf & _
                "                           when t.mes + 2 > 12              " & vbCrLf & _
                "                             then t.ano + 1                 " & vbCrLf & _
                "                             else t.ano                     " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "       and MONTH(DATA) = case                               " & vbCrLf & _
                "                           when t.mes + 2 > 12              " & vbCrLf & _
                "                             then t.mes + 2 - 12            " & vbCrLf & _
                "                             else t.Mes + 2                 " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "         then 1                                             " & vbCrLf & _
                "         else 0                                             " & vbCrLf & _
                "    end) Mes03,                                             " & vbCrLf & _
                "  SUM(case                                                  " & vbCrLf & _
                "      when Year(data)  = case                               " & vbCrLf & _
                "                           when t.mes + 3 > 12              " & vbCrLf & _
                "                             then t.ano + 1                 " & vbCrLf & _
                "                             else t.ano                     " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "       and MONTH(DATA) = case                               " & vbCrLf & _
                "                           when t.mes + 3 > 12              " & vbCrLf & _
                "                             then t.mes + 3 - 12            " & vbCrLf & _
                "                             else t.Mes + 3                 " & vbCrLf & _
                "                         End                                " & vbCrLf & _
                "         then 1                                             " & vbCrLf & _
                "         else 0                                             " & vbCrLf & _
                "    end) Mes04,                                             " & vbCrLf & _
                "    SUM(case                                                " & vbCrLf & _
                "      when Year(data)  = case                               " & vbCrLf & _
                "                           when t.mes + 4 > 12              " & vbCrLf & _
                "                             then t.ano + 1                 " & vbCrLf & _
                "                             else t.ano                     " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "       and MONTH(DATA) = case                               " & vbCrLf & _
                "                           when t.mes + 4 > 12              " & vbCrLf & _
                "                             then t.mes + 4 - 12            " & vbCrLf & _
                "                             else t.Mes + 4                 " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "         then 1                                             " & vbCrLf & _
                "         else 0                                             " & vbCrLf & _
                "    end) Mes05,                                             " & vbCrLf & _
                "    SUM(case                                                " & vbCrLf & _
                "      when Year(data)  = case                               " & vbCrLf & _
                "                           when t.mes + 5 > 12              " & vbCrLf & _
                "                             then t.ano + 1                 " & vbCrLf & _
                "                             else t.ano                     " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "       and MONTH(DATA) = case                               " & vbCrLf & _
                "                           when t.mes + 5 > 12              " & vbCrLf & _
                "                             then t.mes + 5 - 12            " & vbCrLf & _
                "                             else t.Mes + 5                 " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "         then 1                                             " & vbCrLf & _
                "         else 0                                             " & vbCrLf & _
                "    end) Mes06,                                             " & vbCrLf & _
                "    SUM(case                                                " & vbCrLf & _
                "      when Year(data)  = case                               " & vbCrLf & _
                "                           when t.mes + 6 > 12              " & vbCrLf & _
                "                             then t.ano + 1                 " & vbCrLf & _
                "                             else t.ano                     " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "       and MONTH(DATA) = case                               " & vbCrLf & _
                "                           when t.mes + 6 > 12              " & vbCrLf & _
                "                             then t.mes + 6 - 12            " & vbCrLf & _
                "                             else t.Mes + 6                 " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "         then 1                                             " & vbCrLf & _
                "         else 0                                             " & vbCrLf & _
                "    end) Mes07,                                             " & vbCrLf & _
                "    SUM(case                                                " & vbCrLf & _
                "      when Year(data)  = case                               " & vbCrLf & _
                "                           when t.mes + 7 > 12              " & vbCrLf & _
                "                             then t.ano + 1                 " & vbCrLf & _
                "                             else t.ano                     " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "       and MONTH(DATA) = case                               " & vbCrLf & _
                "                           when t.mes + 7 > 12              " & vbCrLf & _
                "                             then t.mes + 7 - 12            " & vbCrLf & _
                "                             else t.Mes + 7                 " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "         then 1                                             " & vbCrLf & _
                "         else 0                                             " & vbCrLf & _
                "    end) Mes08,                                             " & vbCrLf & _
                "    SUM(case                                                " & vbCrLf & _
                "      when Year(data)  = case                               " & vbCrLf & _
                "                           when t.mes + 8 > 12              " & vbCrLf & _
                "                             then t.ano + 1                 " & vbCrLf & _
                "                             else t.ano                     " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "       and MONTH(DATA) = case                               " & vbCrLf & _
                "                           when t.mes + 8 > 12              " & vbCrLf & _
                "                             then t.mes + 8 - 12            " & vbCrLf & _
                "                             else t.Mes + 8                 " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "         then 1                                             " & vbCrLf & _
                "         else 0                                             " & vbCrLf & _
                "    end) Mes09,                                             " & vbCrLf & _
                "    SUM(case                                                " & vbCrLf & _
                "      when Year(data)  = case                               " & vbCrLf & _
                "                           when t.mes + 9 > 12              " & vbCrLf & _
                "                             then t.ano + 1                 " & vbCrLf & _
                "                             else t.ano                     " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "       and MONTH(DATA) = case                               " & vbCrLf & _
                "                           when t.mes + 9 > 12              " & vbCrLf & _
                "                             then t.mes + 9 - 12            " & vbCrLf & _
                "                             else t.Mes + 9                 " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "         then 1                                             " & vbCrLf & _
                "         else 0                                             " & vbCrLf & _
                "    end) Mes10,                                             " & vbCrLf & _
                "    SUM(case                                                " & vbCrLf & _
                "      when Year(data)  = case                               " & vbCrLf & _
                "                           when t.mes + 10 > 12             " & vbCrLf & _
                "                             then t.ano + 1                 " & vbCrLf & _
                "                             else t.ano                     " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "       and MONTH(DATA) = case                               " & vbCrLf & _
                "                           when t.mes + 10 > 12             " & vbCrLf & _
                "                             then t.mes + 10 - 12           " & vbCrLf & _
                "                             else t.Mes + 10                " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "         then 1                                             " & vbCrLf & _
                "         else 0                                             " & vbCrLf & _
                "    end) Mes11,                                             " & vbCrLf & _
                "    SUM(case                                                " & vbCrLf & _
                "      when Year(data)  = case                               " & vbCrLf & _
                "                           when t.mes + 11 > 12             " & vbCrLf & _
                "                             then t.ano + 1                 " & vbCrLf & _
                "                             else t.ano                     " & vbCrLf & _
                "                         end                                " & vbCrLf & _
                "       and MONTH(DATA) = case                               " & vbCrLf & _
                "                           when t.mes +11 > 12              " & vbCrLf & _
                "                             then t.mes + 11 - 12           " & vbCrLf & _
                "                             else t.Mes + 11                " & vbCrLf & _
                "                          end                               " & vbCrLf & _
                "         then 1                                             " & vbCrLf & _
                "         else 0                                             " & vbCrLf & _
                "    end) Mes12                                              " & vbCrLf & _
                "  into #TempVisita " & vbCrLf & _
                "  from Visita V" & vbCrLf & _
                " inner join #Temp t " & vbCrLf & _
                "    on 1 = 1 " & vbCrLf & _
                " inner join Clientes c " & vbCrLf & _
                "    on V.Cliente    = c.Cliente_Id" & vbCrLf & _
                "   and V.EndCliente = c.Endereco_Id" & vbCrLf & _
                " inner join CRMXParametroVisita CxPv" & vbCrLf & _
                "    ON V.Cliente     = CxPv.Cliente_Id " & vbCrLf & _
                "   AND V.EndCliente  = CxPv.EndCliente_Id" & vbCrLf & _
                " where CxPv.Empresa_Id     = '" & CodEmpresa & "'" & vbCrLf & _
                "   and CxPv.EndEmpresa_Id  =" & EndEmpresa & vbCrLf & _
                "   and CxPv.Ano_Id         = " & ano & vbCrLf & _
                "   and CxPv.Consolidado_Id = 0   " & vbCrLf & _
                " group by V.Cliente, " & vbCrLf & _
                "          V.EndCliente," & vbCrLf & _
                "          C.Nome," & vbCrLf & _
                "          CxPv.MinimoVisitas, " & vbCrLf & _
                "          t.datasafra," & vbCrLf & _
                "          t.mesatual " & vbCrLf & _
                " SELECT CodigoCliente,                                                                                            " & vbCrLf & _
                "        EndCliente,                                                                                               " & vbCrLf & _
                "        NomeCliente,                                                                                              " & vbCrLf & _
                "        InicioCRM,                                                                                                " & vbCrLf & _
                "        MesAtual,                                                                                                 " & vbCrLf & _
                "        MinimoVisitas,                                                                                            " & vbCrLf & _
                "        Mes01, Mes02, Mes03, Mes04, Mes05, Mes06, Mes07, Mes08, Mes09, Mes10, Mes11, Mes12,                       " & vbCrLf & _
                "       (Mes01+ Mes02+ Mes03+ Mes04+ Mes05+ Mes06+ Mes07+ Mes08+ Mes09+ Mes10+ Mes11+ Mes12) as TotalAno,          " & vbCrLf & _
                "       (Mes01+ Mes02+ Mes03+ Mes04+ Mes05+ Mes06+ Mes07+ Mes08+ Mes09+ Mes10+ Mes11+ Mes12)                       " & vbCrLf & _
                "         - (MinimoVisitas*MesAtual) as SaldoDeVisitaObrigatorio,                                                            " & vbCrLf & _
                "       CEILING(CONVERT(NUMERIC(18,2), MinimoVisitas/                                                             " & vbCrLf & _
                "       (case when MinimoVisitas=0.00 then 1.00 else MinimoVisitas end*12)* MesAtual )*100) AS IPVMin,            " & vbCrLf & _
                "	    CEILING((Mes01+ Mes02+ Mes03+ Mes04+ Mes05+ Mes06+ Mes07+ Mes08+ Mes09+ Mes10+ Mes11+ Mes12)               " & vbCrLf & _
                "		  / CONVERT(NUMERIC(18,2), (case when MinimoVisitas=0 then 1 else MinimoVisitas end * 12)) * 100) AS IPVAno " & vbCrLf & _
                " FROM #TempVisita ")

            ds = Banco.ConsultaDataSet(sql.ToString.Trim, "Visitas")

            If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                DataInicial = ds.Tables(0).Rows(0)("InicioCRM").ToString
                grdGerenciamentoVisita.DataSource = ds
                grdGerenciamentoVisita.DataBind()
            End If

        End If
    End Sub

    Protected Sub grdGerenciamentoVisita_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grdGerenciamentoVisita.RowDataBound
        'Coluna com o Header Mês 01'         
        Dim pos As Integer = 3
        If e.Row.RowType = DataControlRowType.Header Then
            For i As Integer = 1 To 12 Step 1
                e.Row.Cells(pos).Text = DataInicial.Month & "/" & DataInicial.Year
                pos = pos + 1
                DataInicial = DateAdd(DateInterval.Month, 1, DataInicial)
            Next
        End If
    End Sub


    Protected Sub lnkFormularioCadVisita_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFormularioCadVisita.Click
        Try
            If Funcoes.VerificaPermissao("CRMxVisita", "RELATORIO") Then
                ds = New DataSet()
                Funcoes.BindReport(Me.Page, ds, "Cr_VisitaFormulario", eExportType.PDF, Nothing, True)
            Else
                MsgBox(Me.Page, "Usuario sem permissao para tirar Relatório")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjudaCadVisita_Click(sender As Object, e As EventArgs) Handles lnkAjudaCadVisita.Click
        Try
            Funcoes.Ajuda(Me.Page, "CRMxVisita")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class