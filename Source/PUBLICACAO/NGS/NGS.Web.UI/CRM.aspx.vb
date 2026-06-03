Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Reflection
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class CRM
    Inherits BasePage

    Dim ListGrupoCompra As [Lib].Negocio.ListGrupoProduto
    Dim CRMParametros As [Lib].Negocio.CRM
    Dim CRMListaConsulta As ListCRM

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            'If Funcoes.VerificaPermissao("CRM", "ACESSAR") Then
            GerarNovoHID()

            txtPesoMargemBrutaVenda_TextChanged(Nothing, Nothing)
            txtPesoParticipacaoSOC_TextChanged(Nothing, Nothing)
            txtPesoRatingDeCredito_TextChanged(Nothing, Nothing)

            ucSelecaoProdutoVendas.NomeUC = "VENDAS / INSUMOS"
            ucSelecaoProdutoCompras.NomeUC = "COMPRAS / GRAOS"

            ddl.Carregar(ddlEmpresaVenda, CarregarDDL.Tabela.Empresas, "")
            ddl.Carregar(ddlEmpresaCompra, CarregarDDL.Tabela.Empresas, "")
            ddl.Carregar(ddlAno, CarregarDDL.Tabela.Ano, "2016;10;C", True)
            ddl.Carregar(ddlAnoConsulta, CarregarDDL.Tabela.Ano, "2016;10;C", True)
            LiberaEmpresa()
            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
            '    Exit Sub
            'End If
        End If
    End Sub

#Region "SESSAO"
    Private Sub SessaoSalvarCRMParametros()
        Session("CRMParametros" & HID.Value) = CRMParametros
    End Sub

    Private Sub SessaoRecuperaCRMParametros()
        If Session("CRMParametros" & HID.Value) Is Nothing Then
            CRMParametros = New [Lib].Negocio.CRM()
        Else
            CRMParametros = Session("CRMParametros" & HID.Value)
        End If
    End Sub

    Private Sub SessaoSalvarCRMListaConsulta()
        Session("CRMListaConsulta" & HID.Value) = CRMListaConsulta
    End Sub

    Private Sub SessaoRecuperaCRMListaConsulta()
        CRMListaConsulta = Session("CRMListaConsulta" & HID.Value)
    End Sub
#End Region

#Region "Premissas"
    Protected Sub txtPesoVenda_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If CDec(txtPesoVenda.Text) < 0 Or CDec(txtPesoVenda.Text) > 100 Then
            txtPesoVenda.Text = "70"
            txtPesoCompra.Text = "30"
            Exit Sub
        End If

        txtPesoCompra.Text = 100 - CDec(txtPesoVenda.Text)
    End Sub

    Protected Sub txtPesoCompra_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If CDec(txtPesoCompra.Text) < 0 Or CDec(txtPesoCompra.Text) > 100 Then
            txtPesoVenda.Text = "70"
            txtPesoCompra.Text = "30"
            Exit Sub
        End If

        txtPesoVenda.Text = 100 - CDec(txtPesoCompra.Text)
    End Sub

    Protected Sub txtPesoMargemBrutaVenda_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        lblMBVMenor.Text = CDec(txtPesoMargemBrutaVenda.Text) * 9 / 100
        lblMBVEntre.Text = CDec(txtPesoMargemBrutaVenda.Text) * 18 / 100
        lblMBVMaior.Text = CDec(txtPesoMargemBrutaVenda.Text) * 27 / 100
    End Sub

    Protected Sub txtPesoParticipacaoSOC_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        lblSocMenor.Text = CDec(txtPesoParticipacaoSOC.Text) * 9 / 100
        lblSocEntre.Text = CDec(txtPesoParticipacaoSOC.Text) * 18 / 100
        lblSocMaior.Text = CDec(txtPesoParticipacaoSOC.Text) * 27 / 100
    End Sub

    Protected Sub txtPesoRatingDeCredito_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        lblRatingC.Text = CDec(txtPesoRatingDeCredito.Text) * 9 / 100
        lblRatingB.Text = CDec(txtPesoRatingDeCredito.Text) * 18 / 100
        lblRatingA.Text = CDec(txtPesoRatingDeCredito.Text) * 27 / 100
    End Sub
#End Region

    Protected Sub btnPremissas_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPremissas.Click
        SessaoRecuperaCRMParametros()
        If CRMParametros.IUD = "U" Then
            TabCRM.ActiveTabIndex = 2
            Exit Sub
        End If

        If ddlEmpresaVenda.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione a Empresa")
            Exit Sub
        End If

        If ddlAno.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione o Ano")
            Exit Sub
        End If

        Dim xv As String = ""
        If Not ucSelecaoProdutoVendas.TemSelecionado Then
            MsgBox(Me.Page, "Selecione o(s) Grupo(s) e o(s) Produto(s) para a curva de Venda")
            Exit Sub
        End If

        Dim xc As String = ""
        If chkConsiderarCompras.Checked Then
            If Not ucSelecaoProdutoCompras.TemSelecionado Then
                MsgBox(Me.Page, "Selecione o(s) Grupo(s) e o(s) Produto(s) para a curva de Compra")
                Exit Sub
            End If
        End If

        '*********************************************
        '********* Carregamento **********************
        '*********************************************
        Dim ObjAnalise As [Lib].Negocio.ParametrosAnaliseDeCredito = New [Lib].Negocio.ParametrosAnaliseDeCredito(ddlAno.SelectedValue)

        Dim Emp As String()
        Emp = ddlEmpresaVenda.SelectedValue.Split("-")

        CRMParametros = New [Lib].Negocio.CRM(Emp(0), Emp(1), ddlAno.SelectedValue, chkConsolidarEmpresaVenda.Checked)

        If CRMParametros.IUD = "U" Then
            MsgBox(Me.Page, "Já Existe esta analise configurada, ela sera carregada, para faze-la novamente consulte e exclua.")
            SessaoSalvarCRMParametros()
            AtualizaformComAClasse()
            Exit Sub
        End If

        CRMParametros.IUD = "I"
        CRMParametros.Definicao = ObjAnalise.DefinicaoAno
        CRMParametros.Mercado = IIf(rdTodosMercados.Checked, "T", IIf(rdInterno.Checked, "I", "E"))


        CRMParametros.ProdutosVenda = ucSelecaoProdutoVendas.GetStringGrupoProdutoSelecionado

        If chkConsiderarCompras.Checked Then
            CRMParametros.ConsideraCompra = True
            CRMParametros.ProdutosCompra = ucSelecaoProdutoCompras.GetStringGrupoProdutoSelecionado
            If Not chkConsolidarEmpresaVenda.Checked Then
                If ddlEmpresaCompra.SelectedIndex > 0 Then
                    Dim EmpresaCompra As String()
                    EmpresaCompra = ddlEmpresaCompra.SelectedValue.Split("-")
                    CRMParametros.CodigoEmpresaCompra = EmpresaCompra(0)
                    CRMParametros.EndEmpresaCompra = EmpresaCompra(1)
                Else
                    CRMParametros.CodigoEmpresaCompra = ""
                    CRMParametros.EndEmpresaCompra = 0
                End If
            End If
        Else
            CRMParametros.ConsideraCompra = False
            CRMParametros.CodigoEmpresaCompra = ""
            CRMParametros.EndEmpresaCompra = 0
            CRMParametros.ProdutosCompra = ""
        End If

        CRMParametros.CrmClientes.Atualizar()
        gridSelecaoClientes.DataSource = CRMParametros.CrmClientes.ToArray
        gridSelecaoClientes.DataBind()

        SessaoSalvarCRMParametros()
        lblCorte.Text = CRMParametros.CRMValorCorte.ToString("N2")

        TabCRM.ActiveTabIndex = 2

    End Sub

    Protected Sub btnSelecaoClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSelecaoClientes.Click
        SessaoRecuperaCRMParametros()
        If CRMParametros.IUD = "U" Then
            TabCRM.ActiveTabIndex = 3
            Exit Sub
        End If

        SelecaoCliente()
        TabCRM.ActiveTabIndex = 3
    End Sub

    Protected Sub btnLimparPremissas_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLimparPremissas.Click
        LimparPremissas()
        SelecaoCliente()
    End Sub

    Public Sub SelecaoCliente()
        If CDec(txtPesoMargemBrutaVenda.Text) + CDec(txtPesoParticipacaoSOC.Text) + CDec(txtPesoRatingDeCredito.Text) <> 100 Then
            MsgBox(Me.Page, "A soma dos pesos da Margem Bruta + SOC + Rating de Credito deve ser igual a 100")
            Exit Sub
        End If

        If CDec(txtPesoVenda.Text) + CDec(txtPesoCompra.Text) <> 100 Then
            MsgBox(Me.Page, "A soma dos pesos da Venda e Compra no SOC deve ser igual a 100")
            Exit Sub
        End If

        SessaoRecuperaCRMParametros()
        CRMParametros.CRMPercentualCorte = txtPercentualCorte.Text

        CRMParametros.MBVPercentualMenor = txtMBVPercMenor.Text
        CRMParametros.MBVPercentualMaior = txtMBVPercMaior.Text

        CRMParametros.SOCPercentualMenor = txtSOCPercMenor.Text
        CRMParametros.SOCPercentualMaior = txtSOCPercMaior.Text

        CRMParametros.MBVPeso = txtPesoMargemBrutaVenda.Text
        CRMParametros.SOCPeso = txtPesoParticipacaoSOC.Text
        CRMParametros.RatingPeso = txtPesoRatingDeCredito.Text

        CRMParametros.SOCPesoVenda = txtPesoVenda.Text
        CRMParametros.SOCPesoCompra = txtPesoCompra.Text

        CRMParametros.CrmClientes.Atualizar()
        SessaoSalvarCRMParametros()
        AtualizaformComAClasse()
    End Sub

    Public Sub AtualizaformComAClasse()
        SessaoRecuperaCRMParametros()
        '*********************************************************************
        '*********************  CURVA  ***************************************
        '*********************************************************************
        ddlEmpresaVenda.SelectedValue = CRMParametros.CodigoEmpresa & "-" & CRMParametros.EndEmpresa
        chkConsolidarEmpresaVenda.Checked = CRMParametros.Consolidado

        pnlEmpresaCompra.Visible = Not CRMParametros.Consolidado

        If CRMParametros.CodigoEmpresaCompra.Length > 0 Then
            ddlEmpresaCompra.SelectedValue = CRMParametros.CodigoEmpresaCompra & "-" & CRMParametros.EndEmpresaCompra
        End If

        ddlAno.SelectedValue = CRMParametros.Ano

        Select Case CRMParametros.Mercado
            Case "I" : rdInterno.Checked = True
            Case "E" : rdExterno.Checked = True
            Case "T" : rdTodosMercados.Checked = True
        End Select

        chkConsiderarCompras.Checked = CRMParametros.ConsideraCompra

        ucSelecaoProdutoCompras.PreenheGrid(CRMParametros.ProdutosCompra)
        ucSelecaoProdutoVendas.PreenheGrid(CRMParametros.ProdutosVenda)

        '*********************************************************************
        '*******************  PREMISSAS **************************************
        '*********************************************************************
        txtPercentualCorte.Text = CRMParametros.CRMPercentualCorte.ToString("N0")
        lblCorte.Text = CRMParametros.CRMValorCorte.ToString("N2")

        txtMBVPercMenor.Text = CRMParametros.MBVPercentualMenor.ToString("N0")
        txtMBVPercDe.Text = CRMParametros.MBVPercentualEntreDe.ToString("N0")
        txtMBVPercAte.Text = CRMParametros.MBVPercentualEntreAte.ToString("N0")
        txtMBVPercMaior.Text = CRMParametros.MBVPercentualMaior.ToString("N0")

        txtSOCPercMenor.Text = CRMParametros.SOCPercentualMenor.ToString("N0")
        txtSOCPercDe.Text = CRMParametros.SOCPercentualEntreDe.ToString("N0")
        txtSOCPerAte.Text = CRMParametros.SOCPercentualEntreAte.ToString("N0")
        txtSOCPercMaior.Text = CRMParametros.SOCPercentualMaior.ToString("N0")

        txtPesoMargemBrutaVenda.Text = CRMParametros.MBVPeso.ToString("N0")
        txtPesoParticipacaoSOC.Text = CRMParametros.SOCPeso.ToString("N0")
        txtPesoRatingDeCredito.Text = CRMParametros.RatingPeso.ToString("N0")

        txtPesoVenda.Text = CRMParametros.SOCPesoVenda.ToString("N0")
        txtPesoCompra.Text = CRMParametros.SOCPesoCompra.ToString("N0")

        lblMBVMenor.Text = CRMParametros.MBVPontuacaoMenor.ToString("N2")
        lblMBVEntre.Text = CRMParametros.MBVPontuacaoEntre.ToString("N2")
        lblMBVMaior.Text = CRMParametros.MBVPontuacaoMaior.ToString("N2")

        lblSocMenor.Text = CRMParametros.SOCPontuacaoMenor.ToString("N2")
        lblSocEntre.Text = CRMParametros.SOCPontuacaoEntre.ToString("N2")
        lblSocMaior.Text = CRMParametros.SOCPontuacaoMaior.ToString("N2")

        lblRatingA.Text = CRMParametros.RatingPontuacaoA.ToString("N2")
        lblRatingB.Text = CRMParametros.RatingPontuacaoB.ToString("N2")
        lblRatingC.Text = CRMParametros.RatingPontuacaoC.ToString("N2")

        '*********************************************************************
        '*******************  CLIENTES ***************************************
        '*********************************************************************
        gridSelecaoClientes.DataSource = CRMParametros.CrmClientes.ToArray
        gridSelecaoClientes.DataBind()

        SessaoSalvarCRMParametros()
    End Sub

    Public Sub LimparPremissas()
        txtPercentualCorte.Text = "80"
        lblCorte.Text = "0,00"

        txtPesoMargemBrutaVenda.Text = "34"
        txtPesoParticipacaoSOC.Text = "33"
        txtPesoRatingDeCredito.Text = "33"

        txtPesoVenda.Text = "70"
        txtPesoCompra.Text = "30"

        txtMBVPercMenor.Text = "10"
        txtMBVPercDe.Text = "10"
        txtMBVPercAte.Text = "15"
        txtMBVPercMaior.Text = "15"

        txtSOCPercMenor.Text = "50"
        txtSOCPercDe.Text = "50"
        txtSOCPerAte.Text = "70"
        txtSOCPercMaior.Text = "70"

        lblMBVMenor.Text = "3"
        lblMBVEntre.Text = "6"
        lblMBVMaior.Text = "9"

        lblSocMenor.Text = "3"
        lblSocEntre.Text = "6"
        lblSocMaior.Text = "9"

        lblRatingC.Text = "3"
        lblRatingB.Text = "6"
        lblRatingA.Text = "9"

    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresaVenda.Enabled = False
        End If
    End Sub

    Public Sub AbrirAccordion(ByVal AccordionId As String, ByVal AbrirAccordion As Boolean, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('#" & AccordionId & "').accordion('activate',  " & IIf(AbrirAccordion, 0, 1) & ");"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Protected Sub chkConsiderarCompras_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkConsiderarCompras.CheckedChanged
        If Not chkConsiderarCompras.Checked Then
            ucSelecaoProdutoCompras.Visible = False

            txtPesoVenda.Enabled = False
            txtPesoVenda.Text = "100"

            txtPesoCompra.Enabled = False
            txtPesoCompra.Text = "0"
            pnlEmpresaCompra.Visible = False
        Else
            ucSelecaoProdutoCompras.Visible = True

            txtPesoVenda.Enabled = True
            txtPesoVenda.Text = "70,00"

            txtPesoCompra.Enabled = True
            txtPesoCompra.Text = "30,00"
            pnlEmpresaCompra.Visible = True
        End If
    End Sub

    Protected Sub txtPercentualCorte_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaCRMParametros()
        CRMParametros.CRMPercentualCorte = txtPercentualCorte.Text
        CRMParametros.CrmClientes.Atualizar()
        lblCorte.Text = CRMParametros.CRMValorCorte.ToString("N2")
        SessaoSalvarCRMParametros()
    End Sub

    Protected Sub txtMBVPercMenor_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        txtMBVPercDe.Text = txtMBVPercMenor.Text
    End Sub

    Protected Sub txtMBVPercMaior_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        txtMBVPercAte.Text = txtMBVPercMaior.Text
    End Sub

    Protected Sub txtSOCPercMenor_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        txtSOCPercDe.Text = txtSOCPercMenor.Text
    End Sub

    Protected Sub txtSOCPercMaior_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        txtSOCPerAte.Text = txtSOCPercMaior.Text
    End Sub

    Protected Sub ddlAnoConsulta_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlAnoConsulta.SelectedIndexChanged
        CRMListaConsulta = New ListCRM(ddlAnoConsulta.SelectedValue)
        gridConsulta.DataSource = CRMListaConsulta.ToArray
        gridConsulta.DataBind()
        SessaoSalvarCRMListaConsulta()
    End Sub

    Protected Sub btnSalvar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSalvar.Click
        SessaoRecuperaCRMParametros()
        If CRMParametros.Salvar Then
            MsgBox(Me.Page, "CRM salvo com Sucesso.", eTitulo.Sucess)
        Else
            MsgBox(Me.Page, "Erro ao salvar CRM")
        End If
    End Sub

    Protected Sub gridConsulta_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles gridConsulta.SelectedIndexChanged
        SessaoRecuperaCRMListaConsulta()

        Session.Remove("CRMParametros" & HID.Value)
        Session.Remove("CRMListaConsulta" & HID.Value)

        GerarNovoHID()

        CRMParametros = CRMListaConsulta(gridConsulta.SelectedIndex)
        CRMParametros.IUD = "U"

        SessaoSalvarCRMParametros()
        SessaoSalvarCRMListaConsulta()
        AtualizaformComAClasse()

        TabCRM.ActiveTabIndex = 1
        'AbrirAccordion("ucSelecaoProdutoCompras", False, True, 100)
        'AbrirAccordion("ucSelecaoProdutoVendas", False, True, 100)
        'AbrirAccordion("divCurvaABC", True, True, 100)
    End Sub

    Public Sub GerarNovoHID()
        HID.Value = Guid.NewGuid.ToString
    End Sub

    Protected Sub ddlTipo_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        SessaoRecuperaCRMParametros()

        Dim ddltipo As DropDownList = CType(sender, DropDownList)
        Dim row As GridViewRow = CType(ddltipo.NamingContainer, GridViewRow)

        CRMParametros.CrmClientes(row.RowIndex).TipoClienteCRM = ddltipo.SelectedValue

        CRMParametros.CrmClientes.Atualizar()
        gridSelecaoClientes.DataSource = CRMParametros.CrmClientes.ToArray
        gridSelecaoClientes.DataBind()
        SessaoSalvarCRMParametros()
        lblCorte.Text = CRMParametros.CRMValorCorte.ToString("N2")
    End Sub

    Protected Sub ddlQualitativo_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        SessaoRecuperaCRMParametros()

        Dim ddlQualitativo As DropDownList = CType(sender, DropDownList)
        Dim row As GridViewRow = CType(ddlQualitativo.NamingContainer, GridViewRow)

        CRMParametros.CrmClientes(row.RowIndex).TipoClienteQualitativo = ddlQualitativo.SelectedValue
        SessaoSalvarCRMParametros()
    End Sub

    Protected Sub btnNovaAnalise_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnNovaAnalise.Click
        ddlEmpresaVenda.SelectedIndex = 0
        ddlAno.SelectedIndex = 0
        rdTodosMercados.Checked = True
        chkConsiderarCompras.Checked = True
        CType(ucSelecaoProdutoVendas.FindControlRecursive("rdnivel1"), CheckBox).Checked = True
        ucSelecaoProdutoVendas.CarregarNivel(1)
        CType(ucSelecaoProdutoCompras.FindControlRecursive("rdnivel1"), CheckBox).Checked = True
        ucSelecaoProdutoCompras.CarregarNivel(1)
        TabCRM.ActiveTabIndex = 1


        SessaoRecuperaCRMListaConsulta()
        Session.Remove("CRMParametros" & HID.Value)
        Session.Remove("CRMListaConsulta" & HID.Value)
        GerarNovoHID()
        CRMParametros = New [Lib].Negocio.CRM
        CRMParametros.IUD = "I"
        SessaoSalvarCRMParametros()
        SessaoSalvarCRMListaConsulta()
    End Sub

    Protected Sub imgImpressaoPend_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim btnImprimir As ImageButton = CType(sender, ImageButton)
        Dim Gridrow As GridViewRow = CType(btnImprimir.NamingContainer, GridViewRow)

        SessaoRecuperaCRMListaConsulta()
        CRMParametros = CRMListaConsulta(Gridrow.RowIndex)

        Dim ds As New DataSet
        Dim dsCli As New DataSet

        Dim c As New Structure2DataTable(Of [Lib].Negocio.CRM)
        ds.Merge(c.GetDataTable(CRMParametros))

        Dim c1 As New Structure2DataTable(Of [Lib].Negocio.CRMxCliente)
        For Each obj In CRMParametros.CrmClientes
            If Not chkImprimirCrm.Checked Then
                If ds.Tables.Count = 1 Then
                    ds.Tables.Add(c1.GetDataTable(obj))
                Else
                    ds.Tables(1).ImportRow((c1.GetDataTable(obj).Rows(0)))
                End If
            ElseIf obj.TipoClienteCRM <> "V" Then
                If ds.Tables.Count = 1 Then
                    ds.Tables.Add(c1.GetDataTable(obj))
                Else
                    ds.Tables(1).ImportRow((c1.GetDataTable(obj).Rows(0)))
                End If
            End If
        Next

        Dim soma As Integer
        dsCli = Banco.ConsultaDataSet(Sql, "TipoCliente")

        ds.Tables(0).TableName = "CRM"
        ds.Tables(1).TableName = "CLIENTES"

        ds.Tables.Add("TipoCliente")
        ds.Tables("TipoCliente").Columns.Add("DescricaoTipo")
        ds.Tables("TipoCliente").Columns.Add("Qtde")

        For Each row As DataRow In dsCli.Tables("TipoCliente").Rows
            Dim r As DataRow = ds.Tables("TipoCliente").NewRow()
            r("DescricaoTipo") = row("DescricaoTipo")
            r("Qtde") = row("Qtde")
            soma += CInt(row("Qtde"))
            ds.Tables("TipoCliente").Rows.Add(r)
        Next

        '*************************************************************************
        Dim cols As New ArrayList
        Dim dsRelatorio As New Ds_CRM

        For Each col As Data.DataColumn In ds.Tables("CRM").Columns
            If Not dsRelatorio.Tables("CRM").Columns.Contains(col.ColumnName) Then
                cols.Add(col)
            End If
        Next
        For i = 0 To cols.Count - 1
            ds.Tables("CRM").Columns.Remove(cols(i))
        Next

        cols.Clear()
        For Each col As Data.DataColumn In ds.Tables("CLIENTES").Columns
            If Not dsRelatorio.Tables("CLIENTES").Columns.Contains(col.ColumnName) Then
                cols.Add(col)
            End If
        Next
        For i = 0 To cols.Count - 1
            ds.Tables("CLIENTES").Columns.Remove(cols(i))
        Next

        '*************************************************************************
        Dim logo As New DataTable("Logotipo")
        logo.Columns.Add("path", GetType(String))
        logo.Columns.Add("Imagem", GetType(System.Byte()))
        Dim drImagem As DataRow = logo.NewRow()
        Dim caminhoImagem As String = HttpContext.Current.Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
        drImagem("path") = caminhoImagem
        drImagem("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(caminhoImagem)
        logo.Rows.Add(drImagem)
        ds.Tables.Add(logo)

        Dim rpt As New ReportDocument()

        rpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/Cr_CRM.rpt")
        rpt.Load(rpt.FileName, OpenReportMethod.OpenReportByDefault)

        Dim UrlArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
        Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(UrlArquivo)

        rpt.SetDataSource(ds)

        Dim parameters = New Dictionary(Of String, Object)()
        parameters("Total") = soma

        If CRMParametros.ProdutosVenda.Length > 0 Then
            parameters("ProdutosVenda") = ucSelecaoProdutoVendas.GetSqlEParametrosRelatorio("", "", CRMParametros.ProdutosVenda)(1)
        Else
            parameters("ProdutosVenda") = ""
        End If

        If CRMParametros.ProdutosCompra.Length > 0 Then
            parameters("ProdutosCompra") = ucSelecaoProdutoCompras.GetSqlEParametrosRelatorio("", "", CRMParametros.ProdutosCompra)(1)
        Else
            parameters("ProdutosCompra") = ""
        End If

        Funcoes.BindParameters(rpt, parameters)

        If Dir(NomeArquivo).Length > 0 Then
            Kill(NomeArquivo)
        End If

        Try
            rpt.ExportToDisk(ExportFormatType.PortableDocFormat, NomeArquivo)
            If System.IO.File.Exists(NomeArquivo) Then
                Funcoes.AbrirArquivo(Page, UrlArquivo)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            rpt.Close()
            rpt.Dispose()
        End Try
    End Sub

    Function Sql() As String
        Return "            " & vbCrLf & _
                            "  select case when TipoCliente = 'C' then 'CRM' " & vbCrLf & _
                            "              when TipoCliente = 'E' then 'Estrategico' " & vbCrLf & _
                            "              when TipoCliente = 'M' then 'Cliente De Massa' " & vbCrLf & _
                            "			  when TipoCliente = 'P' then 'Prospect'" & vbCrLf & _
                            "			  when TipoCliente = 'V' then 'CRM Nao Participante'" & vbCrLf & _
                            "		 end as DescricaoTipo, COUNT(*) as Qtde" & vbCrLf & _
                            "   from CRMParametroXCliente" & vbCrLf & _
                            "  group by TipoCliente" & vbCrLf
    End Function

    Protected Sub imgExcluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim btnExcluir As ImageButton = CType(sender, ImageButton)
        Dim Gridrow As GridViewRow = CType(btnExcluir.NamingContainer, GridViewRow)

        SessaoRecuperaCRMListaConsulta()
        CRMParametros = CRMListaConsulta(Gridrow.RowIndex)
        CRMParametros.IUD = "D"
        If CRMParametros.Salvar Then
            MsgBox(Me.Page, "Crm Excluido com Sucesso.", eTitulo.Sucess)
            CRMListaConsulta = New ListCRM(ddlAnoConsulta.SelectedValue)
            gridConsulta.DataSource = CRMListaConsulta.ToArray
            gridConsulta.DataBind()
            SessaoSalvarCRMListaConsulta()
        Else
            MsgBox(Me.Page, "Erro ao Excluir o CRM")
        End If
    End Sub

End Class


Public Class Structure2DataTable(Of T)
    Private dt As New DataTable()

    Public Function GetDataTable(ByVal t As T) As DataTable
        Dim dt As New DataTable()

        Dim tipo As Type = t.[GetType]()

        If tipo.IsArray Then
            Dim arr As Array = TryCast(t, Array)

            For Each obj As Object In arr
                tipo = obj.[GetType]()
                Dim propriedades As PropertyInfo() = tipo.GetProperties()

                If propriedades.Length > 0 Then
                    Dim row As DataRow = dt.NewRow()

                    For Each p As PropertyInfo In propriedades
                        If Not dt.Columns.Contains(p.Name) Then
                            dt.Columns.Add(p.Name, p.PropertyType)
                        End If
                        row(p.Name) = p.GetValue(obj, Nothing)
                    Next

                    dt.Rows.Add(row)
                    dt.AcceptChanges()
                End If
            Next
        Else
            Dim propriedades As PropertyInfo() = tipo.GetProperties()

            If propriedades.Length > 0 Then
                Dim row As DataRow = dt.NewRow()

                For Each p As PropertyInfo In propriedades
                    If Not dt.Columns.Contains(p.Name) Then
                        dt.Columns.Add(p.Name, p.PropertyType)
                    End If
                    row(p.Name) = p.GetValue(t, Nothing)
                Next

                dt.Rows.Add(row)
                dt.AcceptChanges()
            End If
        End If

        Return dt
    End Function

End Class