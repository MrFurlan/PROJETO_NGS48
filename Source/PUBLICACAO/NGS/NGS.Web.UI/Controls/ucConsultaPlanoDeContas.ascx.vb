Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaPlanoDeContas
    Inherits BaseUserControl

    Private strSQL As String
    Private Conta As String
    Private Titulo As String
    Private Cliente As String
    Private TipoDeCliente As Integer
    Private Produto As String
    Private CentroDeCusto As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            ddl.Carregar(ddlTipoDeConta, CarregarDDL.Tabela.TipoDaContaContabil, "")
        End If
        btnFechar.Parent.Visible = False
    End Sub

    Public Sub BindGridView(ByVal pPermitirTodasAsContas As Boolean)
        pnlParametros.Visible = True
        GridGruposDeContas.Columns(0).Visible = pPermitirTodasAsContas
        CargaGruposDeContas(1)
    End Sub

    Public Sub BindGridViewTemEncargo(ByVal pPermitirTodasAsContas As Boolean, ByVal pReceberPagar As String)
        GridGruposDeContas.Columns(0).Visible = pPermitirTodasAsContas
        ddlTemEncargoFinanceiro.SelectedValue = "S"
        ddlTemEncargoFinanceiro.Enabled = False
        ViewState.Add("TemEncargo", True)
        ViewState.Add("pReceberPagar", pReceberPagar)
        CargaGruposDeContas(4, pReceberPagar)
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        txtConta.Text = ""
        txtTitulo.Text = ""
        ddlTipoDeConta.SelectedIndex = 0
        Dim opcao As Integer = 1
        ddlTemCliente.SelectedIndex = 0
        ddlTemProduto.SelectedIndex = 0
        ddlTemCentroCusto.SelectedIndex = 0
        If Not ViewState("TemEncargo") IsNot Nothing Then
            ddlTemEncargoFinanceiro.SelectedIndex = 0
            ddlTemEncargoFinanceiro.Enabled = True
            opcao = 4
        End If
        ddlTemPedido.SelectedIndex = 0
        ddlAdiantamento.SelectedIndex = 0
        GridContasAnaliticas.Parent.Visible = False

        lnkConsultar.Parent.Visible = True
        lnkLimpar.Parent.Visible = True

        ddlTemEncargoFinanceiro.ClearSelection()
        ddlTemEncargoFinanceiro.Enabled = True
        ViewState.Remove("TemEncargo")
        ViewState.Remove("pReceberPagar")

        GridGruposDeContas.DataSource = Nothing
        GridGruposDeContas.DataBind()
        GridContasAnaliticas.DataSource = Nothing
        GridContasAnaliticas.DataBind()

        CargaGruposDeContas(opcao)
        GridGruposDeContas.Parent.Visible = True
    End Sub

    Protected Overrides Sub Selecionar(ByVal args As String)
        Try
            Dim objConta As New [Lib].Negocio.PlanoDeConta("", 0, args)
            Session(Session("ssTipoRetorno")) = objConta
            Session("Debito_Conta") = objConta.Conta
            Session("Debito_Titulo") = objConta.Titulo
            Session("TemCliente" & HID.Value) = IIf(objConta.TemCliente, "S", "N")
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim ucName = MainUserControl.ClientID.Split("_")
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objConta)
                Else
                    CType(Me.Page, IBasePage).Carregar(objConta)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaPlanoDeContas")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaGruposDeContas(ByVal opcao As Integer, Optional ByVal pReceberPagar As String = "C")
        If txtConta.Text.Length > 0 AndAlso Not IsNumeric(txtConta.Text) Then txtConta.Text = ""
        Dim ds As New DataSet

        strSQL = " SELECT Conta_Id, " & vbCrLf & _
                 " case " & vbCrLf & _
                 " when len(Conta_Id) = 9" & vbCrLf & _
                 "  then (select titulo from planodecontas c where c.conta_id = left(PlanoDeContas.conta_id,7)) + ' - ' " & vbCrLf & _
                 "  else ''" & vbCrLf & _
                 "  end + Titulo AS Titulo, " & vbCrLf & _
                 " Cliente, TipoDeCliente, Produto, CentroDeCusto" & vbCrLf & _
                 "  FROM PlanoDeContas " & vbCrLf & _
                 " WHERE 1 = 1"

        If opcao <> 3 Then
            If txtConta.Text.Length > 0 Then strSQL &= " And Conta_Id like '" & txtConta.Text.Trim() & "%'"
            If txtTitulo.Text.Length > 0 Then strSQL &= " And Titulo like '%" & txtTitulo.Text.Trim().ToUpper() & "%'"
            If ddlTipoDeConta.SelectedIndex > 0 Then strSQL &= " And TipoDeConta =" & ddlTipoDeConta.SelectedValue

            If ddlTemCliente.SelectedIndex > 0 Then strSQL &= " And case when isnull(cliente,'') = '' then 'N' else cliente end = '" & ddlTemCliente.SelectedValue & "'"
            If ddlTemProduto.SelectedIndex > 0 Then strSQL &= " And case when isnull(Produto,'') = '' then 'N' else Produto end = '" & ddlTemProduto.SelectedValue & "'"
            If ddlTemCentroCusto.SelectedIndex > 0 Then strSQL &= " And case when isnull(CentroDeCusto,'') = '' then 'N' else CentroDeCusto end = '" & ddlTemCentroCusto.SelectedValue & "'"
            If ddlTemEncargoFinanceiro.SelectedIndex > 0 Then strSQL &= " And isnull(Encargo,0) = " & IIf(ddlTemEncargoFinanceiro.SelectedValue = "S", "1", "0")
            If ddlTemPedido.SelectedIndex > 0 Then strSQL &= " And isnull(pedido,0) = " & IIf(ddlTemPedido.SelectedValue = "S", "1", "0")
            If ddlAdiantamento.SelectedIndex > 0 Then strSQL &= " And isnull(Adiantamento,0) = " & IIf(ddlAdiantamento.SelectedValue = "S", "1", "0")
        End If

        If opcao = 1 Then
            strSQL &= " And len(conta_id) = 7 "
            ds = Banco.ConsultaDataSet(strSQL, "PlanoDeContas")
            GridGruposDeContas.DataSource = ds
            GridGruposDeContas.DataBind()
        ElseIf opcao = 2 Then
            strSQL &= "   And conta_id LIKE '" & txtConta.Text & "%'"
            strSQL &= "    and len(Conta_Id) = 9 Order By Conta_Id"

            ds = Banco.ConsultaDataSet(strSQL, "PlanoDeContas")

            If ds Is Nothing Then
                MsgBox(Me.Page, "Grupo de Contas nao Tem Cliente e não tem contas relacionadas.")
            ElseIf ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Grupo de Contas nao Tem Cliente e não tem contas relacionadas.")
            Else
                GridContasAnaliticas.Parent.Visible = True
                GridContasAnaliticas.DataSource = ds
                GridContasAnaliticas.DataBind()
            End If
        ElseIf opcao = 3 Then
            strSQL &= "   And conta_id LIKE '" & GridGruposDeContas.SelectedRow.Cells(1).Text & "%'"
            strSQL &= "    and len(Conta_Id) = 9 Order By Conta_Id"
            ds = Banco.ConsultaDataSet(strSQL, "PlanoDeContas")

            Dim objConta As PlanoDeConta = New [Lib].Negocio.PlanoDeConta("", 0, GridGruposDeContas.SelectedRow.Cells(1).Text)

            If (ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0) AndAlso Not objConta.TemCliente Then
                MsgBox(Me.Page, "Grupo de Contas nao Tem Cliente e não tem contas relacionadas.")
            ElseIf ds.Tables(0).Rows.Count = 0 AndAlso objConta.TemCliente Then
                Selecionar(GridGruposDeContas.SelectedRow.Cells(1).Text)
            Else
                GridContasAnaliticas.Parent.Visible = True
                GridContasAnaliticas.DataSource = ds
                GridContasAnaliticas.DataBind()
            End If
        ElseIf opcao = 4 Then
            strSQL &= " AND EXISTS(SELECT 1 FROM encargosplanodecontas AS e WHERE e.conta_id = PlanoDeContas.conta_Id) "
            If pReceberPagar.Equals("P") Then
                strSQL &= " AND LEN(isnull(pagar,'')) > 0 "
            ElseIf pReceberPagar.Equals("R") Then
                strSQL &= " AND LEN(isnull(receber,'')) > 0 "
            End If
            strSQL &= " ORDER BY PlanoDeContas.conta_id"
            ds = Banco.ConsultaDataSet(strSQL, "PlanoDeContas")
            GridGruposDeContas.Parent.Visible = False
            GridContasAnaliticas.Parent.Visible = True
            GridContasAnaliticas.DataSource = ds
            GridContasAnaliticas.DataBind()
        End If
    End Sub

    Protected Sub GridGruposDeContas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridGruposDeContas.SelectedIndexChanged
        Try
            CargaGruposDeContas(3)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridContasAnaliticas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridContasAnaliticas.SelectedIndexChanged
        Try
            strSQL = "SELECT Conta_Id, Titulo, Cliente, TipoDeCliente, Produto, CentroDeCusto  " & _
                             "  FROM PlanoDeContas " & _
                             " Where conta_id = '" & GridContasAnaliticas.SelectedRow.Cells(1).Text & "'"

            For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "PlanoDeContas").Tables(0).Rows
                Conta = Dr("Conta_Id")
                Titulo = Dr("Titulo")
                Cliente = Dr("Cliente")
                TipoDeCliente = Dr("TipoDeCliente")
                Produto = Dr("Produto")
                CentroDeCusto = Dr("CentroDeCusto")
            Next
            Selecionar(Conta)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim btnAdd As Button = CType(sender, Button)
            Dim row As GridViewRow = CType(btnAdd.NamingContainer, GridViewRow)
            CargaGruposDeContas(3)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub CarregarContaAdiantamentoSubOperacao(ByVal Conta As String)
        Dim ds As New DataSet
        strSQL = " SELECT Conta_Id , Titulo, Cliente, TipoDeCliente, Produto, CentroDeCusto" & _
                 " FROM PlanoDeContas " & _
                 " Where Conta_Id =  '" & Conta & "'"

        ds = Banco.ConsultaDataSet(strSQL, "PlanoDeContas")
        GridContasAnaliticas.DataSource = ds
        GridContasAnaliticas.DataBind()

        pnlParametros.Visible = False
        GridGruposDeContas.Parent.Visible = False
        GridContasAnaliticas.Parent.Visible = True
        btnFechar.Parent.Visible = True
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            GridContasAnaliticas.DataSource = Nothing
            GridContasAnaliticas.DataBind()
            GridContasAnaliticas.Parent.Visible = False

            GridGruposDeContas.DataSource = Nothing
            GridGruposDeContas.DataBind()
            CargaGruposDeContas(IIf(Not ddlTemEncargoFinanceiro.Enabled, 4, 1), IIf(ViewState("pReceberPagar") IsNot Nothing, ViewState("pReceberPagar"), String.Empty))
            GridGruposDeContas.Parent.Visible = True

            If GridGruposDeContas.Rows.Count = 0 Then
                GridGruposDeContas.Parent.Visible = False
                GridContasAnaliticas.Parent.Visible = False
                CargaGruposDeContas(IIf(Not ddlTemEncargoFinanceiro.Enabled, 4, 2), IIf(ViewState("pReceberPagar") IsNot Nothing, ViewState("pReceberPagar"), String.Empty))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultarSub_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            GridGruposDeContas.Parent.Visible = False
            GridContasAnaliticas.Parent.Visible = False
            CargaGruposDeContas(2)
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

    Protected Sub lnkSair_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkSair.Click
        Popup.CloseDialog(Me.Page, "divConsultaPlanoDeContas")
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaPlanoDeContas")
    End Sub

End Class