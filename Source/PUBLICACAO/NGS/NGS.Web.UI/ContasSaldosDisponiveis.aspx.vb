Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ContasSaldosDisponiveis
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ContasSaldosDisponiveis", "ACESSAR") Then
                Limpar()
                CarregarGrid()
                HID.Value = Guid.NewGuid().ToString
                ucConsultaEmpresas.SetarHID(HID.Value)
                ucConsultaPlanoDeContas.SetarHID(HID.Value)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objEmpresaCSD" & HID.Value) Is Nothing Then
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objEmpresaCSD" & HID.Value), [Lib].Negocio.Cliente))
            txtEmpresa.Text = itemEmpresa.Text
            txtCodigoEmpresa.Value = itemEmpresa.Value
            Session.Remove("objEmpresaCSD" & HID.Value)
        ElseIf Session("objContaINI" & HID.Value) IsNot Nothing Then
            If CType(Session("objContaINI" & HID.Value), [Lib].Negocio.PlanoDeConta).TemCliente Then
                MsgBox(Me.Page, "Conta com cliente não pode ser usada!")
            Else
                txtContaInicial.Text = CType(Session("objContaINI" & HID.Value), [Lib].Negocio.PlanoDeConta).Conta
                txtNomeContaInicial.Text = CType(Session("objContaINI" & HID.Value), [Lib].Negocio.PlanoDeConta).Titulo
            End If
            Session.Remove("objContaINI" & HID.Value)
        ElseIf Not Session("objContaFIM" & HID.Value) Is Nothing Then
            If CType(Session("objContaFIM" & HID.Value), [Lib].Negocio.PlanoDeConta).TemCliente Then
                MsgBox(Me.Page, "Conta com cliente não pode ser usada!")
            Else
                txtContaFinal.Text = CType(Session("objContaFIM" & HID.Value), [Lib].Negocio.PlanoDeConta).Conta
                txtNomeContaFinal.Text = CType(Session("objContaFIM" & HID.Value), [Lib].Negocio.PlanoDeConta).Titulo
            End If
            Session.Remove("objContaFIM" & HID.Value)
        End If
    End Sub

    Private Sub Limpar()
        txtEmpresa.Text = ""
        txtCodigoEmpresa.Value = ""
        txtContaInicial.Text = ""
        txtNomeContaInicial.Text = ""
        txtContaFinal.Text = ""
        txtNomeContaFinal.Text = ""
        txtDescricao.Text = ""
        lnkNovo.Enabled = True
        lnkExcluir.Enabled = False
        SqlArray.Clear()
    End Sub

    Private Sub Desabilitar()
        imgEmpresa.Enabled = False
        imgContaInicial.Enabled = False
        imgContaFinal.Enabled = False
        lnkNovo.Enabled = False
        lnkLimpar.Enabled = False
        lnkExcluir.Enabled = False
        MsgBox(Me.Page, "Usuário sem permissão para essa página!")
    End Sub

    Private Sub CarregarGrid()
        Dim ds As New DataSet

        Sql = "SELECT     Empresa_Id, EndEmpresa_Id, ContaInicial_Id, ContaFinal_Id, Descricao "
        Sql &= "FROM         ContasSaldosDisponiveis "
        Sql &= "ORDER BY Empresa_Id"

        ds = Banco.ConsultaDataSet(Sql, "ContasSaldosDisponiveis")

        gridContasSaldosDisponiveis.DataSource = ds
        gridContasSaldosDisponiveis.DataBind()
    End Sub

    Private Sub BuscarConta(ByVal Tipo As String)
        Session("ssCampo") = "ContasSaldosDisponiveis"
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, Tipo & HID.Value)
    End Sub

    Private Sub Excluir()
        If Funcoes.VerificaPermissao("ContasSaldosDisponiveis", "EXCLUIR") Then
            Dim Empresa() As String = txtCodigoEmpresa.Value.Split("-")
            Sql = "DELETE FROM ContasSaldosDisponiveis " & _
                  " WHERE  Empresa_Id = '" & Empresa(0) & "' AND EndEmpresa_Id = " & Empresa(1) & _
                  "   AND  ContaInicial_Id = '" & txtContaInicial.Text & "' AND ContaFinal_Id = '" & txtContaFinal.Text & "'"
            SqlArray.Add(Sql)
            If Banco.GravaBanco(SqlArray) Then
                Limpar()
                CarregarGrid()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
        End If
    End Sub

    Protected Sub imgEmpresa_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Session("ssCampo") = "Livre"
        ucConsultaEmpresas.Limpar()
        Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaCSD" & HID.Value)
    End Sub

    Protected Sub imgContaInicial_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If String.IsNullOrWhiteSpace(txtCodigoEmpresa.Value) Then
            MsgBox(Me.Page, "Empresa não foi selecionanda!")
        Else
            BuscarConta("objContaINI")
        End If
    End Sub

    Protected Sub imgContaFinal_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If String.IsNullOrWhiteSpace(txtCodigoEmpresa.Value) Then
            MsgBox(Me.Page, "Empresa não foi selecionanda!")
        ElseIf String.IsNullOrWhiteSpace(txtContaInicial.Text) Then
            MsgBox(Me.Page, "Conta inicial não foi selecionada!")
        Else
            BuscarConta("objContaFIM")
        End If
    End Sub

    Protected Sub gridContasSaldosDisponiveis_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim objCliente As New [Lib].Negocio.Cliente(gridContasSaldosDisponiveis.SelectedRow.Cells(1).Text(), gridContasSaldosDisponiveis.SelectedRow.Cells(2).Text())
        Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objCliente)
        txtEmpresa.Text = itemEmpresa.Text
        txtCodigoEmpresa.Value = itemEmpresa.Value

        Dim objContaInicial As New [Lib].Negocio.PlanoDeConta("", 0, gridContasSaldosDisponiveis.SelectedRow.Cells(3).Text)
        txtContaInicial.Text = objContaInicial.Conta
        txtNomeContaInicial.Text = objContaInicial.Titulo

        Dim objContaFinal As New [Lib].Negocio.PlanoDeConta("", 0, gridContasSaldosDisponiveis.SelectedRow.Cells(4).Text)
        txtContaFinal.Text = objContaFinal.Conta
        txtNomeContaFinal.Text = objContaFinal.Titulo

        txtDescricao.Text = Server.HtmlDecode(gridContasSaldosDisponiveis.SelectedRow.Cells(5).Text)
        lnkNovo.Enabled = False
        lnkExcluir.Enabled = True
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("ContasSaldosDisponiveis", "GRAVAR") Then
            Dim Empresa() As String = txtCodigoEmpresa.Value.Split("-")
            Sql = "INSERT INTO ContasSaldosDisponiveis " & _
                  "                      (Empresa_Id, EndEmpresa_Id, ContaInicial_Id, ContaFinal_Id, Descricao)" & _
                  "VALUES     ('" & Empresa(0) & "', " & Empresa(1) & ", '" & txtContaInicial.Text & "', '" & txtContaFinal.Text & "', " & _
                  "            '" & RTrim(Funcoes.EliminarCaracteresEspeciais(txtDescricao.Text)).ToUpper & "')"
            SqlArray.Add(Sql)
            If Banco.GravaBanco(SqlArray) Then
                Limpar()
                CarregarGrid()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            Excluir()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ContasSaldosDisponiveis")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class