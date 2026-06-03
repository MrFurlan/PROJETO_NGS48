Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Etapas
    Inherits BasePage

    Dim Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Etapas", "ACESSAR") Then
                BindGridView()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub grd_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grd.SelectedIndexChanged
        Dim ds As DataSet
        Dim pCodigo As String = grd.SelectedRow.Cells(1).Text()
        Dim sql As String = "Select Etapa_Id, Descricao From Etapas where Etapa_Id = " & pCodigo
        ds = Banco.ConsultaDataSet(sql, "TblEtapas")
        If ds.Tables(0).Rows.Count > 0 Then
            txtCodigo.Text = ds.Tables(0).Rows(0)("Etapa_Id")
            txtDescricao.Text = ds.Tables(0).Rows(0)("Descricao")
        End If
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = True
        lnkExcluir.Parent.Visible = True
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Sub BindGridView()
        Sql = "Select Etapa_Id, Descricao From Etapas order by Etapa_Id"
        grd.DataSource = Banco.ConsultaDataSet(Sql, "Etapas")
        grd.DataBind()
    End Sub

    Private Function Validar() As Boolean
        Dim aux = True
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            aux = False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            aux = False
        End If
        Return aux
    End Function

    Private Sub Incluir(ByVal Codigo As String, ByVal Descricao As String)
        If Funcoes.VerificaPermissao("Etapas", "GRAVAR") Then
            Dim SqlArray As New ArrayList
            Sql = "INSERT INTO Etapas(Etapa_Id, Descricao) Values(" & Codigo & ",'" & Descricao & "')"
            SqlArray.Add(Sql)
            If Not Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
            Limpar()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If
    End Sub

    Private Sub Alterar(ByVal Codigo As String, ByVal Descricao As String)
        If Funcoes.VerificaPermissao("Etapas", "ALTERAR") Then
            Dim SqlArray As New ArrayList
            Sql = "Update Etapas Set Descricao = '" & Descricao & "' Where Etapa_Id = " & Codigo & ""
            SqlArray.Add(Sql)
            If Not Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
            Limpar()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub Excluir(ByVal Codigo As String)
        If Funcoes.VerificaPermissao("Etapas", "EXCLUIR") Then
            Dim SqlArray As New ArrayList
            Sql = "Delete From Etapas Where (Etapa_Id = " & Codigo & ")"
            SqlArray.Add(Sql)
            If Not Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
            Limpar()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir regirsto.")
        End If
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("Etapas", "GRAVAR") Then
            If Not Validar() Then
                MsgBox(Me.Page, "É necessário preencher os campos obrigatórios!")
                Exit Sub
            End If
            Incluir(txtCodigo.Text.Trim(), txtDescricao.Text.Trim())
            BindGridView()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        If Funcoes.VerificaPermissao("Etapas", "ALTERAR") Then
            If Not Validar() Then
                MsgBox(Me.Page, "É necessário preencher os campos obrigatórios!")
                Exit Sub
            End If
            Alterar(txtCodigo.Text.Trim(), txtDescricao.Text.Trim())
            BindGridView()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        If Funcoes.VerificaPermissao("Etapas", "EXCLUIR") Then
            If Not Validar() Then
                MsgBox(Me.Page, "É necessário preencher os campos obrigatórios!")
                Exit Sub
            End If
            Excluir(txtCodigo.Text.Trim())
            BindGridView()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir regirsto.")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("Etapas", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Etapa_Id AS varchar))) + CAST(Etapa_Id AS varchar) as Codigo, Descricao From Etapas Order by Etapa_Id"
                Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Etapas.")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")

                Funcoes.BindReport(Me.Page, ds, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Etapas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class