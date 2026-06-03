Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class UnidadesDeMedida
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("UnidadesDeMedida", "ACESSAR") Then
                Limpar()
                AtualizarGrid()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub IniciarUnidade(ByVal Tipo As String)
        If ValidarCampos() Then
            Dim objUnidade As New [Lib].Negocio.UnidadeDeMedida
            objUnidade.Unidade = txtCodigo.Text
            objUnidade.Descricao = txtDescricao.Text
            objUnidade.UnidadeIndea = txtCodigoIndea.Text
            objUnidade.IUD = Tipo
            If objUnidade.Salvar Then
                Limpar()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        End If
    End Sub

    Private Function ValidarCampos() As Boolean
        If txtCodigo.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Código da unidade não foi informado.")
            Return False
        ElseIf txtDescricao.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Descrição da unidade não foi informada.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub AtualizarGrid()
        Dim Lista As New [Lib].Negocio.ListUnidadeDeMedida
        GridUnidadesDeMedidas.DataSource = Lista.ToArray
        GridUnidadesDeMedidas.DataBind()
    End Sub

    Private Sub Limpar()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        txtCodigo.Enabled = True
        txtDescricao.Enabled = True
        txtCodigoIndea.Enabled = True

        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigoIndea.Text = ""

        AtualizarGrid()

        txtCodigo.Focus()
    End Sub

    Protected Sub GridUnidadesDeMedidas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridUnidadesDeMedidas.SelectedIndexChanged
        Try
            Dim objUnidade As New [Lib].Negocio.UnidadeDeMedida(GridUnidadesDeMedidas.SelectedRow.Cells(1).Text())

            txtCodigo.Text = objUnidade.Unidade
            txtDescricao.Text = objUnidade.Descricao
            txtCodigoIndea.Text = objUnidade.UnidadeIndea
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtCodigo.Enabled = False
            txtDescricao.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("UnidadesDeMedida", "GRAVAR") Then
                IniciarUnidade("I")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("UnidadesDeMedida", "ALTERAR") Then
                IniciarUnidade("U")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("UnidadesDeMedida", "EXCLUIR") Then
                IniciarUnidade("D")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
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

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("UnidadesDeMedida", "RELATORIO") Then
                Dim ds As New DataSet
                Dim Sql As String

                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Unidade_Id AS varchar))) + CAST(Unidade_Id AS varchar) as Codigo, Descricao From UnidadeDeMedida Order by Unidade_Id"

                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Tabela de Unidades De Medidas")
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "UnidadesDeMedida")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class