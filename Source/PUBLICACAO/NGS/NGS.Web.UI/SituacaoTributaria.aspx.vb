Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class SituacaoTributaria
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Dim objSituacaoTributaria As [Lib].Negocio.SituacaoTributaria

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("SituacaoTributaria", "ACESSAR") Then
                CarregarSituacaoTributaria()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarSituacaoTributaria()
        If Funcoes.VerificaPermissao("SituacaoTributaria", "LEITURA") Then
            GridSituacaoTributaria.DataSource = New ListSituacaoTributaria(True)
            GridSituacaoTributaria.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridSituacaoTributarias_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            objSituacaoTributaria = New [Lib].Negocio.SituacaoTributaria(GridSituacaoTributaria.SelectedRow.Cells(1).Text())
            txtCodigo.Text = objSituacaoTributaria.Codigo
            txtDescricao.Text = objSituacaoTributaria.Descricao
            txtCodigo.Enabled = False
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            MsgBox(Me.Page, "Código não foi informado")
            txtCodigo.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição não foi informada")
            txtDescricao.Focus()
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributaria", "GRAVAR") Then
                If ValidarCampos() Then
                    objSituacaoTributaria = New [Lib].Negocio.SituacaoTributaria
                    objSituacaoTributaria.IUD = "I"
                    objSituacaoTributaria.Codigo = txtCodigo.Text
                    objSituacaoTributaria.Descricao = txtDescricao.Text

                    If objSituacaoTributaria.Salvar Then
                        Limpar()
                        CarregarSituacaoTributaria()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributaria", "ALTERAR") Then
                If ValidarCampos() Then
                    objSituacaoTributaria = New [Lib].Negocio.SituacaoTributaria(txtCodigo.Text)
                    objSituacaoTributaria.IUD = "U"
                    objSituacaoTributaria.Descricao = txtDescricao.Text

                    If objSituacaoTributaria.Salvar Then
                        Limpar()
                        CarregarSituacaoTributaria()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para alterar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributaria", "EXCLUIR") Then
                If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
                    objSituacaoTributaria = New [Lib].Negocio.SituacaoTributaria(txtCodigo.Text)
                    objSituacaoTributaria.IUD = "D"

                    If objSituacaoTributaria.Salvar Then
                        Limpar()
                        CarregarSituacaoTributaria()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para excluir registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributaria", "RELATORIO") Then

                Dim Titulo As String = "Tabela de Situações Tributárias"
                Dim Codigo As String = "Código"
                Dim Descricao As String = "Descrição"

                Titulo = "Tabela de Situações Tributárias"
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(SituacaoTributaria_Id AS varchar))) + CAST(SituacaoTributaria_Id AS varchar) as Codigo, Descricao From SituacaoTributaria Order by SituacaoTributaria_Id"

                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", Titulo)
                parameters.Add("Codigo", Codigo)
                parameters.Add("Descricao", Descricao)
                
                Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuario sem permissao para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "SituacaoTributaria")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class