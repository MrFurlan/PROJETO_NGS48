Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Situacoes
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Dim objSituacao As Situacao

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Situacoes", "ACESSAR") AndAlso Request.Url.Host.ToUpper.Contains("LOCALHOST") Then
                CarregarSituacoes()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarSituacoes()
        If Funcoes.VerificaPermissao("Situacoes", "LEITURA") Then
            'Dim ListSituacoes As New ListSituacao(True)
            GridSituacoes.DataSource = New ListSituacao(True)
            GridSituacoes.DataBind()
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

    Protected Sub GridSituacoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            objSituacao = New Situacao(GridSituacoes.SelectedRow.Cells(1).Text())
            txtCodigo.Text = objSituacao.Codigo
            txtDescricao.Text = objSituacao.Descricao
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
            MsgBox(Me.Page, "Código não foi informado!")
            txtCodigo.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição não foi informada!")
            txtDescricao.Focus()
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Situacoes", "GRAVAR") Then
                If ValidarCampos() Then
                    objSituacao = New Situacao()
                    objSituacao.IUD = "I"
                    objSituacao.Codigo = txtCodigo.Text
                    objSituacao.Descricao = txtDescricao.Text

                    If objSituacao.Salvar Then
                        Limpar()
                        CarregarSituacoes()
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
            If Funcoes.VerificaPermissao("Situacoes", "ALTERAR") Then
                If ValidarCampos() Then
                    objSituacao = New Situacao(txtCodigo.Text)
                    objSituacao.IUD = "U"
                    objSituacao.Codigo = txtCodigo.Text
                    objSituacao.Descricao = txtDescricao.Text

                    If objSituacao.Salvar Then

                        Limpar()
                        CarregarSituacoes()
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
            If Funcoes.VerificaPermissao("Situacoes", "EXCLUIR") Then
                'Sql = "DELETE FROM Situacoes"
                'Sql &= " WHERE Situacao_Id = '" & txtCodigo.Text & "' "
                'SqlArray.Add(Sql)
                objSituacao = New Situacao(txtCodigo.Text)
                objSituacao.IUD = "D"
                If objSituacao.Salvar Then
                    Limpar()
                    CarregarSituacoes()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para excluir registro.")
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
            If Funcoes.VerificaPermissao("Situacoes", "RELATORIO") Then

                Dim Titulo As String = "Tabela de Situações"
                Dim Codigo As String = "Código"
                Dim Descricao As String = "Descrição"

                Titulo = "Tabela de Situações"
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Situacao_Id AS varchar))) + CAST(Situacao_Id AS varchar) as Codigo, Descricao From Situacoes Order by Situacao_Id"

                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", Titulo)
                parameters.Add("Codigo", Codigo)
                parameters.Add("Descricao", Descricao)
               
                Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuario sem permissao para tirar Relatório")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Situacoes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class