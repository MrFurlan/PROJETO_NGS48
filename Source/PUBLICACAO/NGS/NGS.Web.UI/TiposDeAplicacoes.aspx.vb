Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class TiposDeAplicacoes
    Inherits BasePage

    Dim DS As DataSet
    Dim objTipoDeAplicacao As TipoDeAplicacao

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("TiposDeAplicacoes", "ACESSAR") Then
                CarregarTiposDeAplicacoes()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página;", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarTiposDeAplicacoes()
        If Funcoes.VerificaPermissao("TiposDeAplicacoes", "LEITURA") Then
            Dim ListTipoDeAplicacoes = New ListTipoDeAplicacao()
            GridTiposDeAplicacoes.DataSource = ListTipoDeAplicacoes
            GridTiposDeAplicacoes.DataBind()
        End If
    End Sub

    Protected Sub GridTiposDeAplicacoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            objTipoDeAplicacao = New TipoDeAplicacao(GridTiposDeAplicacoes.SelectedRow.Cells(1).Text())
            txtCodigo.Text = objTipoDeAplicacao.Codigo
            txtDescricao.Text = objTipoDeAplicacao.Descricao
            txtCodigo.Enabled = False
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function ValidarCampos() As Boolean
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

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("TiposDeAplicacoes", "GRAVAR") Then
                If ValidarCampos() Then
                    objTipoDeAplicacao = New TipoDeAplicacao()
                    objTipoDeAplicacao.IUD = "I"
                    objTipoDeAplicacao.Codigo = txtCodigo.Text
                    objTipoDeAplicacao.Descricao = txtDescricao.Text

                    If objTipoDeAplicacao.Salvar Then
                        Limpar()
                        CarregarTiposDeAplicacoes()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("TiposDeAplicacoes", "ALTERAR") Then
                If ValidarCampos() Then
                    objTipoDeAplicacao = New TipoDeAplicacao(txtCodigo.Text)
                    objTipoDeAplicacao.IUD = "U"
                    objTipoDeAplicacao.Descricao = txtDescricao.Text

                    If objTipoDeAplicacao.Salvar Then
                        Limpar()
                        CarregarTiposDeAplicacoes()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("TiposDeAplicacoes", "EXCLUIR") Then
                objTipoDeAplicacao = New TipoDeAplicacao(txtCodigo.Text)
                objTipoDeAplicacao.IUD = "D"
                If objTipoDeAplicacao.Salvar Then
                    Limpar()
                    CarregarTiposDeAplicacoes()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
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
            If Funcoes.VerificaPermissao("TiposDeAplicacoes", "RELATORIO") Then
                Dim Sql = "Select REPLICATE('0', 3 - LEN(CAST(Codigo_Id AS varchar))) + CAST(Codigo_Id AS varchar) as Codigo, Descricao From TiposDeAplicacoes Order by Codigo_Id"

                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Tabela de Tipos De Aplicações")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
               
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
            Funcoes.Ajuda(Me.Page, "TiposDeAplicacoes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class