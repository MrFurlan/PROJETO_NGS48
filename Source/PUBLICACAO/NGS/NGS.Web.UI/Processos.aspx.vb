Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Processos
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Processos", "ACESSAR") Then
                    Limpar()
                    CarregarProcessos()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarProcessos()
        If Funcoes.VerificaPermissao("Estados", "LEITURA") Then
            Dim ListProcesso As New [Lib].Negocio.Processos(eServidor.ServidorLocal)

            'ListProcesso tem que buscar no UOL se é PRIMEIRA VEZ
            'LISTA QUE ESTÁ NA UOL VERIFICAR SE EXISTE NO LOCAL, O QUE EXISTIR UPDATE, E O QUE NÃO EXISTIR INSERT
            'BUSCA OS PROCESSOS NO CLIENTE E PREENCHE O GRID

            GridProcessos.DataSource = ListProcesso
            GridProcessos.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtProcesso.Text = String.Empty
        txtDescricao.Text = String.Empty
        txtManual.Text = String.Empty
        txtProcesso.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkAtualizarTudo.Enabled = False
        lnkAtualizarTudo.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        If Session("ssNomeUsuario") = "KAYNÃ" Or Session("ssNomeUsuario") = "FURLAN" Then
            lnkAtualizarTudo.Enabled = True
            lnkAtualizarTudo.Parent.Visible = True
        End If
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(txtProcesso.Text) Then
            param &= "Processo: " & txtProcesso.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            param &= "Descrição: " & txtDescricao.Text
        End If

        Return param
    End Function

    Protected Sub GridProcessos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtProcesso.Text = Server.HtmlDecode(GridProcessos.SelectedRow.Cells(1).Text())
            txtDescricao.Text = Server.HtmlDecode(GridProcessos.SelectedRow.Cells(2).Text())
            txtProcesso.Enabled = False
            txtManual.Text = Server.HtmlDecode(GridProcessos.SelectedRow.Cells(3).Text)
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtDescricao.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Processos", "GRAVAR") Then
                Dim obj As [Lib].Negocio.Processo = preenheObjeto("I")

                If obj.Salvar(eServidor.ServidorUOL) AndAlso obj.Salvar(eServidor.ServidorLocal) Then
                    MsgBox(Me.Page, "Sucesso na inclusão.")
                    Limpar()
                    CarregarProcessos()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function preenheObjeto(ByVal IUD As String) As Processo
        Dim obj As New [Lib].Negocio.Processo()
        obj.IUD = IUD
        obj.Processo = txtProcesso.Text.Trim
        obj.Descricao = txtDescricao.Text.Trim
        obj.Manual = txtManual.Text.Trim
        obj.DataAtualizacao = DateTime.Now()
        Return obj
    End Function

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Processos", "ALTERAR") Then
                Dim obj As [Lib].Negocio.Processo = preenheObjeto("U")

                If obj.Salvar(eServidor.ServidorUOL) AndAlso obj.Salvar(eServidor.ServidorLocal) Then
                    MsgBox(Me.Page, "Sucesso na alteração.")
                    Limpar()
                    CarregarProcessos()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizarTudo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizarTudo.Click
        Try
            Dim ListProcesso As New [Lib].Negocio.Processos(eServidor.ServidorUOL)
            For Each objUol As Object In ListProcesso

                If objUol.Manual.Trim() = String.Empty Then
                    Continue For
                End If
                Dim obj As New [Lib].Negocio.Processo()
                obj.IUD = "U"
                obj.Processo = objUol.Processo
                obj.Descricao = objUol.Descricao
                obj.Manual = objUol.Manual
                obj.DataAtualizacao = DateTime.Now()

                If Not obj.Salvar(eServidor.ServidorLocal) Then
                    MsgBox(Me.Page, "Erro ao atualizar o processo: " & objUol.Processo)
                End If
            Next

            MsgBox(Me.Page, "Processos atualizados com sucesso.")
            Limpar()
            CarregarProcessos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Processos", "EXCLUIR") Then
                With New [Lib].Negocio.Processo()
                    .Processo = txtProcesso.Text.Trim
                    .IUD = "D"
                    If .Salvar(eServidor.ServidorLocal) Then
                        MsgBox(Me.Page, "Sucesso na exclusão.")
                        Limpar()
                        CarregarProcessos()
                    End If
                End With
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
            If Funcoes.VerificaPermissao("Processos", "RELATORIO") Then
                Dim sql As String = " Select Processo_Id, Descricao From Processos Order by Processo_Id "
                Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Processos")

                Dim param As New Dictionary(Of String, Object)
                Funcoes.BindReport(Me.Page, ds, "Cr_Processos", eExportType.PDF, param)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Processos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class