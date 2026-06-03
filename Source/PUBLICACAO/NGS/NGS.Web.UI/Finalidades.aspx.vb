Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Finalidades
    Inherits BasePage

    Private Sql As String
    Private SqlArray As New ArrayList
    Private DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Finalidades", "ACESSAR") AndAlso Request.Url.Host.ToUpper.Contains("LOCALHOST") Then
                CarregarFinalidades()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarFinalidades()
        If Funcoes.VerificaPermissao("Finalidades", "LEITURA") Then
            Sql = "  SELECT Finalidade_id as Codigo, Descricao " & vbCrLf & _
                            " FROM Finalidades " & vbCrLf & _
                            " ORDER BY Finalidade_id" & vbCrLf

            GridFinalidades.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridFinalidades.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        'apenas via banco - furlan - 27/10/2014
        'lnkNovo.Parent.Visible = True
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridFinalidades_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridFinalidades.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridFinalidades.SelectedRow.Cells(2).Text()
            txtCodigo.Enabled = False
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            'apenas via banco - furlan - 27/10/2014
            'lnkAtualizar.Parent.Visible = True
            'lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Finalidades", "GRAVAR") Then
                Sql = "INSERT Into Finalidades(Finalidade_id, Descricao) " & vbCrLf & _
                     " Values('" & txtCodigo.Text & "' " & vbCrLf & _
                     ",'" & UCase(txtDescricao.Text) & "')" & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) Then
                    Limpar()
                    CarregarFinalidades()
                Else
                    MsgBox(Me.Page, Session("ssMessage"))
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
            If Funcoes.VerificaPermissao("Finalidades", "ALTERAR") Then
                Sql = "UPDATE Finalidades" & vbCrLf & _
                      " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                      " WHERE Finalidade_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarFinalidades()
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
            If Funcoes.VerificaPermissao("Finalidades", "EXCLUIR") Then
                Sql = "DELETE FROM Finalidades" & vbCrLf & _
                      " WHERE Finalidade_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarFinalidades()
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
            If Funcoes.VerificaPermissao("Finalidades", "RELATORIO") Then

                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Finalidade_Id AS varchar))) + CAST(Finalidade_Id AS varchar) as Codigo, Descricao From Finalidades Order by Finalidade_Id"
                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Finalidades.")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
                
                Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir o relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Finalidades")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class