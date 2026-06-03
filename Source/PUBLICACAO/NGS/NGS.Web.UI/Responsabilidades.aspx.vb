Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Responsabilidades
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Responsabilidades", "ACESSAR") Then
                CarregarResponsabilidades()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarResponsabilidades()
        If Globais.GPermiteLeitura = "S" Then
            Sql = "  SELECT Codigo_id as Codigo, Descricao " & _
                            " FROM Responsabilidades " & _
                            " ORDER BY Codigo_id"
            GridResponsabilidades.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridResponsabilidades.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        Try
            txtCodigo.Text = ""
            txtDescricao.Text = ""
            txtCodigo.Enabled = True
            lnkNovo.Parent.Visible = True
            lnkAtualizar.Parent.Visible = False
            lnkExcluir.Parent.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridResponsabilidades_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Limpar()
        txtCodigo.Text = GridResponsabilidades.SelectedRow.Cells(1).Text()
        txtDescricao.Text = GridResponsabilidades.SelectedRow.Cells(2).Text()
        txtCodigo.Enabled = False
        txtDescricao.Focus()
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = True
        lnkExcluir.Parent.Visible = True
    End Sub

    Private Sub Relatorio()
        Try
            If Funcoes.VerificaPermissao("Responsabilidades", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Codigo_Id AS varchar))) + CAST(Codigo_Id AS varchar) as Codigo, Descricao From Responsabilidades Order by Codigo_Id"

                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Tabela de Responsabilidades")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")

                Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Responsabilidades", "GRAVAR") Then
                Sql = "INSERT Into Responsabilidades(Codigo_id, Descricao) "
                Sql &= " Values('" & txtCodigo.Text & "' "
                Sql &= ",'" & UCase(txtDescricao.Text) & "')"
                SqlArray.Add(Sql)
                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarResponsabilidades()
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
            If Funcoes.VerificaPermissao("Responsabilidades", "ATUALIZAR") Then
                Sql = "UPDATE Responsabilidades"
                Sql &= " Set Descricao = '" & txtDescricao.Text & "' "
                Sql &= " WHERE Codigo_Id = '" & txtCodigo.Text & "' "
                SqlArray.Add(Sql)
                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarResponsabilidades()
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
            If Funcoes.VerificaPermissao("Responsabilidades", "EXCLUIR") Then
                Sql = "DELETE FROM Responsabilidades"
                Sql &= " WHERE Codigo_Id = '" & txtCodigo.Text & "' "
                SqlArray.Add(Sql)
                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarResponsabilidades()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Relatorio()
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Responsabilidades")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class