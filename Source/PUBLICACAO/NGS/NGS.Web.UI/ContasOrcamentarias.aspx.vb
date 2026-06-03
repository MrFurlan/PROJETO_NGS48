Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ContasOrcamentarias
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim ds As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ContasOrcamentaria", "ACESSAR") Then
                    CarregarContasOrcamentarias()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarContasOrcamentarias()
        If Globais.GPermiteLeitura = "S" Then
            Sql = "  SELECT Codigo_id as Codigo, Descricao " & vbCrLf & _
                  " FROM ContasOrcamentarias " & vbCrLf & _
                  " ORDER BY Codigo_id" & vbCrLf
            GridContasOrcamentarias.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridContasOrcamentarias.DataBind()
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

    Protected Sub GridContasOrcamentarias_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridContasOrcamentarias.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridContasOrcamentarias.SelectedRow.Cells(2).Text()
            txtCodigo.Enabled = False
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ContasOrcamentarias", "GRAVAR") Then
                Sql = "INSERT Into ContasOrcamentarias(Codigo_id, Descricao) " & vbCrLf & _
                      " Values('" & txtCodigo.Text & "' " & vbCrLf & _
                      ",'" & UCase(txtDescricao.Text) & "')" & vbCrLf
                SqlArray.Add(Sql)
                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarContasOrcamentarias()
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
            If Funcoes.VerificaPermissao("ContasOrcamentarias", "ALTERAR") Then
                Sql = "UPDATE ContasOrcamentarias" & vbCrLf & _
                      " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                      " WHERE Codigo_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)
                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarContasOrcamentarias()
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
            If Funcoes.VerificaPermissao("ContasOrcamentarias", "EXCLUIR") Then
                Sql = "DELETE FROM ContasOrcamentarias" & vbCrLf & _
                      " WHERE Codigo_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)
                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarContasOrcamentarias()
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
            If Funcoes.VerificaPermissao("ContasOrcamentarias", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Codigo_Id AS varchar))) + CAST(Codigo_Id AS varchar) as Codigo, Descricao From ContasOrcamentarias Order by Codigo_Id"

                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatórios De Contas Orçamentarias.")
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
            Funcoes.Ajuda(Me.Page, "ContasOrcamentaria")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class