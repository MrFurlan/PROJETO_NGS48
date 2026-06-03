Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Historicos
    Inherits BasePage

    Dim SqlArray As New ArrayList
    Dim ds As DataSet
    Dim Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Historicos", "ACESSAR") Then
                CarregarHistoricosPadroes()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarHistoricosPadroes()
        Sql = "  SELECT Historico_id as Codigo, Descricao " & vbCrLf & _
                        " FROM Historicos " & vbCrLf & _
                        " ORDER BY Historico_id" & vbCrLf
        GridHistoricosPadroes.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
        GridHistoricosPadroes.DataBind()
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridHistoricosPadroes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridHistoricosPadroes.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridHistoricosPadroes.SelectedRow.Cells(2).Text()
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
            If Funcoes.VerificaPermissao("Historicos", "GRAVAR") Then
                Sql = "INSERT Into Historicos(Historico_id, Descricao) " & vbCrLf & _
                      " Values('" & txtCodigo.Text & "' " & vbCrLf & _
                      ",'" & UCase(txtDescricao.Text) & "')" & vbCrLf
                SqlArray.Add(Sql)
                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarHistoricosPadroes()
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
            If Funcoes.VerificaPermissao("Historicos", "ALTERAR") Then
                Sql = "UPDATE Historicos" & vbCrLf & _
                       " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                       " WHERE Historico_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarHistoricosPadroes()
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
            If Funcoes.VerificaPermissao("Historicos", "EXCLUIR") Then
                Sql = "DELETE FROM Historicos" & vbCrLf & _
                      " WHERE Historico_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)
                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarHistoricosPadroes()
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
            If Funcoes.VerificaPermissao("Historicos", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Historico_Id AS varchar))) + CAST(Historico_Id AS varchar) as Codigo, Descricao From Historicos Order by Historico_Id"
                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Históricos Padrões.")
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
            Funcoes.Ajuda(Me.Page, "Historicos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class