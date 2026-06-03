Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class GruposDeAtivos
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Patrimonio)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("GruposDeAtivos", "ACESSAR") Then
                CarregarGruposDeAtivos()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarGruposDeAtivos()
        If Funcoes.VerificaPermissao("GruposDeAtivos", "LEITURA") Then
            Sql = "  SELECT Grupo_Id as Grupo, Descricao, PercentualDepreciacao as Depreciacao" & vbCrLf & _
                            " FROM GruposDeAtivos " & vbCrLf & _
                            " ORDER BY Grupo_Id" & vbCrLf

            GridGruposDeAtivos.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridGruposDeAtivos.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtGrupo.Text = ""
        txtDescricao.Text = ""
        txtDepreciacao.Text = ""
        txtGrupo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(txtGrupo.Text) Then
            param &= "Grupo: " & txtGrupo.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            param &= "Descrição: " & txtDescricao.Text
        End If

        Return param
    End Function

    Protected Sub GridGruposDeAtivo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtGrupo.Text = Server.HtmlDecode(GridGruposDeAtivos.SelectedRow.Cells(1).Text())
            txtDescricao.Text = Server.HtmlDecode(GridGruposDeAtivos.SelectedRow.Cells(2).Text())
            txtDepreciacao.Text = Server.HtmlDecode(GridGruposDeAtivos.SelectedRow.Cells(3).Text())
            txtGrupo.Enabled = False
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
            If Funcoes.VerificaPermissao("GruposDeAtivos", "GRAVAR") Then
                Sql = "INSERT Into GruposDeAtivos (Grupo_Id, Descricao, PercentualDepreciacao) " & vbCrLf & _
                      " Values('" & txtGrupo.Text & "' " & vbCrLf & _
                      ", '" & UCase(txtDescricao.Text) & "' " & vbCrLf & _
                      ", " & Replace(txtDepreciacao.Text, ",", ".") & ")" & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                Else
                    Limpar()
                    CarregarGruposDeAtivos()
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
            If Funcoes.VerificaPermissao("GruposDeAtivos", "ALTERAR") Then
                Sql = "UPDATE GruposDeAtivos" & vbCrLf & _
                      " Set Descricao = '" & txtDescricao.Text & "', " & vbCrLf & _
                      " PercentualDepreciacao = " & Replace(Replace(txtDepreciacao.Text, ".", ""), ",", ".") & vbCrLf & _
                      " WHERE Grupo_Id = '" & txtGrupo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarGruposDeAtivos()
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
            If Funcoes.VerificaPermissao("GruposDeAtivos", "EXCLUIR") Then
                Sql = "DELETE FROM GruposDeAtivos" & vbCrLf & _
                      " WHERE Grupo_Id = '" & txtGrupo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarGruposDeAtivos()
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
            If Funcoes.VerificaPermissao("GruposDeAtivos", "RELATORIO") Then

                Sql = "  SELECT Grupo_Id as Codigo, Descricao, PercentualDepreciacao as Depreciacao" & vbCrLf & _
                      " FROM GruposDeAtivos " & vbCrLf & _
                      " ORDER BY Grupo_Id" & vbCrLf
                DS = Banco.ConsultaDataSet(Sql, "GruposDeAtivos")

                Funcoes.BindReport(Me.Page, DS, "Cr_GruposDeAtivos", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "GruposDeAtivos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class