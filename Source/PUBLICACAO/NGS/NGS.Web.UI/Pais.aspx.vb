Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Pais
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim Ds As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Pais", "ACESSAR") Then
                    CarregarPais()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarPais()
        If Funcoes.VerificaPermissao("Pais", "LEITURA") Then
            Sql = "  SELECT Pais_Id as Codigo, Descricao " & vbCrLf & _
                  " FROM Pais " & vbCrLf & _
                  " ORDER BY Pais_Id" & vbCrLf
            GridPais.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridPais.DataBind()
        End If
    End Sub

    Protected Sub GridPais_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridPais.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridPais.SelectedRow.Cells(2).Text()
            txtCodigo.Enabled = False
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

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
            If Funcoes.VerificaPermissao("Pais", "GRAVAR") Then
                Sql = "INSERT Into Pais(Pais_id, Descricao) " & vbCrLf & _
                      " Values('" & txtCodigo.Text & "' " & vbCrLf & _
                      ",'" & UCase(txtDescricao.Text) & "')" & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarPais()
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
            If Funcoes.VerificaPermissao("Pais", "ALTERAR") Then
                Sql = "UPDATE Pais" & vbCrLf & _
                      " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                      " WHERE Pais_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    CarregarPais()
                    Limpar()
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
            If Funcoes.VerificaPermissao("Pais", "EXCLUIR") Then
                Sql = "DELETE FROM Pais" & vbCrLf & _
                      " WHERE Pais_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, Session("ssMessage"))
                Else
                    Limpar()
                    CarregarPais()
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
            If Funcoes.VerificaPermissao("Pais", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 4 - LEN(CAST(Pais_Id AS varchar))) + CAST(Pais_Id AS varchar) as Codigo, " & vbCrLf & _
                      "Descricao" & vbCrLf & _
                      "From Pais" & vbCrLf & _
                      "Order by Descricao" & vbCrLf
                Ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Países.")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
               
                Funcoes.BindReport(Me.Page, Ds, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Pais")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class