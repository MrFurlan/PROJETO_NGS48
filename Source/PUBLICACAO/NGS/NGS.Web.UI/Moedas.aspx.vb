Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Moedas
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Moedas", "ACESSAR") Then
                CarregarMoedas()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarMoedas()
        If Funcoes.VerificaPermissao("Moedas", "LEITURA") Then
            Sql = "  SELECT Moeda_id as Codigo, Descricao " & vbCrLf & _
                            " FROM Moedas " & vbCrLf & _
                            " ORDER BY Moeda_id" & vbCrLf
            GridMoedas.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")

            GridMoedas.DataBind()
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

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            MsgBox(Me.Page, "Código não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição não foi informada!", eTitulo.Info)
            Return False
        End If

        Return True
    End Function

    Protected Sub GridMoedas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridMoedas.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridMoedas.SelectedRow.Cells(2).Text()
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
            If Funcoes.VerificaPermissao("Moedas", "GRAVAR") Then
                If ValidarCampos() Then
                    Sql = "INSERT Into Moedas(Moeda_id, Descricao) " & vbCrLf & _
                                          " Values('" & txtCodigo.Text & "' " & vbCrLf & _
                                          ",'" & UCase(txtDescricao.Text) & "')" & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) Then
                        Limpar()
                        CarregarMoedas()
                        MsgBox(Me.Page, "Registro salvo com Sucesso.", eTitulo.Sucess)
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
            If Funcoes.VerificaPermissao("Moedas", "ALTERAR") Then
                If ValidarCampos() Then
                    Sql = "UPDATE Moedas" & vbCrLf & _
                                          " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                                          " WHERE Moeda_Id = '" & txtCodigo.Text & "' " & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) Then
                        Limpar()
                        CarregarMoedas()
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para alterar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Moedas", "EXCLUIR") Then
                Sql = "DELETE FROM Moedas" & vbCrLf & _
                      " WHERE Moeda_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    Limpar()
                    CarregarMoedas()
                    MsgBox(Me.Page, "Registro deletado com Sucesso.", eTitulo.Sucess)
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para excluir registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("Moedas", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Moeda_Id AS varchar))) + CAST(Moeda_Id AS varchar) as Codigo, Descricao From Moedas Order by Moeda_Id"
                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Moedas.")
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

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Moedas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class
