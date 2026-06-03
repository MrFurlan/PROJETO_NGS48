Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Permissoes
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Permissoes", "ACESSAR") Then
                    CarregarPermissoes()
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

    Private Sub CarregarPermissoes()
        If Funcoes.VerificaPermissao("Permissoes", "LEITURA") Then
            Sql = "  SELECT Permissao_Id As Permissao, Descricao " & vbCrLf & _
                            " FROM Permissoes " & vbCrLf & _
                            " ORDER BY Permissao_Id" & vbCrLf

            GridPermissao.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridPermissao.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtPermissao.Text = ""
        txtDescricao.Text = ""
        txtPermissao.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridPermissoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtPermissao.Text = GridPermissao.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridPermissao.SelectedRow.Cells(2).Text()
            txtPermissao.Enabled = False
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
            If Funcoes.VerificaPermissao("Permissoes", "GRAVAR") Then
                Sql = "INSERT Into Permissoes (Permissao_id, Descricao) " & vbCrLf & _
                      " Values('" & txtPermissao.Text & "' " & vbCrLf & _
                      ",'" & UCase(txtDescricao.Text) & "')" & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarPermissoes()
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
            If Funcoes.VerificaPermissao("Permissoes", "ALTERAR") Then
                Sql = "UPDATE Permissoes" & vbCrLf & _
                      " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                      " WHERE Permissao_Id = '" & txtPermissao.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarPermissoes()
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
            If Funcoes.VerificaPermissao("Permissoes", "EXCLUIR") Then
                Sql = "DELETE FROM Permissoes" & vbCrLf & _
                      " WHERE Permissao_Id = '" & txtPermissao.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarPermissoes()
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
            If Funcoes.VerificaPermissao("Permissoes", "RELATORIO") Then
                Sql = " Select Permissao_Id as Codigo, Descricao From Permissoes Order by Permissao_Id "
                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Permissões.")
                parameters.Add("Codigo", "Permissão")
                parameters.Add("Descricao", "Descricao")
                Dim objEmpresa As Cliente = New Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
                parameters.Add("Empresa", objEmpresa.Nome.Trim())
                parameters.Add("CidadeCnpj", objEmpresa.Cidade.Trim() & "/" & objEmpresa.Estado.Codigo.Trim() & vbCrLf & "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))

                Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatorio.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Permissoes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class