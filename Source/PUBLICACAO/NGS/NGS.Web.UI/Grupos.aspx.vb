Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Grupos
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Grupos", "ACESSAR") Then
                CarregarUsuarios()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarUsuarios()
        If Funcoes.VerificaPermissao("Grupos", "LEITURA") Then
            Sql = "  SELECT Grupo_Id as Grupo, Descricao " & vbCrLf & _
                            " FROM Grupos " & vbCrLf & _
                            " ORDER BY Grupo_Id" & vbCrLf

            GridGruposDeUsuarios.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridGruposDeUsuarios.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtGrupo.Text = ""
        txtDescricao.Text = ""
        txtGrupo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridUsuarios_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtGrupo.Text = Server.HtmlDecode(GridGruposDeUsuarios.SelectedRow.Cells(1).Text())
            txtDescricao.Text = Server.HtmlDecode(GridGruposDeUsuarios.SelectedRow.Cells(2).Text())
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
        If Funcoes.VerificaPermissao("Grupos", "GRAVAR") Then
            Sql = "INSERT Into Grupos (Grupo_id, Descricao) " & vbCrLf & _
                  " Values('" & txtGrupo.Text & "' " & vbCrLf & _
                  ",'" & UCase(txtDescricao.Text) & "')" & vbCrLf

            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Else
                Limpar()
                CarregarUsuarios()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        If Funcoes.VerificaPermissao("Grupos", "ALTERAR") Then
            Sql = "UPDATE Grupos" & vbCrLf & _
                  " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                  " WHERE Grupo_Id = '" & txtGrupo.Text & "' " & vbCrLf
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Else
                Limpar()
                CarregarUsuarios()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        If Funcoes.VerificaPermissao("Grupos", "EXCLUIR") Then
            Sql = "DELETE FROM Grupos" & vbCrLf & _
                  " WHERE Grupo_Id = '" & txtGrupo.Text & "' " & vbCrLf
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Else
                Limpar()
                CarregarUsuarios()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
        End If
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
            If Funcoes.VerificaPermissao("Grupos", "RELATORIO") Then
                Sql = " Select Grupo_Id as Codigo, Descricao From Grupos Order by Grupo_Id "
                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Grupos de Usuários")
                Dim objEmpresa As Cliente = New Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
                parameters.Add("Empresa", objEmpresa.Nome.Trim())
                parameters.Add("CidadeCnpj", objEmpresa.Cidade.Trim() & "/" & objEmpresa.Estado.Codigo.Trim() & vbCrLf & "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
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

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Grupos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class