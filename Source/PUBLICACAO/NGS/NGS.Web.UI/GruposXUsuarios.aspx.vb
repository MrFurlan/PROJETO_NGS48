Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class GruposXUsuarios
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("GruposXUsuarios", "ACESSAR") Then
                CarregarGrupos()
                CarregarUsuarios()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarGrupos()
        Sql = "Select Grupo_Id as Grupo From Grupos Order By Grupo_Id"
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlGrupos.Items.Add(New ListItem(UCase(Dr("Grupo")), UCase(Dr("Grupo"))))
        Next
        DdlGrupos.Items.Insert(0, "")
        DdlGrupos.SelectedIndex = 0
    End Sub

    Private Sub CarregarUsuarios()
        Sql = "Select Usuario_Id as Usuario From Usuarios Order By Usuario_Id"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlUsuarios.Items.Add(New ListItem(UCase(Dr("Usuario")), UCase(Dr("Usuario"))))
        Next
        DdlUsuarios.Items.Insert(0, "")
        DdlUsuarios.SelectedIndex = 0
    End Sub

    Private Sub CarregarGruposXUsuarios()
        Sql = "Select Usuario_Id as Usuario, Grupo_Id as Grupo From GruposXUsuarios "
        Sql &= " Where Grupo_Id = '" & UCase(DdlGrupos.SelectedValue) & "' Order By Usuario_Id"

        GridUsuarios.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
        GridUsuarios.DataBind()
    End Sub

    Private Sub Limpar()
        DdlUsuarios.SelectedIndex = 0
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            DdlGrupos.SelectedValue = GridUsuarios.SelectedRow.Cells(1).Text()
            DdlUsuarios.SelectedValue = GridUsuarios.SelectedRow.Cells(2).Text()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlGrupos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarGruposXUsuarios()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("GruposXUsuarios", "GRAVAR") Then
                Sql = "INSERT Into GruposXUsuarios(Grupo_Id, Usuario_id) " & vbCrLf & _
                      " Values('" & UCase(DdlGrupos.SelectedValue) & "'" & vbCrLf & _
                      ",'" & UCase(DdlUsuarios.SelectedValue) & "')" & vbCrLf
                SqlArray.Add(Sql)
                If Not Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarGruposXUsuarios()
                    lnkNovo.Parent.Visible = False
                    lnkAtualizar.Parent.Visible = True
                    lnkExcluir.Parent.Visible = True
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("GruposXUsuarios", "EXCLUIR") Then
                Sql = "DELETE FROM GruposXUsuarios" & vbCrLf & _
                      " WHERE Grupo_Id = '" & UCase(DdlGrupos.SelectedValue) & "' And Usuario_ID = '" & UCase(DdlUsuarios.SelectedValue) & "'" & vbCrLf
                SqlArray.Add(Sql)
                If Not Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarGruposXUsuarios()
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
            If Funcoes.VerificaPermissao("GruposXUsuarios", "RELATORIO") Then
                If DdlGrupos.Text = "" Then
                    Sql = "Select Grupo_Id as Codigo, Usuario_Id as Descricao From GruposXUsuarios Order by Grupo_Id, Usuario_Id"
                Else
                    Sql = "Select Grupo_Id as Codigo, Usuario_Id as Descricao From GruposXUsuarios " & vbCrLf & _
                          " Where Grupo_Id = '" & DdlGrupos.SelectedValue & "' Order by Usuario_Id" & vbCrLf
                End If
                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Grupos X Usuários")
                parameters.Add("Codigo", "Grupo")
                parameters.Add("Descricao", "Usuário")
                Dim objEmpresa As Cliente = New Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
                parameters.Add("Empresa", objEmpresa.Nome.Trim())
                parameters.Add("CidadeCnpj", objEmpresa.Cidade.Trim() & "/" & objEmpresa.Estado.Codigo.Trim() & vbCrLf & "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))

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
            Funcoes.Ajuda(Me.Page, "GruposXUsuarios")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class