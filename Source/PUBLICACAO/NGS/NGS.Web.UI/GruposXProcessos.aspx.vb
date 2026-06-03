Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class GruposXProcessos
    Inherits BasePage

    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Dim Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("GruposXProcessos", "ACESSAR") Then
                    CarregarGrupos()
                    CarregarProcessos()
                    Limpar()
                    DdlGrupos.SelectedIndex = 0
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarGrupos()
        Sql = "Select Grupo_Id as Grupo From Grupos Order By Grupo_Id"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlGrupos.Items.Add(New ListItem(UCase(Dr("Grupo")), UCase(Dr("Grupo"))))
        Next
        DdlGrupos.Items.Insert(0, "")
        DdlGrupos.SelectedIndex = 0
    End Sub

    Private Sub CarregarProcessos()
        Sql = "Select Processo_Id as Processo From Processos Order By Processo_Id"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlProcessos.Items.Add(New ListItem(UCase(Dr("Processo")), UCase(Dr("Processo"))))
        Next
        DdlProcessos.Items.Insert(0, "")
        DdlProcessos.SelectedIndex = 0
    End Sub

    Private Sub CarregarGruposXProcessos()
        Sql = "Select upper(Grupo_Id) as Grupo, upper(Processo_Id) as Processo From GruposXProcessos "
        Sql &= " Where Grupo_Id = '" & UCase(DdlGrupos.SelectedValue) & "' Order By Processo_Id"

        GridProcessos.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
        GridProcessos.DataBind()
    End Sub

    Private Sub Limpar()

        DdlProcessos.SelectedIndex = 0
        lnkNovo.Parent.Visible = True
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            DdlGrupos.SelectedValue = GridProcessos.SelectedRow.Cells(1).Text()
            DdlProcessos.SelectedValue = GridProcessos.SelectedRow.Cells(2).Text()
            lnkNovo.Parent.Visible = False
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlGrupos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarGruposXProcessos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("GruposXProcessos", "GRAVAR") Then
                Sql = "INSERT Into GruposXProcessos(Grupo_Id, Processo_id) " & vbCrLf & _
                      " Values('" & UCase(DdlGrupos.SelectedValue) & "'" & vbCrLf & _
                      ",'" & UCase(DdlProcessos.SelectedValue) & "')" & vbCrLf
                SqlArray.Add(Sql)
                If Not Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarGruposXProcessos()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("GruposXProcessos", "EXCLUIR") Then
                Sql = "DELETE FROM GruposXProcessos" & vbCrLf & _
                      " WHERE Grupo_Id = '" & UCase(DdlGrupos.SelectedValue) & "' And Processo_ID = '" & UCase(DdlProcessos.SelectedValue) & "'" & vbCrLf
                SqlArray.Add(Sql)
                If Not Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarGruposXProcessos()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("GruposXProcessos", "RELATORIO") Then
                If DdlGrupos.Text = "" Then
                    Sql = "Select Grupo_Id as Codigo, Processo_Id as Descricao From GruposXProcessos Order by Grupo_Id, Processos_Id"
                Else
                    Sql = "Select Grupo_Id as Codigo, Processo_Id as Descricao From GruposXProcessos " & vbCrLf & _
                          " Where Grupo_Id = '" & DdlGrupos.SelectedValue & "' Order by Processo_Id" & vbCrLf
                End If
                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Grupos X Processos")
                parameters.Add("Codigo", "Codigo")
                parameters.Add("Descricao", "Descricao")
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "GruposXProcessos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class