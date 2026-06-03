Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ClasseAmbiental
    Inherits BasePage

    Dim objClasseAmbiental As [Lib].Negocio.ClasseAmbiental

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ClasseAmbiental", "ACESSAR") Then
                AtualizarGrid()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub gridClasseAmbiental_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtCodigo.Enabled = False
            txtCodigo.Text = gridClasseAmbiental.SelectedRow.Cells(1).Text()
            txtDescricao.Text = gridClasseAmbiental.SelectedRow.Cells(2).Text()
            'ss(Page.FindControl("txtDescricao"))
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGrid()
        Dim Lista As New ListClasseAmbiental(True)
        gridClasseAmbiental.DataSource = Lista.ToArray
        gridClasseAmbiental.DataBind()
    End Sub

    Private Sub Limpar()
        objClasseAmbiental = New [Lib].Negocio.ClasseAmbiental
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        txtCodigo.Enabled = True

        Dim textAnotacion As TextBox = Me.Page.FindControl(txtCodigo.UniqueID)
        Dim scriptManager As ScriptManager = scriptManager.GetCurrent(Me.Page)
        textAnotacion.Focus()
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = "Select ClasseAmbiental_Id As Codigo, Descricao From ClasseAmbiental where 1=1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            sql &= "And ClasseAmbiental_Id = " & txtCodigo.Text
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            sql &= "And Descricao = '" & txtDescricao.Text & "'"
        End If

        sql &= "order by ClasseAmbiental_Id"

        Return Banco.ConsultaDataSet(sql, "Tabelas")
    End Function

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ClasseAmbiental", "GRAVAR") Then
                objClasseAmbiental = New [Lib].Negocio.ClasseAmbiental
                objClasseAmbiental.Codigo = txtCodigo.Text
                objClasseAmbiental.Descricao = txtDescricao.Text
                objClasseAmbiental.IUD = "I"
                If objClasseAmbiental.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("ClasseAmbiental", "ALTERAR") Then
                objClasseAmbiental = New [Lib].Negocio.ClasseAmbiental
                objClasseAmbiental.Codigo = gridClasseAmbiental.SelectedRow.Cells(1).Text()
                objClasseAmbiental.Descricao = txtDescricao.Text
                objClasseAmbiental.IUD = "U"
                If objClasseAmbiental.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("ClasseAmbiental", "EXCLUIR") Then
                objClasseAmbiental = New [Lib].Negocio.ClasseAmbiental
                objClasseAmbiental.Codigo = gridClasseAmbiental.SelectedRow.Cells(1).Text()
                objClasseAmbiental.IUD = "D"
                If objClasseAmbiental.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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
            If Funcoes.VerificaPermissao("ClasseAmbiental", "RELATORIO") Then
                Dim ds As DataSet = getDataSet()
                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório de Categoria da Praga")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
                Dim objEmpresa As Cliente = New Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
                parameters.Add("Empresa", objEmpresa.Nome.Trim())
                parameters.Add("CidadeCnpj", objEmpresa.Cidade.Trim() & "/" & objEmpresa.Estado.Codigo.Trim() & vbCrLf & "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))

                Funcoes.BindReport(Me.Page, ds, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ClasseAmbiental")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class