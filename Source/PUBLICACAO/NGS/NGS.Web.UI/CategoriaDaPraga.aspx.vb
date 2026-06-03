Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CategoriaDaPraga
    Inherits BasePage

    Dim objCategoriaPraga As CategoriaPraga

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CategoriaDaPraga", "ACESSAR") Then
                Limpar()
                AtualizarGrid()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub gridCategoriaDaPraga_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = True
        lnkExcluir.Parent.Visible = True
        txtCodigo.Enabled = False
        txtCodigo.Text = gridCategoriaDaPraga.SelectedRow.Cells(1).Text()
        txtDescricao.Text = gridCategoriaDaPraga.SelectedRow.Cells(2).Text()
        txtDescricao.Focus()
    End Sub

    Sub AtualizarGrid()
        Dim Lista As New ListCategoriaPraga(True)
        gridCategoriaDaPraga.DataSource = Lista.ToArray()
        gridCategoriaDaPraga.DataBind()
    End Sub

    Sub Limpar()
        objCategoriaPraga = New CategoriaPraga
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        txtCodigo.Enabled = True
        txtCodigo.Focus()
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CategoriaDaPraga", "GRAVAR") Then
                objCategoriaPraga = New CategoriaPraga
                objCategoriaPraga.Codigo = txtCodigo.Text
                objCategoriaPraga.Descricao = txtDescricao.Text
                objCategoriaPraga.IUD = "I"
                If objCategoriaPraga.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para Gravar Registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("CategoriaDaPraga", "ALTERAR") Then
                objCategoriaPraga = New CategoriaPraga
                objCategoriaPraga.Codigo = gridCategoriaDaPraga.SelectedRow.Cells(1).Text()
                objCategoriaPraga.Descricao = txtDescricao.Text
                objCategoriaPraga.IUD = "U"
                If objCategoriaPraga.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
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
            If Funcoes.VerificaPermissao("CategoriaDaPraga", "EXCLUIR") Then
                objCategoriaPraga = New CategoriaPraga
                objCategoriaPraga.Codigo = gridCategoriaDaPraga.SelectedRow.Cells(1).Text()
                objCategoriaPraga.IUD = "D"
                If objCategoriaPraga.Salvar Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("CategoriaDaPraga", "RELATORIO") Then
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

    Private Function getDataSet() As DataSet
        Dim sql As String = "Select Categoria_Id as Codigo, Descricao From Categoria where 1=1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            sql &= "And Categoria_Id = " & txtCodigo.Text
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            sql &= "And Descricao = '" & txtDescricao.Text & "'"
        End If

        sql &= "order by Categoria_Id"

        Return Banco.ConsultaDataSet(sql, "Tabelas")
    End Function

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CategoriaDaPraga")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class