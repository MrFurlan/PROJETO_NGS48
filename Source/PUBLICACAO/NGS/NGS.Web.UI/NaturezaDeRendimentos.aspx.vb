Imports System.Data
Imports System.Collections
Imports System.IO
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis


Public Class NaturezaDeRendimentos
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("NaturezaDeRendimentos", "ACESSAR") Then
                    Limpar()
                    CarregarNaturezaDeRendimentos()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarNaturezaDeRendimentos()
        Dim nome As String = String.Empty
        Dim lstNaturezaDeRendimentos As New [Lib].Negocio.NaturezaDeRendimentos("")

        Dim dtNDR As New DataTable("ItemEPI")
        dtNDR.Columns.Add("Codigo", Type.GetType("System.Int64"))
        dtNDR.Columns.Add("Descricao", Type.GetType("System.String"))
        dtNDR.Columns.Add("TipoPessoa", Type.GetType("System.String"))

        For Each item In lstNaturezaDeRendimentos

            Dim drNDR As DataRow = dtNDR.NewRow()
            drNDR("Codigo") = item.Codigo
            drNDR("Descricao") = item.Descricao

            If item.TipoPessoa = eTipoPessoa.Fisica Then
                drNDR("TipoPessoa") = "FÍSICA"
            ElseIf item.TipoPessoa = eTipoPessoa.Juridica Then
                drNDR("TipoPessoa") = "JURÍDICA"
            Else
                drNDR("TipoPessoa") = "AMBAS"
            End If

            dtNDR.Rows.Add(drNDR)
        Next

        GridNaturezaDeRendimentos.DataSource = dtNDR
        GridNaturezaDeRendimentos.DataBind()
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        ddlPessoa.SelectedIndex = 0
        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            MsgBox(Me.Page, "Código do Estado não foi informado!", eTitulo.Info)
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição do Estado não foi informada!", eTitulo.Info)
            Return False
        ElseIf ddlPessoa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Tipo da Pessoa não foi selecionado!", eTitulo.Info)
            Return False
        End If
        Return True
    End Function

    Protected Sub GridNaturezaDeRendimentos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            txtCodigo.Text = GridNaturezaDeRendimentos.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridNaturezaDeRendimentos.SelectedRow.Cells(2).Text()

            If GridNaturezaDeRendimentos.SelectedRow.Cells(3).Text() = "FÍSICA" Then
                ddlPessoa.SelectedValue = eTipoPessoa.Fisica
            ElseIf GridNaturezaDeRendimentos.SelectedRow.Cells(3).Text() = "JURÍDICA" Then
                ddlPessoa.SelectedValue = eTipoPessoa.Juridica
            Else
                ddlPessoa.SelectedValue = eTipoPessoa.Ambas
            End If

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtCodigo.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("NaturezaDeRendimentos", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim objNaturezaDeRendimento As New NaturezaDeRendimento()
                    objNaturezaDeRendimento.Codigo = CInt(txtCodigo.Text)
                    objNaturezaDeRendimento.Descricao = RTrim(txtDescricao.Text)
                    objNaturezaDeRendimento.TipoPessoa = ddlPessoa.SelectedValue
                    objNaturezaDeRendimento.IUD = "I"
                    If objNaturezaDeRendimento.Salvar Then
                        MsgBox(Me.Page, "Informação inserida com sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarNaturezaDeRendimentos()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("NaturezaDeRendimentos", "ALTERAR") Then
                If ValidarCampos() Then
                    Dim objNaturezaDeRendimento As New NaturezaDeRendimento(CInt(txtCodigo.Text))
                    objNaturezaDeRendimento.Descricao = RTrim(txtDescricao.Text)
                    objNaturezaDeRendimento.TipoPessoa = ddlPessoa.SelectedValue
                    objNaturezaDeRendimento.IUD = "U"
                    If objNaturezaDeRendimento.Salvar Then
                        MsgBox(Me.Page, "Informação alterada com sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarNaturezaDeRendimentos()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("NaturezaDeRendimentos", "EXCLUIR") Then
                If ValidarCampos() Then
                    Dim objNaturezaDeRendimento As New NaturezaDeRendimento(CInt(txtCodigo.Text))

                    If objNaturezaDeRendimento.temNF Then
                        MsgBox(Me.Page, "Código não pode ser removido pois está sendo utilizado em Nota Fiscal.", eTitulo.Info)
                        Exit Sub
                    End If

                    objNaturezaDeRendimento.Descricao = RTrim(txtDescricao.Text)
                    objNaturezaDeRendimento.IUD = "D"
                    If objNaturezaDeRendimento.Salvar Then
                        MsgBox(Me.Page, "Informação excluída com sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarNaturezaDeRendimentos()
                    End If
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "NaturezaDeRendimentos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class