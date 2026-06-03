Imports NGS.Lib.Negocio

Public Class TabelasDePrecos
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Revenda)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("TabelasDePrecos", "ACESSAR") Then
                    CarregarTabelasDePrecos()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarTabelasDePrecos()
        Dim lstTabelasDePrecos As New [Lib].Negocio.ListTabelaDePreco()

        GridTabelasDePrecos.DataSource = lstTabelasDePrecos
        GridTabelasDePrecos.DataBind()
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = String.Empty
        txtDescricao.Text = String.Empty
        chkAtivo.Checked = True
        txtCodigo.Enabled = False

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição não foi informado!", eTitulo.Info)
        ElseIf Not chkAtivo.Checked Then
            MsgBox(Me.Page, "Ativo não foi marcado!", eTitulo.Info)
            Return False
        End If
        Return True
    End Function

    Protected Sub GridTabelasDePrecos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            txtCodigo.Text = GridTabelasDePrecos.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridTabelasDePrecos.SelectedRow.Cells(2).Text()

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("TabelasDePrecos", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim objTabelaDePreco As New TabelaDePreco()
                    objTabelaDePreco.Descricao = RTrim(txtDescricao.Text)
                    objTabelaDePreco.Ativo = UCase(RTrim(chkAtivo.Checked))
                    objTabelaDePreco.IUD = "I"
                    If objTabelaDePreco.Salvar Then
                        MsgBox(Me.Page, "Informação inserida com sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarTabelasDePrecos()
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
            If Funcoes.VerificaPermissao("TabelasDePrecos", "ALTERAR") Then
                If ValidarCampos() Then
                    Dim objTabelaDePreco As New TabelaDePreco()
                    objTabelaDePreco.Codigo = txtCodigo.Text
                    objTabelaDePreco.Descricao = UCase(RTrim(txtDescricao.Text))
                    objTabelaDePreco.IUD = "U"
                    If objTabelaDePreco.Salvar Then
                        MsgBox(Me.Page, "Informação alterada com sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarTabelasDePrecos()
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
            If Funcoes.VerificaPermissao("TabelasDePrecos", "EXCLUIR") Then
                Dim objTabelaDePreco As New TabelaDePreco(txtCodigo.Text)
                objTabelaDePreco.IUD = "D"
                If objTabelaDePreco.Salvar Then
                    MsgBox(Me.Page, "Informação excluída com sucesso.", eTitulo.Sucess)
                    Limpar()
                    CarregarTabelasDePrecos()
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
            Funcoes.Ajuda(Me.Page, "TabelasDePrecos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class