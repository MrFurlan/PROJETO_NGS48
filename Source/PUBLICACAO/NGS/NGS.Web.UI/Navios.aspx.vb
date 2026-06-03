Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Navios
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                'If Funcoes.VerificaPermissao("Navios", "ACESSAR") Then
                limpar()
                'Else
                '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                '    Exit Sub
                'End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Private Sub limpar()
        txtCodigo.Text = String.Empty
        txtDescricao.Text = String.Empty
        txtCodigo.Enabled = False
        chkAtivo.Checked = True

        lnkNovo.Parent.Visible = True
        lnkExcluir.Parent.Visible = False
        lnkAtualizar.Parent.Visible = False

        carregarNavios()
    End Sub

    Private Sub carregarNavios()
        Dim lstNavios As New ListNavios(True)

        GridNavios.DataSource = lstNavios
        GridNavios.DataBind()
    End Sub

    Protected Sub GridNavios_SelectedIndexChanged(ByVal sernder As Object, ByVal e As System.EventArgs) Handles GridNavios.SelectedIndexChanged
        Try
            txtCodigo.Text = GridNavios.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridNavios.SelectedRow.Cells(2).Text()
            chkAtivo.Checked = GridNavios.SelectedRow.Cells(3).Text()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkNovo.Click
        Try
            'If Funcoes.VerificaPermissao("Navios", "GRAVAR") Then
            Dim ObjNavio As New Navio()
            ObjNavio.Descricao = RTrim(txtDescricao.Text)
            ObjNavio.Ativo = chkAtivo.Checked
            ObjNavio.IUD = "I"

            If ObjNavio.Salvar Then
                MsgBox(Me.Page, "Registro inserido com sucesso.", eTitulo.Sucess)
                limpar()
                carregarNavios()
            End If
            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para gravar registro.", eTitulo.Info)
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAtualizar.Click
        Try
            'If Funcoes.VerificaPermissao("Navios", "ALTERAR") Then
            Dim ObjNavio As New Navio()
            ObjNavio.IUD = "U"
            ObjNavio.Codigo = RTrim(txtCodigo.Text)
            ObjNavio.Descricao = RTrim(txtDescricao.Text)
            ObjNavio.Ativo = chkAtivo.Checked

            If ObjNavio.Salvar Then
                MsgBox(Me.Page, "Registro alterado com sucesso.", eTitulo.Sucess)
                limpar()
                carregarNavios()
            End If
            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            'If Funcoes.VerificaPermissao("Navios", "EXCLUIR") Then
            Dim ObjNavio As New Navio()
            ObjNavio.IUD = "D"
            ObjNavio.Codigo = txtCodigo.Text

            If ObjNavio.Salvar Then
                MsgBox(Me.Page, "Registro excluído com sucesso.", eTitulo.Sucess)
                limpar()
                carregarNavios()
            End If
            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sernder As Object, ByVal e As System.EventArgs) Handles lnkLimpar.Click
        limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("Navios", "RELATORIO") Then

            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir o relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Ñavios")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class