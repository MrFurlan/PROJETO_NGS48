Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ModoDeAplicacao
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            'If Funcoes.VerificaPermissao("ModoDeAplicacao", "ACESSAR") Then
            carregarFormaDeAplicacao()
            Limpar()
            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Gestao.aspx")
            '    Exit Sub
            'End If
        End If
    End Sub

    Private Sub carregarddlCulturaPragaFito()
        'Dim Lista As New Negocio.ListFormaDeAplicaca

        'ddlCulturaPragaFito.DataValueField = "Codigo"
        'ddlCulturaPragaFito.DataTextField = "Descricao"
        'ddlCulturaPragaFito.DataSource = Lista.ToDataTable()
        'ddlCulturaPragaFito.DataBind()

        'Funcoes.InserirLinhaEmBranco(ddlCulturaPragaFito)
    End Sub

    Private Sub carregarFormaDeAplicacao()
        Dim Lista As New [Lib].Negocio.ListFormaDeAplicacao(True)

        ddlFormaDeAplicacao.DataValueField = "Codigo"
        ddlFormaDeAplicacao.DataTextField = "Descricao"
        ddlFormaDeAplicacao.DataSource = Lista.ToArray()
        ddlFormaDeAplicacao.DataBind()

        Funcoes.InserirLinhaEmBranco(ddlFormaDeAplicacao)
    End Sub

    Private Sub carregarGrid(ByVal Codigo As String, ByVal Forma As String)
        Dim Lista As New [Lib].Negocio.ListModoDeAplicacao(Codigo, Forma, "", "", "")

        gridModoDeAplicacao.DataSource = Lista.ToArray()
        gridModoDeAplicacao.DataBind()
    End Sub

    Private Sub Limpar()
        ddlCulturaPragaFito.Enabled = True
        ddlFormaDeAplicacao.Enabled = True
        txtDescricao.Enabled = True

        lnkNovo.Enabled = True
        lnkAtualizar.Enabled = True
        lnkExcluir.Enabled = True
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click

    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click

    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click

    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click

    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ModoDeAplicacao")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class