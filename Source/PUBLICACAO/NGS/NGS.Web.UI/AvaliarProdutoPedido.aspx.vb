Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AvaliarProdutoPedido
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            'If Funcoes.VerificaPermissao("AvaliarProdutoPedido", "ACESSAR") Then
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "")
            ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")
            Limpar()
            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Revenda.aspx")
            '    Exit Sub
            'End If
        End If
    End Sub

    Private Sub Limpar()
        ddlSafra.SelectedValue = ""
        Funcoes.VerificaEmpresa(ddlEmpresa)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If ddlEmpresa.SelectedIndex = 0 Or ddlSafra.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Selecione a Empresa e Safra antes de Avaliar.")
                Exit Sub
            End If
            Dim val As New PedidoAvaliacaoProduto()
            Dim Empresa As String() = ddlEmpresa.SelectedValue.ToString.Split("-")
            If val.AvaliarPedidos(Empresa(0), Empresa(1), ddlSafra.SelectedValue) Then
                MsgBox(Me.Page, "Pedidos Avaliados com Sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, "Erro Durante a Avaliaçao dos Pedidos", eTitulo.Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Try
                Funcoes.Ajuda(Me.Page, "AvaliarProdutoPedido")
            Catch ex As Exception
                MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class