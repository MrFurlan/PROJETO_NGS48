Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio

Public Class CodificaTexto
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnGerar_Click(sender As Object, e As EventArgs) Handles btnGerar.Click
        Try
            lblResult64Bits.Text = FuncoesStrings.CodificarPara64Bits(txtTexto.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub btnDecript_Click(sender As Object, e As EventArgs) Handles btnDecript.Click
        Try
            lblResult64Bits.Text = FuncoesStrings.DecodificarDe64Bits(txtTexto.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub
End Class