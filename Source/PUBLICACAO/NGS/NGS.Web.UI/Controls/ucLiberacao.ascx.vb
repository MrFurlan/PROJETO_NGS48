Imports AjaxControlToolkit.HtmlEditor.Popups
Imports NGS.[Lib].Negocio
Imports NGS.Lib.Uteis

Public Class ucLiberacao
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            GerarRandomico()
        End If
    End Sub
    Private Sub GerarRandomico()
        txtNumero.Text = (New Random).Next
    End Sub

    Public Overrides Sub Limpar()
        txtNumero.Text = ""
        txtNumeroLiberacao.Text = ""
    End Sub

    Protected Sub lnkLiberar_Click(sender As Object, e As EventArgs) Handles lnkLiberar.Click
        If (txtNumero.Text * 10 - 11 + 1971) = txtNumeroLiberacao.Text Then

            Dim User As New Usuario()
            Session(Session("ssTipoRetorno")) = True
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(User)
                Else
                    CType(Me.Page, IBasePage).Carregar(User)
                End If
                Popup.CloseDialog(Me.Page, "divLiberacao")
            End If
        Else
            MsgBox(Me.Page, "Numero digitado não confere.")
            txtNumeroLiberacao.Text = ""
        End If
    End Sub

    Protected Sub lnkGerar_Click(sender As Object, e As EventArgs) Handles lnkGerar.Click
        GerarRandomico()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divLiberacao")
    End Sub

End Class