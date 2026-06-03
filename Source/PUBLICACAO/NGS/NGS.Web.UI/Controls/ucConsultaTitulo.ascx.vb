Imports Microsoft.Win32
Imports MoreLinq
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaTitulo
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            Me.Limpar()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        If IsNumeric(txtRegistro.Text) AndAlso CInt(txtRegistro.Text) > 0 Then

            Session(Session("ssTipoRetorno")) = txtRegistro.Text
            Dim ObjTitulo As New NGS.Lib.Negocio.TituloV(txtRegistro.Text)

            If Session("ssTipoRetorno") IsNot Nothing Then
                Dim str As String = txtRegistro.Text
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(ObjTitulo)
                Else
                    CType(Me.Page, IBasePage).Carregar(ObjTitulo)
                End If

                Popup.CloseDialog(Me.Page, "divConsultaTitulo")
            End If
        Else
            MsgBox(Me.Page, "Digite o número do título para Consultar.")
        End If

    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaTitulo")
    End Sub

    Public Overrides Sub Limpar()
        txtRegistro.Text = String.Empty
    End Sub
End Class