Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Pedidos
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
        End If
    End Sub

End Class