Imports NGS.Lib.Negocio
Public Class SetoresXCentrosDeCusto
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                'If Funcoes.VerificaPermissao("SetoresXCentrosDeCusto", "ACESSAR") Then

                'Else
                '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "")
                '    Exit Sub
                'End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridSetores_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridCentrosDeCusto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


End Class