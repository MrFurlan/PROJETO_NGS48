Imports NGS.Lib.Negocio

Public Class ucCadastrarTipoDeCertidao
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Informe a descrição.", eTitulo.Info)
            Return False
        End If
        Return True
    End Function

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        txtDescricao.Text = String.Empty
    End Sub

    Protected Sub btnConfirmar_Click(sender As Object, e As EventArgs) Handles btnConfirmar.Click
        Try
            If ValidaCampos() Then
                If Session("ssTipoRetorno") IsNot Nothing Then
                    Dim sql As String = "Insert Into TipoDeCertidao (Descricao) Values ('" & txtDescricao.Text.Trim & "')"
                    If Banco.GravaBanco(sql) Then
                        MsgBox(Me.Page, "Tipo adicionado com sucesso.", eTitulo.Sucess)
                        Popup.CloseDialog(Me.Page, "divCadastrarTipoDeCertidao")

                        Session(Session("ssTipoRetorno")) = "1"
                        CType(Me.Page, IBasePage).Carregar("")
                        Popup.CloseDialog(Me.Page, "divCadastrarTipoDeCertidao")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divCadastrarTipoDeCertidao")
    End Sub
End Class