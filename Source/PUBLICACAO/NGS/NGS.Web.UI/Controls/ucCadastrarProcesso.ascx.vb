Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucCadastrarProcesso
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then

        End If
    End Sub

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtProcesso.Text) Then
            MsgBox(Me.Page, "Informe o processo.")
            Return False
        ElseIf Len(txtProcesso.Text) < 3 Then
            MsgBox(Me.Page, "Informe no Mínimo 3 caracteres para o processo.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Informe a descrição.")
            Return False
        ElseIf Len(txtDescricao.Text) < 8 Then
            MsgBox(Me.Page, "Informe no Mínimo 8 caracteres para a descrição.")
            Return False
        End If
        Return True
    End Function

    Protected Sub btnConfirmar_Click(sender As Object, e As EventArgs) Handles btnConfirmar.Click
        Try
            If ValidaCampos() Then
                If Session("ssTipoRetorno") IsNot Nothing Then
                    Dim sql As String = "Insert Into Processos (Processo_Id, Descricao) Values ('" & txtProcesso.Text.Trim() & "', '" & txtDescricao.Text.Trim & "')"
                    If Banco.GravaBanco(sql) Then
                        MsgBox(Me.Page, "Processo adicionado com Sucesso.", eTitulo.Sucess)
                        Selecionar()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Limpar()
        txtDescricao.Text = String.Empty
        txtProcesso.Text = String.Empty
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Protected Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divCadastrarProcesso")
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Session(Session("ssTipoRetorno")) = txtProcesso.Text.ToUpper()
            CType(Me.Page, IBasePage).Carregar("")
            Popup.CloseDialog(Me.Page, "divCadastrarProcesso")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class