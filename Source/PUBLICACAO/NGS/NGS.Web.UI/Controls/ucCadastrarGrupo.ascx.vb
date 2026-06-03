Imports NGS.Lib.Negocio

Public Class ucCadastrarGrupo
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtGrupo.Text) Then
            MsgBox(Me.Page, "Informe o grupo.", eTitulo.Info)
            Return False
        ElseIf Len(txtGrupo.Text) < 2 Then
            MsgBox(Me.Page, "Informe no Mínimo 3 caracteres para o grupo.", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Informe a descrição.", eTitulo.Info)
            Return False
        ElseIf Len(txtDescricao.Text) < 6 Then
            MsgBox(Me.Page, "Informe no Mínimo 6 caracteres para a descrição.", eTitulo.Info)
            Return False
        End If
        Return True
    End Function

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Protected Sub btnConfirmar_Click(sender As Object, e As EventArgs) Handles btnConfirmar.Click
        Try
            If ValidaCampos() Then
                If Session("ssTipoRetorno") IsNot Nothing Then
                    Dim sql As String = "Insert Into Grupos (Grupo_Id, Descricao) Values ('" & txtGrupo.Text.Trim() & "', '" & txtDescricao.Text.Trim & "')"
                    If Banco.GravaBanco(sql) Then
                        MsgBox(Me.Page, "Grupo adicionado com Sucesso.", eTitulo.Sucess)
                        Selecionar()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Session(Session("ssTipoRetorno")) = "1"
            CType(Me.Page, IBasePage).Carregar("")
            Popup.CloseDialog(Me.Page, "divCadastrarGrupo")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divCadastrarGrupo")
    End Sub

    Public Overrides Sub Limpar()
        txtDescricao.Text = String.Empty
        txtGrupo.Text = String.Empty
    End Sub

End Class