Imports System.Security
Imports System.Security.Cryptography
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class SenhaNova
    Inherits BasePage

    Dim Sql As String = ""

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not IsPostBack And IsConnect Then
            txtSenha.Focus()
        End If
    End Sub

    Private Sub Confirmar(ByVal Senha As String)
        Dim SqlArray As New ArrayList
        Dim dataToHash As Byte() = Funcoes.ConvertStringToByteArray(Senha)
        Dim hashValue As Byte() = New MD5CryptoServiceProvider().ComputeHash(dataToHash)
        UsuarioServidor.KeyCodeActive = True
        Sql = "UPDATE Usuarios SET Senha = '" & BitConverter.ToString(hashValue) & "' WHERE (Usuario_Id = '" & IIf(HttpContext.Current.Session("NmUsuario") IsNot Nothing, HttpContext.Current.Session("NmUsuario"), HttpContext.Current.Session("ssNomeUsuario")) & "')"
        SqlArray.Add(Sql)
        If Not Banco.GravaBanco(SqlArray) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
        UsuarioServidor.KeyCodeActive = False
        Response.Redirect("~/Login.aspx")
    End Sub

    Protected Sub btnConfirmar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConfirmar.Click
        If (String.IsNullOrWhiteSpace(txtSenha.Text)) Then
            MsgBox(Me.Page, "É necessário informar o campo senha!")
            Exit Sub
        End If

        If (String.IsNullOrWhiteSpace(txtConfirmacao.Text)) Then
            MsgBox(Me.Page, "É necessário informar o campo confirmação de senha!")
            Exit Sub
        End If

        If (txtSenha.Text.Trim() <> txtConfirmacao.Text.Trim()) Then
            MsgBox(Me.Page, "Os campos de senhas não conferem!")
            Exit Sub
        End If

        Try
            Confirmar(txtSenha.Text.Trim())
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCancelar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancelar.Click
        Response.Redirect("~/Index.aspx")
    End Sub

End Class