Imports System.Security.Cryptography
Imports System.Text
Imports Org.BouncyCastle.Utilities
Public Class RecuperarSenha
    Private _sUsuario As String
    Private _sBanco As String
    Private _lsEmails As List(Of String)
    Private _sToken As String
    Private _sHashToken As String
    Private _bValidarToken As Boolean

    Sub New(sUsuario As String, sBanco)
        _sUsuario = sUsuario
        _lsEmails = Funcoes.BuscarEmailPorUsuario(_sUsuario)
    End Sub

    Private Sub RegistrarToken()
        Try
            Dim bytes(31) As Byte
            Using rng As RandomNumberGenerator = RandomNumberGenerator.Create()
                rng.GetBytes(bytes)
            End Using

            _sToken = Convert.ToBase64String(bytes)

            _sToken = _sToken.Replace("+", "-").
                              Replace("/", "_").
                              Replace("=", "")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub GravarToken()
        Try
            Dim sql As String = "
                INSERT INTO TokensRecuperarSenha
                
            "
        Catch ex As Exception

        End Try
    End Sub

    Public Function EnviarEmailRecuperacao() As Exception
        RegistrarToken()
    End Function

    Private Sub ValidarToken(sToken As String)
        Try
            Dim sql As String = "
                SELECT TokenHash FROM TokensRecuperarSenha
                WHERE Usuario_Id = '" & _sUsuario & "'
                Banco_Id = '" & _sBanco & "'"

            Dim banco As AcessaBanco

            If sToken <> _sHashToken Then
                Throw New Exception("Token inválido!")
            End If
            _bValidarToken = True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub DeletarToken()

    End Sub

End Class
