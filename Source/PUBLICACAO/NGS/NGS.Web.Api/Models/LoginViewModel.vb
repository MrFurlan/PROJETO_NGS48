Imports System.ComponentModel.DataAnnotations

Namespace Models

    ''' <summary>
    ''' Classe para mapeamento das informações de Login para acesso ao sistema
    ''' </summary>
    Public Class LoginViewModel

#Region "Propriedades"


        Public Property BancosDeDados As List(Of SelectListItem)
        Public Property BancoId As String

        ''' <summary>
        ''' Usuário para acesso ao sistema
        ''' </summary>
        <Display(Name:="Login")>
        <Required(ErrorMessage:="Login é Obrigatório")>
        <StringLength(255, ErrorMessage:="Máximo de 255 caracteres")>
        Public Property UsuarioId As String

        ''' <summary>
        ''' Senha de acesso ao sistema
        ''' </summary>
        <Display(Name:="Senha")>
        <Required(ErrorMessage:="Senha é Obrigatória")>
        <StringLength(16, ErrorMessage:="Máximo de 16 caracteres")>
        Public Property Password As String

        ''' <summary>
        ''' Variável para Url de Retorno quando tenta navegar e cai a sessão
        ''' </summary>
        Public Property ReturnUrl As String

#End Region

    End Class

End Namespace
