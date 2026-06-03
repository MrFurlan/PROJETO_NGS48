Imports Microsoft.IdentityModel.Tokens
Imports System.IdentityModel.Tokens.Jwt
Imports System.Security.Claims
Imports System.Text
Imports System.Web.Http

Public Class TokenController
    Inherits ApiController

    Private ReadOnly secretKey As String = "NGS-2024-APIKwf1Z7b2ZcC5qB8xKd3NvT"

    <HttpPost>
    <Route("api/token")>
    <AllowAnonymous>
    Public Function GenerateToken(<FromBody> ByVal request As LoginRequest) As IHttpActionResult
        ' Verificar as credenciais do usuário (substitua com sua lógica de autenticação)
        If IsValidUser(request.Username, request.Password) Then
            ' Gerar token JWT
            Dim tokenHandler As New JwtSecurityTokenHandler()
            Dim key As Byte() = Encoding.ASCII.GetBytes(secretKey)
            Dim tokenDescriptor As New SecurityTokenDescriptor() With {
                .Subject = New ClaimsIdentity(New List(Of Claim)() From {
                    New Claim(ClaimTypes.Name, request.Username)
                }),
                .Expires = DateTime.UtcNow.AddDays(7), ' Define a validade do token
                .SigningCredentials = New SigningCredentials(New SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha256Signature)
            }
            Dim token As SecurityToken = tokenHandler.CreateToken(tokenDescriptor)
            Dim tokenString As String = tokenHandler.WriteToken(token)

            Return Ok(New With {
                .token = tokenString
            })
        Else
            Return Unauthorized()
        End If
    End Function

    <HttpGet>
    <Authorize>
    <Route("api/protegido")>
    Public Function GetProtegido() As IHttpActionResult
        ' Este método só será acessado se o usuário estiver autenticado
        Dim identity As ClaimsIdentity = TryCast(User.Identity, ClaimsIdentity)
        Dim username As String = identity.FindFirst(ClaimTypes.Name)?.Value

        Return Ok($"Olá, {username}! Você acessou uma rota protegida.")
    End Function

    Private Function IsValidUser(ByVal username As String, ByVal password As String) As Boolean
        ' Aqui consultamos no banco o usuario criado, para acessar o sistema
        ' Este é apenas um exemplo simplificado:
        Return username = "usuario" AndAlso password = "senha"

    End Function

    ' GET api/token/hashpassword?password=yourpassword
    'http://localhost:53957/api/token/hashpassword?password=fgfgfg
    <HttpGet>
    <AllowAnonymous>
    <Route("api/token/hashpassword")>
    Public Function HashPassword(<FromUri> password As String) As IHttpActionResult
        If String.IsNullOrEmpty(password) Then
            Return BadRequest("Password cannot be empty.")
        End If

        Try
            Dim hashedPassword As String = DataAcess.CreateMD5Hash(password)
            Return Ok(New With {Key .HashedPassword = hashedPassword})
        Catch ex As Exception
            Return InternalServerError(ex)
        End Try
    End Function

    ' GET api/token/encrypt?text=otextoparacriptografar
    <AllowAnonymous>
    <HttpGet>
    <Route("api/token/encrypt")>
    Public Function Encrypt(<FromUri> text As String) As IHttpActionResult
        If String.IsNullOrEmpty(text) Then
            Return BadRequest("Text cannot be empty.")
        End If

        Try
            Dim encryptedText As String = EncryptionService.Encrypt(text)
            Return Ok(New With {Key .EncryptedText = encryptedText})
        Catch ex As Exception
            Return InternalServerError(ex)
        End Try
    End Function

    ' GET api/token/decrypt?encryptedText=textoparadescriptografar
    <AllowAnonymous>
    <HttpGet>
    <Route("api/token/decrypt")>
    Public Function Decrypt(<FromUri> encryptedText As String) As IHttpActionResult
        If String.IsNullOrEmpty(encryptedText) Then
            Return BadRequest("Encrypted text cannot be empty.")
        End If

        Try
            Dim decryptedText As String = EncryptionService.Decrypt(encryptedText)
            Return Ok(New With {Key .DecryptedText = decryptedText})
        Catch ex As Exception
            Return InternalServerError(ex)
        End Try
    End Function

End Class

Public Class LoginRequest
    Public Property Username As String
    Public Property Password As String
End Class
