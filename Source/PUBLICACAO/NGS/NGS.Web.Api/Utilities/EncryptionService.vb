Imports System.Security.Cryptography
Imports System.Text

Public Class EncryptionService
    Private Shared ReadOnly Key As Byte()
    Private Shared ReadOnly IV As Byte()

    ' Static constructor to initialize the key and IV
    Shared Sub New()
        ' Using a fixed password and salt for the key generation
        Dim password As String = "SuperSecretPassword"
        Dim salt As String = "S@1tValue12345678"
        Dim keyGenerator As New Rfc2898DeriveBytes(password, Encoding.UTF8.GetBytes(salt))

        ' Generating a 32-byte key and a 16-byte IV
        Key = keyGenerator.GetBytes(32)
        IV = keyGenerator.GetBytes(16)
    End Sub

    Public Shared Function Encrypt(plainText As String) As String
        Using aes As Aes = Aes.Create()
            aes.Key = Key
            aes.IV = IV

            Dim encryptor As ICryptoTransform = aes.CreateEncryptor(aes.Key, aes.IV)

            Using ms As New IO.MemoryStream()
                Using cs As New CryptoStream(ms, encryptor, CryptoStreamMode.Write)
                    Using sw As New IO.StreamWriter(cs)
                        sw.Write(plainText)
                    End Using
                End Using
                Return Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
    End Function

    Public Shared Function Decrypt(cipherText As String) As String
        Dim buffer As Byte() = Convert.FromBase64String(cipherText)

        Using aes As Aes = Aes.Create()
            aes.Key = Key
            aes.IV = IV

            Dim decryptor As ICryptoTransform = aes.CreateDecryptor(aes.Key, aes.IV)

            Using ms As New IO.MemoryStream(buffer)
                Using cs As New CryptoStream(ms, decryptor, CryptoStreamMode.Read)
                    Using sr As New IO.StreamReader(cs)
                        Return sr.ReadToEnd()
                    End Using
                End Using
            End Using
        End Using
    End Function
End Class
