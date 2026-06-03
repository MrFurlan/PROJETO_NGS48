Imports System.Data.SqlClient

Public Class BancoDados

#Region "Propriedades"

    Public Property Banco_Id As String
    Public Property HostServidor As String
    Public Property UsuarioBanco As String
    Public Property SenhaBanco As String
    Public Property IP As String

#End Region

#Region "Métodos"

    Public Shared Function Listar() As List(Of BancoDados)
        Try
            Dim lsBancos As New List(Of BancoDados)
            Using conexao As New SqlConnection(AcessaBanco.GetStringBancoUsuarios())
                If (conexao IsNot Nothing) Then
                    conexao.Open()
                    Dim strSQL As String = "SELECT Banco_Id, HostDoServidor, UsuarioDoBanco, SenhaDoBanco, IP " &
                                           "FROM Bancos B " &
                                           "ORDER BY Banco_Id"
                    Using cmd As New SqlCommand(strSQL, conexao)
                        Dim dr As SqlDataReader = cmd.ExecuteReader()
                        Do While dr.Read()
                            Dim objBanco As New BancoDados
                            objBanco.Banco_Id = dr("Banco_Id").ToString()
                            objBanco.HostServidor = dr("HostDoServidor").ToString()
                            objBanco.UsuarioBanco = dr("UsuarioDoBanco").ToString()
                            objBanco.SenhaBanco = dr("SenhaDoBanco").ToString()
                            objBanco.IP = dr("IP").ToString()
                            lsBancos.Add(objBanco)
                        Loop
                    End Using
                Else
                    Throw New Exception("Não foi possível conectar ao banco de dados!")
                End If
            End Using
            Return lsBancos
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function ListarBancoPorUsuarios(ByVal sUsuario As String) As List(Of BancoDados)
        Try
            Dim lsBancos As New List(Of BancoDados)
            Using conexao As New SqlConnection(AcessaBanco.GetStringBancoUsuarios())
                If (conexao IsNot Nothing) Then
                    conexao.Open()
                    Dim strSQL As String = " SELECT Banco_Id, HostDoServidor, UsuarioDoBanco, SenhaDoBanco, IP
                                             
                                             FROM Bancos B
                                             WHERE Banco_Id in 
                                                ( SELECT NomeDoBanco FROM Usuarios
                                                    WHERE Usuario_Id = '" & sUsuario & "')
                                             ORDER BY Banco_Id"
                    Using cmd As New SqlCommand(strSQL, conexao)
                        Dim dr As SqlDataReader = cmd.ExecuteReader()
                        Do While dr.Read()
                            Dim objBanco As New BancoDados
                            objBanco.Banco_Id = dr("Banco_Id").ToString()
                            objBanco.HostServidor = dr("HostDoServidor").ToString()
                            objBanco.UsuarioBanco = dr("UsuarioDoBanco").ToString()
                            objBanco.SenhaBanco = dr("SenhaDoBanco").ToString()
                            objBanco.IP = dr("IP").ToString()
                            lsBancos.Add(objBanco)
                        Loop
                    End Using
                Else
                    Throw New Exception("Não foi possível conectar ao banco de dados!")
                End If
            End Using
            Return lsBancos
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

End Class