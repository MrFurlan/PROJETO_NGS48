Imports System.Data.SqlClient
Imports NGS.Lib.Uteis

Public Class UsuarioLocal

#Region "Métodos"

    Public Shared Function GetStringConexao(ByVal Usuario As String, ByRef NomeServidor As String) As String
        Dim strStringConexao As String = String.Empty

        Using conexao As New SqlConnection(AcessaBanco.GetStringBancoUsuarios())
            conexao.Open()
            Dim strSQL As String = "SELECT U.Usuario_Id, B.Banco_Id, B.HostDoServidor, B.UsuarioDoBanco, B.SenhaDoBanco " & _
                                   "FROM Bancos B " & _
                                   "INNER JOIN Usuarios U ON B.Banco_Id = U.NomeDoBanco " & _
                                                         "AND B.HostDoServidor = U.HostDoServidor " & _
                                   "WHERE U.Usuario_Id = @Usuario_Id"

            Using cmm As New SqlCommand(strSQL, conexao)
                cmm.Parameters.Add("@Usuario_Id", SqlDbType.NVarChar, 100).Value = Usuario
                Dim dr As SqlDataReader = cmm.ExecuteReader()
                If dr.Read() Then
                    strStringConexao = String.Format("Data Source={0}; Initial Catalog={1}; User Id={2}; Password={3}", _
                                                     dr("HostDoServidor"), dr("Banco_Id"), dr("UsuarioDoBanco"), FuncoesStrings.DecodificarDe64Bits(dr("SenhaDoBanco")))
                    NomeServidor = dr("HostDoServidor")
                End If
            End Using
        End Using

        Return strStringConexao
    End Function

    Public Shared Function GetStringConexao(ByVal Usuario As String, ByRef NomeServidor As String, ByVal Banco As String, ByVal Host As String) As String
        Dim strStringConexao As String = String.Empty
        Using conexao As New SqlConnection(AcessaBanco.GetStringBancoUsuarios())
            conexao.Open()
            Dim strSQL As String = "SELECT  @Usuario_Id AS Usuario, Banco_Id, " +
                                            "CASE WHEN ISNULL(IP, '') = '' THEN " +
                                                        "HostDoServidor " +
                                                "ELSE IP END AS HostDoServidor, UsuarioDoBanco, SenhaDoBanco " & _
                                   "FROM    Bancos " & _
                                   "WHERE   Banco_Id        = @Banco_Id " & _
                                   "AND     HostDoServidor  = @HostServidor"


            Using cmm As New SqlCommand(strSQL, conexao)
                cmm.Parameters.Add("@Usuario_Id", SqlDbType.NVarChar, 100).Value = Usuario
                cmm.Parameters.Add("@Banco_Id", SqlDbType.NVarChar, 50).Value = Banco
                cmm.Parameters.Add("@HostServidor", SqlDbType.NVarChar, 50).Value = Host
                Dim dr As SqlDataReader = cmm.ExecuteReader()
                If dr.Read() Then
                    strStringConexao = String.Format("Data Source={0}; Initial Catalog={1}; User Id={2}; Password={3}",
                                                     dr("HostDoServidor"), dr("Banco_Id"), dr("UsuarioDoBanco"), FuncoesStrings.DecodificarDe64Bits(dr("SenhaDoBanco")))
                    NomeServidor = dr("HostDoServidor")
                End If
            End Using
        End Using
        Return strStringConexao
    End Function

#End Region

End Class