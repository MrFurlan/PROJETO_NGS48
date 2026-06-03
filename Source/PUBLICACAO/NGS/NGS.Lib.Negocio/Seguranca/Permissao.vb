Imports System.Data.SqlClient
Imports System.Web

<Serializable()> _
Public Class PermissaoUsuario

    Public Property Usuario_Id As String
    Public Property Processo_Id As String
    Public Property PermiteGravar As Boolean
    Public Property PermiteAlterar As Boolean
    Public Property PermiteExcluir As Boolean
    Public Property PermiteLeitura As Boolean
    Public Property PermiteAcesso As Boolean
    Public Property PermiteLiberar As Boolean
    Public Property PermiteImprimir As Boolean

    Public Shared Function VerificaPermissoes(ByVal NomePrograma As String, ByVal Usuario As String) As PermissaoUsuario
        Using cnn As New SqlConnection(HttpContext.Current.Session("Conexao"))
            cnn.Open()

            Dim strSQL As String = "SELECT      P.Permissao_Id " & _
                                   "FROM        GruposXProcessosXPermissoes GPP " & _
                                   "INNER JOIN  GruposXUsuarios             GU  ON  GPP.Grupo_Id        = GU.Grupo_Id " & _
                                   "INNER JOIN  Usuarios                    U   ON  GU.Usuario_Id       = U.Usuario_Id " & _
                                   "INNER JOIN  Permissoes                  P   ON  GPP.Permissao_Id    = P.Permissao_Id " & _
                                   "WHERE       GPP.Processo_Id = @Processo_Id " & _
                                   "AND         GU.Usuario_Id   = @Usuario_Id"

            Using cmm As New SqlCommand(strSQL, cnn)
                cmm.Parameters.Add("@Processo_Id", SqlDbType.NVarChar, 100).Value = NomePrograma
                cmm.Parameters.Add("@Usuario_Id", SqlDbType.NVarChar, 100).Value = Usuario

                Dim dr As SqlDataReader = cmm.ExecuteReader()

                Dim objPermissao As New PermissaoUsuario()
                objPermissao.Usuario_Id = Usuario
                objPermissao.Processo_Id = NomePrograma
                objPermissao.PermiteGravar = False
                objPermissao.PermiteAlterar = False
                objPermissao.PermiteExcluir = False
                objPermissao.PermiteLeitura = False
                objPermissao.PermiteAcesso = False
                objPermissao.PermiteLiberar = False
                objPermissao.PermiteImprimir = False

                While dr.Read()
                    Select Case dr("Permissao_Id")
                        Case "GRAVAR" : objPermissao.PermiteGravar = True
                        Case "ALTERAR" : objPermissao.PermiteAlterar = True
                        Case "EXCLUIR" : objPermissao.PermiteExcluir = True
                        Case "LEITURA" : objPermissao.PermiteLeitura = True
                        Case "ACESSAR" : objPermissao.PermiteAcesso = True
                        Case "LIBERAR" : objPermissao.PermiteLiberar = True
                        Case "RELATORIO" : objPermissao.PermiteImprimir = True
                    End Select
                End While

                Return objPermissao
            End Using
        End Using
    End Function

End Class