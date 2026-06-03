Imports System.Web
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Configuration
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

<Serializable()>
Public NotInheritable Class UsuarioServidor

#Region "Construtor"

    Private Sub New()
    End Sub

#End Region

#Region "Sessão"

    Public Shared Property KeyCodeActive() As Boolean
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssKeyCodeActive") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssKeyCodeActive"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As Boolean)
            HttpContext.Current.Session("ssKeyCodeActive") = value
        End Set
    End Property

    Public Shared Property Usuario() As [Lib].Negocio.Usuario
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssUsuario") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssUsuario"), [Lib].Negocio.Usuario)
            End If
            Return Nothing
        End Get
        Set(ByVal value As [Lib].Negocio.Usuario)
            HttpContext.Current.Session("ssUsuario") = value
        End Set
    End Property

    Public Shared Property NomeUsuario() As String
        Get
            If HttpContext.Current Is Nothing Then Return Nothing

            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssNomeUsuario") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssNomeUsuario"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssNomeUsuario") = value
        End Set
    End Property

    Public Shared Property NomeServidor() As String
        Get
            If HttpContext.Current Is Nothing Then Return String.Empty

            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssNomeServidor") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssNomeServidor"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssNomeServidor") = value
        End Set
    End Property

    Public Shared Property EnderecoLocal() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssEnderecoLocal") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssEnderecoLocal"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssEnderecoLocal") = value
        End Set
    End Property

    Public Shared Property BancoDeDados() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssBancoDeDados") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssBancoDeDados"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssBancoDeDados") = value
        End Set
    End Property

    Public Shared Property NmUsuario() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("NmUsuario") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("NmUsuario"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("NmUsuario") = value
        End Set
    End Property

    Public Shared Property CodigoEmpresa() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssEmpresa") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssEmpresa"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssEmpresa") = value
        End Set
    End Property

    Public Shared Property EnderecoEmpresa() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssEndEmpresa") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssEndEmpresa"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssEndEmpresa") = value
        End Set
    End Property

    Public Shared Property NomeEmpresa() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssNomeEmpresa") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssNomeEmpresa"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssNomeEmpresa") = value
        End Set
    End Property

    Public Shared Property CidadeEmpresa() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssCidadeEmpresa") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssCidadeEmpresa"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssCidadeEmpresa") = value
        End Set
    End Property

    Public Shared Property EstadoEmpresa() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssEstadoEmpresa") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssEstadoEmpresa"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssEstadoEmpresa") = value
        End Set
    End Property

    Public Shared Property ReduzidoEmpresa() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssReduzidoEmpresa") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssReduzidoEmpresa"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssReduzidoEmpresa") = value
        End Set
    End Property

    Public Shared Property ImagemEmpresa() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssImagemEmpresa") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssImagemEmpresa"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssImagemEmpresa") = value
        End Set
    End Property

    Public Shared Property MenuDeAcesso() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssMenuDeAcesso") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssMenuDeAcesso"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssMenuDeAcesso") = value
        End Set
    End Property

    Public Shared Property ImprimirDanfe() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssImprimirDanfe") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssImprimirDanfe"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssImprimirDanfe") = value
        End Set
    End Property

    Public Shared Property UsuarioAtivo() As Boolean
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssUsuarioAtivo") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssUsuarioAtivo"), Boolean)
            End If
            Return Nothing
        End Get
        Set(ByVal value As Boolean)
            HttpContext.Current.Session("ssUsuarioAtivo") = value
        End Set
    End Property

    Public Shared Property LiberaEmpresa() As Boolean
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssLiberaEmpresa") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssLiberaEmpresa"), Boolean)
            End If
            Return Nothing
        End Get
        Set(ByVal value As Boolean)
            HttpContext.Current.Session("ssLiberaEmpresa") = value
        End Set
    End Property

    Public Shared Property TrocaEmpresa() As Boolean
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssTrocaEmpresa") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssTrocaEmpresa"), Boolean)
            End If
            Return Nothing
        End Get
        Set(ByVal value As Boolean)
            HttpContext.Current.Session("ssTrocaEmpresa") = value
        End Set
    End Property

    Public Shared Property LiberaJanela() As Boolean
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssLiberaJanela") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssLiberaJanela"), Boolean)
            End If
            Return Nothing
        End Get
        Set(ByVal value As Boolean)
            HttpContext.Current.Session("ssLiberaJanela") = value
        End Set
    End Property


    Public Shared Property Conexao() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("Conexao") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("Conexao"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("Conexao") = value
        End Set
    End Property

    Public Shared Property DataInicioDashboard() As DateTime
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssDataInicioDashboard") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssDataInicioDashboard"), DateTime)
            End If
            Return Nothing
        End Get
        Set(ByVal value As DateTime)
            HttpContext.Current.Session("ssDataInicioDashboard") = value
        End Set
    End Property

    Public Shared Property DataFimDashboard() As DateTime
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssDataFimDashboard") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssDataFimDashboard"), DateTime)
            End If
            Return Nothing
        End Get
        Set(ByVal value As DateTime)
            HttpContext.Current.Session("ssDataFimDashboard") = value
        End Set
    End Property

    Public Shared Property SeguimentoDashboard() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssSeguimentoDashboard") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssSeguimentoDashboard"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssSeguimentoDashboard") = value
        End Set
    End Property

    Public Shared Property EmpresasDashboard() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssEmpresasDashboard") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssEmpresasDashboard"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssEmpresasDashboard") = value
        End Set
    End Property

    Public Shared Property AcessoCusto() As String
        Get
            If HttpContext.Current.Session IsNot Nothing AndAlso HttpContext.Current.Session("ssAcessoCusto") IsNot Nothing Then
                Return CType(HttpContext.Current.Session("ssAcessoCusto"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            HttpContext.Current.Session("ssAcessoCusto") = value
        End Set
    End Property

#End Region

#Region "Variáveis"
    Private Shared _Empresa As Cliente
#End Region

#Region "Funções"

    Public Shared Function Mapper(ByVal dr As SqlDataReader) As Usuario

        Dim obj As New Usuario()
        obj.Usuario_Id = dr("Usuario_Id")
        obj.NomeCompleto = dr("NomeCompleto")
        obj.Senha = IIf(dr("Senha") Is DBNull.Value, "", dr("Senha"))
        obj.Empresa_Id = IIf(dr("Empresa") Is DBNull.Value, "", dr("Empresa"))
        obj.EnderecoEmpresa = IIf(dr("EndEmpresa") Is DBNull.Value, -1, dr("EndEmpresa"))
        obj.MenuDeAcesso = IIf(dr("MenuDeAcesso") Is DBNull.Value, "", dr("MenuDeAcesso"))
        obj.Email = IIf(dr("Email") Is DBNull.Value, "", dr("Email"))
        obj.AcessoUnidade = IIf(dr("AcessoUnidade") Is DBNull.Value, "", dr("AcessoUnidade"))
        obj.AcessoEmpresa = IIf(dr("AcessoEmpresa") Is DBNull.Value, "", dr("AcessoEmpresa"))
        obj.AcessoEnderecoEmpresa = IIf(dr("AcessoEndEmpresa") Is DBNull.Value, -1, dr("AcessoEndEmpresa"))
        obj.TrocarSenha = String.IsNullOrWhiteSpace(obj.Senha)
        obj.ImprimirDanfe = dr("ImprimirDanfe")
        obj.UsuarioAtivo = dr("Ativo")
        obj.LiberaEmpresa = dr("LiberaEmpresa")
        obj.TrocaEmpresa = dr("TrocaEmpresa")
        obj.LiberaJanela = dr("LiberaJanela")
        obj.DataInicioDashboard = dr("DataInicioDashboard")
        obj.DataFimDashboard = dr("DataFimDashboard")
        obj.SeguimentoDashboard = dr("SeguimentoDashboard")
        obj.EmpresasDashboard = dr("EmpresasDashboard")

        Return obj
    End Function

#End Region

#Region "Métodos"

    Public Shared Function Carregar(ByVal login As String) As Usuario
        Using cnn As New SqlConnection(UsuarioServidor.Conexao)
            cnn.Open()
            Dim strSQL As String = "SELECT  Usuario_Id, NomeCompleto, Senha, Empresa, EndEmpresa," & vbCrLf &
                                   "        MenuDeAcesso, Email, AcessoUnidade, AcessoEmpresa," & vbCrLf &
                                   "        AcessoEndEmpresa, SenhaMacros, isnull(Ativo,0) AS Ativo," & vbCrLf &
                                    "        Cliente, EndCliente, isnull(ImprimirDanfe,0) AS ImprimirDanfe," & vbCrLf &
                                   "         ISNULL(LiberaEmpresa, 0) as LiberaEmpresa, " & vbCrLf &
                                   "         ISNULL(TrocaEmpresa, 0) as TrocaEmpresa, " & vbCrLf &
                                   "         ISNULL(LiberaJanela, 0) as LiberaJanela," & vbCrLf &
                                   "         ISNULL(DataInicioDashboard, DATEADD(DAY, -30, GETDATE())) AS DataInicioDashboard," & vbCrLf &
                                   "         ISNULL(DataFimDashboard, GETDATE()) AS DataFimDashboard, " & vbCrLf &
                                   "        ISNULL(SeguimentoDashboard, '') AS SeguimentoDashboard, " & vbCrLf &
                                   "        ISNULL(EmpresasDashboard, 0) AS EmpresasDashboard " & vbCrLf &
                                   "FROM    Usuarios " &
                                   "WHERE   Usuario_Id = @Usuario_Id"
            Using cmd As New SqlCommand(strSQL, cnn)
                cmd.Parameters.Add("@Usuario_Id", SqlDbType.NVarChar, 100).Value = login
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                Dim obj As New Usuario()
                If dr.Read() Then obj = Mapper(dr)
                Return obj
            End Using
        End Using
    End Function

    Public Shared Function Validar(ByVal login As String, ByVal senha As String, Optional ByVal Supervisor As Boolean = False) As Boolean
        Using cnn As New SqlConnection(UsuarioServidor.Conexao)
            cnn.Open()
            Dim obj As New Usuario()
            Dim strSQL As String = "SELECT  Usuario_Id, NomeCompleto, Senha, Empresa, EndEmpresa," & vbCrLf &
                                   "        MenuDeAcesso, Email, AcessoUnidade, AcessoEmpresa," & vbCrLf &
                                   "        AcessoEndEmpresa, SenhaMacros, isnull(Ativo,0) AS Ativo," & vbCrLf &
                                   "        Cliente, EndCliente, isnull(ImprimirDanfe,0) AS ImprimirDanfe," & vbCrLf &
                                   "        ISNULL(LiberaEmpresa, 0) as LiberaEmpresa, " & vbCrLf &
                                   "        ISNULL(TrocaEmpresa, 0) as TrocaEmpresa, " & vbCrLf &
                                   "        ISNULL(LiberaJanela, 0) as LiberaJanela," & vbCrLf &
                                   "        ISNULL(DataInicioDashboard, DATEADD(DAY, -30, GETDATE())) AS DataInicioDashboard," & vbCrLf &
                                   "        ISNULL(DataFimDashboard, GETDATE()) AS DataFimDashboard, " & vbCrLf &
                                   "        ISNULL(SeguimentoDashboard, '') AS SeguimentoDashboard, " & vbCrLf &
                                   "        ISNULL(EmpresasDashboard, 0) AS EmpresasDashboard " & vbCrLf &
                                   "FROM    Usuarios " &
                                   "WHERE   UPPER(Usuario_Id) = UPPER(@Usuario_Id)"

            Using cmd As New SqlCommand(strSQL, cnn)
                cmd.Parameters.Add("@Usuario_Id", SqlDbType.NVarChar, 100).Value = login
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If dr.Read() Then obj = Mapper(dr)
            End Using

            If Not String.IsNullOrWhiteSpace(obj.Senha) OrElse obj.TrocarSenha Then
                If obj.TrocarSenha Then
                    Return True
                End If

                Dim byteHash As Byte() = Funcoes.ConvertStringToByteArray(senha)
                Dim byteMD5 As Byte() = New MD5CryptoServiceProvider().ComputeHash(byteHash)

                If Not Supervisor Then
                    Usuario = obj
                    CodigoEmpresa = IIf(Not String.IsNullOrWhiteSpace(obj.Empresa_Id), obj.Empresa_Id, "")
                    EnderecoEmpresa = IIf(Not String.IsNullOrWhiteSpace(obj.EnderecoEmpresa), obj.EnderecoEmpresa, "")
                    MenuDeAcesso = IIf(Not String.IsNullOrWhiteSpace(obj.MenuDeAcesso), obj.MenuDeAcesso, "Index")
                    NomeUsuario = UCase(Trim(obj.Usuario_Id))
                    ImprimirDanfe = obj.ImprimirDanfe
                    UsuarioAtivo = obj.UsuarioAtivo
                    LiberaEmpresa = obj.LiberaEmpresa
                    TrocaEmpresa = obj.TrocaEmpresa
                    LiberaJanela = obj.LiberaJanela
                    DataInicioDashboard = obj.DataInicioDashboard
                    DataFimDashboard = obj.DataFimDashboard
                    SeguimentoDashboard = obj.SeguimentoDashboard
                    EmpresasDashboard = obj.EmpresasDashboard
                    AcessoCusto = "CapturaDeDados" & ".aspx"

                End If

                Return BitConverter.ToString(byteMD5).ToUpper = obj.Senha.ToUpper() AndAlso VerificarEmpresa()
            End If
            Return False
        End Using
    End Function

    Public Shared Function TrocarSenha(ByVal login As String, ByVal senha As String) As Boolean
        Using cnn As New SqlConnection(UsuarioServidor.Conexao)
            cnn.Open()
            Dim obj As New Usuario()
            Dim strSQL As String = "SELECT  Usuario_Id, NomeCompleto, Senha, Empresa, EndEmpresa," & vbCrLf &
                                   "        MenuDeAcesso, Email, AcessoUnidade, AcessoEmpresa," & vbCrLf &
                                   "        AcessoEndEmpresa, SenhaMacros, isnull(Ativo,0) AS Ativo," & vbCrLf &
                                   "        Cliente, EndCliente, isnull(ImprimirDanfe,0) AS ImprimirDanfe," & vbCrLf &
                                   "        ISNULL(LiberaEmpresa, 0) as LiberaEmpresa, " & vbCrLf &
                                   "        ISNULL(TrocaEmpresa, 0) as TrocaEmpresa, " & vbCrLf &
                                   "        ISNULL(LiberaJanela, 0) as LiberaJanela," & vbCrLf &
                                   "        ISNULL(DataInicioDashboard, DATEADD(DAY, -30, GETDATE())) AS DataInicioDashboard," & vbCrLf &
                                   "        ISNULL(DataFimDashboard, GETDATE()) AS DataFimDashboard, " & vbCrLf &
                                   "        ISNULL(SeguimentoDashboard, '') AS SeguimentoDashboard, " & vbCrLf &
                                   "        ISNULL(EmpresasDashboard, 0) AS EmpresasDashboard " & vbCrLf &
                                   "FROM    Usuarios " &
                                   "WHERE   UPPER(Usuario_Id) = UPPER(@Usuario_Id)"

            Using cmd As New SqlCommand(strSQL, cnn)
                cmd.Parameters.Add("@Usuario_Id", SqlDbType.NVarChar, 100).Value = login
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If dr.Read() Then obj = Mapper(dr)
            End Using

            If String.IsNullOrWhiteSpace(obj.Senha) OrElse obj.TrocarSenha Then
                Return True
            End If
            Return False
        End Using
    End Function

    Public Shared Function AtualizarSenha(ByVal Usuario As String, ByVal Senha As String) As Boolean
        Using cnn As New SqlConnection(UsuarioServidor.Conexao)
            cnn.Open()

            Dim strSQL As String = "UPDATE  Usuarios " &
                                   "Set     Senha       = @Senha " &
                                   "WHERE   Usuario_Id  = @Usuario_Id"

            Using trans As SqlTransaction = cnn.BeginTransaction()
                Try
                    Using cmd As New SqlCommand(strSQL, cnn, trans)
                        Dim byteHash As Byte() = Funcoes.ConvertStringToByteArray(Senha)
                        Dim byteMD5 As Byte() = New MD5CryptoServiceProvider().ComputeHash(byteHash)

                        Dim strSenhaNova As String = BitConverter.ToString(byteMD5).ToUpper

                        cmd.Parameters.Add("@Senha", SqlDbType.NVarChar, 100).Value = strSenhaNova
                        cmd.Parameters.Add("@Usuario_Id", SqlDbType.NVarChar, 100).Value = Usuario

                        Dim intAfetados As Integer = cmd.ExecuteNonQuery()

                        If intAfetados = 1 Then
                            trans.Commit()
                            Return True
                        Else
                            trans.Rollback()
                            Throw New Exception("Usuário não encontrado para atualizar a senha!")
                            Return False
                        End If
                    End Using
                Catch ex As Exception
                    trans.Rollback()
                    Throw New Exception("Problemas ao atualizar a senha!", ex)
                    Return False
                End Try
            End Using
        End Using
    End Function

    Public Shared Function VerificarEmpresa() As Boolean
        Dim Sql = " Select Cliente_Id, Endereco_Id, Nome, Endereco, Cidade, Estado, Reduzido, Imagem " & vbCrLf &
              "  FROM Clientes" & vbCrLf &
              "  WHERE Cliente_Id = '" & HttpContext.Current.Session("ssEmpresa") & "'" & vbCrLf &
              "  AND Endereco_Id = " & HttpContext.Current.Session("ssEndEmpresa") & ""

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(Sql, "Empresa")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            HttpContext.Current.Session("ssEmpresa") = row("Cliente_Id")
            HttpContext.Current.Session("ssEndEmpresa") = row("Endereco_Id")
            HttpContext.Current.Session("ssNomeEmpresa") = UCase(row("Nome"))
            HttpContext.Current.Session("ssEnderecoEmpresa") = UCase(row("Endereco"))
            HttpContext.Current.Session("ssCidadeEmpresa") = UCase(row("Cidade"))
            HttpContext.Current.Session("ssEstadoEmpresa") = UCase(row("Estado"))
            HttpContext.Current.Session("ssReduzidoEmpresa") = UCase(row("Reduzido"))
            HttpContext.Current.Session("ssImagemEmpresa") = UCase(row("Imagem"))
            HttpContext.Current.Session("ssAcessoCusto") = "CapturaDeDados" & ".aspx"

            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Sub CarregarInformacaoParaUsuarioServidor(ByVal login As String)
        Using cnn As New SqlConnection(UsuarioServidor.Conexao)
            cnn.Open()
            Dim obj As New Usuario()
            Dim strSQL As String = "SELECT  Usuario_Id, NomeCompleto, Senha, Empresa, EndEmpresa," & vbCrLf &
                                   "        MenuDeAcesso, Email, AcessoUnidade, AcessoEmpresa," & vbCrLf &
                                   "        AcessoEndEmpresa, SenhaMacros, isnull(Ativo,0) AS Ativo," & vbCrLf &
                                   "        Cliente, EndCliente, isnull(ImprimirDanfe,0) AS ImprimirDanfe," & vbCrLf &
                                   "        ISNULL(LiberaEmpresa, 0) as LiberaEmpresa, " & vbCrLf &
                                   "        ISNULL(TrocaEmpresa, 0) as TrocaEmpresa, " & vbCrLf &
                                   "        ISNULL(LiberaJanela, 0) as LiberaJanela," & vbCrLf &
                                   "        ISNULL(DataInicioDashboard, DATEADD(DAY, -30, GETDATE())) AS DataInicioDashboard," & vbCrLf &
                                   "        ISNULL(DataFimDashboard, GETDATE()) AS DataFimDashboard, " & vbCrLf &
                                   "        ISNULL(SeguimentoDashboard, '') AS SeguimentoDashboard, " & vbCrLf &
                                   "        ISNULL(EmpresasDashboard, 0) AS EmpresasDashboard " & vbCrLf &
                                   "FROM    Usuarios " &
                                   "WHERE   UPPER(Usuario_Id) = UPPER(@Usuario_Id)"

            Using cmd As New SqlCommand(strSQL, cnn)
                cmd.Parameters.Add("@Usuario_Id", SqlDbType.NVarChar, 100).Value = login
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If dr.Read() Then obj = Mapper(dr)
            End Using

            Usuario = obj
            CodigoEmpresa = IIf(Not String.IsNullOrWhiteSpace(obj.Empresa_Id), obj.Empresa_Id, "")
            EnderecoEmpresa = IIf(Not String.IsNullOrWhiteSpace(obj.EnderecoEmpresa), obj.EnderecoEmpresa, "")
            MenuDeAcesso = IIf(Not String.IsNullOrWhiteSpace(obj.MenuDeAcesso), obj.MenuDeAcesso, "Index")
            NomeUsuario = UCase(Trim(obj.Usuario_Id))
            ImprimirDanfe = obj.ImprimirDanfe
            UsuarioAtivo = obj.UsuarioAtivo
            LiberaEmpresa = obj.LiberaEmpresa
            TrocaEmpresa = obj.TrocaEmpresa
            LiberaJanela = obj.LiberaJanela
            DataInicioDashboard = obj.DataInicioDashboard
            DataFimDashboard = obj.DataFimDashboard
            SeguimentoDashboard = obj.SeguimentoDashboard
            EmpresasDashboard = obj.EmpresasDashboard
            AcessoCusto = "CapturaDeDados" & ".aspx"

        End Using
    End Sub

#End Region

End Class