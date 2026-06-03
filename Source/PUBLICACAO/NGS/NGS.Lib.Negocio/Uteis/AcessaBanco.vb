Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Configuration
Imports System.Environment
Imports System.IO
Imports System.Web
Imports NGS.Lib.Uteis

Public Class AcessaBanco

#Region "Fields"
    Private _cmd As New SqlCommand
    Private _cnn As SqlConnection
#End Region

#Region "Construtor"
    ''' <summary>
    '''0 - SERVIDOR DO USUARIO BANCO DO SISTEMA NGS - VAZIO
    '''1 - SERVIDOR DO USUARIO BANCO USUARIOS       - VAZIO
    '''2 - SERVIDOR DA EMPRESA DO PEDIDO BANCO NGS  - NOME DO SERVIDOR DA EMPRESA DO PEDIDO 
    ''' </summary>
    ''' <param name="pBanco"></param>
    ''' <param name="pNomeDoServidor"></param>
    ''' <remarks></remarks>
    Public Sub New(Optional ByVal pBanco As Integer = 0, Optional pNomeDoServidor As String = "")
        _cnn = New SqlConnection()

        If ConfigurationManager.ConnectionStrings("baxi") IsNot Nothing Then
            '_cnn.ConnectionString = "Data Source=MRFURLAN; Initial Catalog=Baxi; User Id=sa; Password=pwd_ngs123"
            _cnn.ConnectionString = "Data Source=SRVBAXI; Initial Catalog=Baxi; User Id=sa; Password=BaxiH24q#p11Zxt"
        Else
            Select Case pBanco
                Case 0 : _cnn.ConnectionString = UsuarioServidor.EnderecoLocal
                Case 1 : _cnn.ConnectionString = GetStringBancoUsuarios()
                Case 2 : _cnn.ConnectionString = GetStringBancoEmpresaPedido(pNomeDoServidor)
                Case 3 : _cnn.ConnectionString = "server=ajuda.ngssolucoes.com.br;User Id=ngs_ajuda;Persist Security Info=True;database=ngs_ajuda;pwd=tbqcp@1fg;port=3306;SslMode=None"
                    'Case 3 : _cnn.ConnectionString = "server=ajudangs.mysql.uhserver.com;User Id=ngs;Persist Security Info=True;database=ajudangs;password=tbqcp@1fg"
            End Select
        End If

        _cmd.Connection = _cnn
        _cmd.CommandType = CommandType.Text
        _cmd.CommandTimeout = 5000
    End Sub
#End Region

#Region "Methods"

    Public Function GravaBanco(ByVal sql As String) As Boolean
        Dim Sqls As New ArrayList
        Sqls.Add(sql)
        Return GravaBanco(Sqls)
    End Function

    Public Function GravaBanco(ByVal Sqls As ArrayList) As Boolean
        If Not UsuarioServidor.KeyCodeActive Then Throw New Exception("Sistema com chave de licença expirada. Entre em contato com o suporte.")
        Dim _transaction As SqlTransaction = Nothing
        'Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/gravabanco.txt"))
        Try
            _cmd.Connection.Open()
            _transaction = _cnn.BeginTransaction(IsolationLevel.ReadCommitted)
            _cmd.Transaction = _transaction
            For index As Integer = 0 To Sqls.Count - 1
                _cmd.CommandText = CStr(Sqls.Item(index))
                'strm.WriteLine(_cmd.CommandText)
                _cmd.CommandTimeout = 5000
                _cmd.ExecuteNonQuery()
            Next
            _transaction.Commit()
        Catch ex As SqlException
            If _transaction IsNot Nothing Then
                _transaction.Rollback()
            End If
            'Violação de primary key.
            If ex.Number.Equals(2627) Then
                HttpContext.Current.Session("ssMessage") = "Registro já cadastrado!"
                Throw New Exception("Registro já cadastrado!")
            ElseIf ex.Number.Equals(547) Then
                Dim pos As Integer = ex.Message.IndexOf("FK_")
                Dim msg As String = ""
                If pos <> -1 Then
                    msg = ex.Message.Substring(pos)
                    msg = msg.Split(" ")(0)
                Else
                    msg = ex.Message.Trim("""", "").Trim(".", "")
                End If
                HttpContext.Current.Session("ssMessage") = "Registro esta sendo utilizado em outras tabelas! " & msg
                Throw New Exception("Registro esta sendo utilizado em outras tabelas! " & msg)
            Else
                Funcoes.WriteLogFile(Me.GetType().Name, GetCurrentMethodName(), ex.Message)
                HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message)
                Throw New Exception(ex.Message)
            End If
        Catch ex As Exception
            If _transaction IsNot Nothing Then
                _transaction.Rollback()
            End If
            Funcoes.WriteLogFile(Me.GetType().Name, GetCurrentMethodName(), ex.Message)
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message)
            Throw New Exception(ex.Message)
        Finally
            'strm.Close()
            'strm.Dispose()
            _cmd.Connection.Close()
        End Try
        Return True
    End Function

    Function GetCurrentMethodName() As String

        Dim stack As New System.Diagnostics.StackFrame(1)

        Return stack.GetMethod().Name

    End Function

    Public Function ExecuteNonQuery(ByVal sql As String) As Boolean
        Dim rowsNumber As Integer = 0
        Dim _transaction As SqlTransaction = Nothing
        Dim errorMsg As New StringBuilder()
        Try
            _cmd.Connection.Open()
            _transaction = _cnn.BeginTransaction(IsolationLevel.ReadCommitted)
            _cmd.Transaction = _transaction
            _cmd.CommandText = sql
            rowsNumber = _cmd.ExecuteNonQuery()
            _transaction.Commit()
        Catch ex As Exception
            If _transaction IsNot Nothing Then
                _transaction.Rollback()
            End If
            Funcoes.WriteLogFile(Me.GetType().Name, GetCurrentMethodName(), ex.Message)
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message)
            Throw New Exception(ex.Message)
        Finally
            _cmd.Connection.Close()
        End Try
        Return rowsNumber > 0
    End Function

    Public Function ConsultaDataSet(ByVal sql As String, ByVal tableName As String, Optional TipoComando As System.Data.CommandType = CommandType.Text, Optional ParametrosSP As System.Data.SqlClient.SqlParameterCollection = Nothing) As DataSet
        Try
            Dim adp As New SqlDataAdapter
            Dim ds As New DataSet
            _cmd.Connection.Open()
            If Not String.IsNullOrWhiteSpace(sql) Then
                If TipoComando = CommandType.StoredProcedure Then
                    _cmd.Parameters.Clear()
                    For i = 0 To ParametrosSP.Count - 1
                        _cmd.Parameters.AddWithValue(ParametrosSP.Item(i).ParameterName, ParametrosSP.Item(i).Value)
                    Next
                End If

                _cmd.CommandType = TipoComando
                _cmd.CommandText = sql
                _cmd.CommandTimeout = 5000
                adp.SelectCommand = _cmd
                adp.Fill(ds, tableName)
            End If
            Return ds
        Catch ex As Exception
            Funcoes.WriteLogFile(Me.GetType().Name, GetCurrentMethodName(), ex.Message)
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message)

            If ex.Message.Contains("A propriedade ConnectionString não foi inicializada") Then
                HttpContext.Current.Response.Redirect("~/Logout.aspx", False)
                'Throw New Exception(ex.Message)
            Else
                Throw New Exception(ex.Message)
            End If

        Finally
            _cmd.Connection.Close()
        End Try
    End Function

    Public Function ConsultaReport(ByVal sql As String, ByVal tableName As String) As DataSet
        Dim adp As New SqlDataAdapter
        Dim _transaction As SqlTransaction = Nothing
        Dim ds As New DataSet
        Try
            If Not String.IsNullOrWhiteSpace(sql) Then
                _cmd.Connection.Open()
                _transaction = _cnn.BeginTransaction(IsolationLevel.ReadUncommitted)
                _cmd.Transaction = _transaction

                _cmd.CommandText = sql
                _cmd.CommandTimeout = 5000
                adp.SelectCommand = _cmd
                adp.Fill(ds, tableName)
                _transaction.Commit()
            End If
            Return ds
        Catch ex As Exception
            Funcoes.WriteLogFile(Me.GetType().Name, GetCurrentMethodName(), ex.Message)
            If _transaction IsNot Nothing Then
                _transaction.Rollback()
            End If
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message)
            Throw New Exception(ex.Message)
        Finally
            _cmd.Connection.Close()
        End Try
    End Function

    Public Shared Function GetStringBancoUsuarios() As String
        Dim strArquivo As String = GetEnvironmentVariable("windir") & "\msedbms.dll"
        Dim strLinha As String
        Dim strStringConexao As String = String.Empty
        Dim strTipo As String = String.Empty
        Dim strEndereco As String = String.Empty
        Dim chrLetra As Char
        Dim intPos As Integer = 0
        Dim shtLetra As Short
        Dim blnTipo As Boolean = False

        If Not File.Exists(strArquivo) Then
            Throw New Exception("Arquivo de configuração do banco de dados não existe!")
        Else
            Using strm As New StreamReader(strArquivo)
                While strm.Read()
                    intPos += 1
                    blnTipo = False
                    strLinha = strm.ReadLine
                    If Not String.IsNullOrEmpty(strLinha) Then
                        For intPosLetra As Integer = 1 To strLinha.Length
                            chrLetra = Mid(strLinha, intPosLetra, 1)
                            shtLetra = Convert.ToInt16(chrLetra) + 9
                            'If Not blnTipo Then blnTipo = (Convert.ToChar(shtLetra) = "#")
                            'If blnTipo Then strTipo &= Convert.ToChar(shtLetra) Else strEndereco &= Convert.ToChar(shtLetra)

                            strEndereco &= Convert.ToChar(shtLetra)
                        Next
                    Else
                        Exit While
                    End If
                End While
            End Using

            Dim x As String = StrReverse(strEndereco)
            Dim pos As Integer = x.IndexOf("#")
            x = StrReverse(x.Substring(pos + 1, Len(strEndereco) - pos - 1))

            If Not String.IsNullOrEmpty(x) Then strStringConexao = x.Replace("Provider=sqloledb;", "")
        End If

        Return strStringConexao
    End Function

    Public Function GetStringBancoEmpresaPedido(ByVal servidor As String) As String
        Dim ds As DataSet
        Try
            Dim sql As String = ""
            sql = "select UsuarioDoBanco, SenhaDoBanco " & vbCrLf & _
                  "  from Bancos" & vbCrLf & _
                  " Where Banco_Id       = '" & UsuarioServidor.BancoDeDados & "'" & vbCrLf & _
                  "   and HostDoServidor = '" & servidor & "'" & vbCrLf

            Dim B As New AcessaBanco(1)
            ds = B.ConsultaDataSet(sql, "ConexaoRemota")

            If ds.Tables(0).Rows.Count = 0 Then Return ""
        Catch ex As Exception
            Return ""
        End Try

        Return "Data Source=" & servidor & ";Initial Catalog=" & UsuarioServidor.BancoDeDados & ";User Id=" & ds.Tables(0).Rows(0)("UsuarioDoBanco") & ";Password=" & FuncoesStrings.DecodificarDe64Bits(ds.Tables(0).Rows(0)("SenhaDoBanco")) & ";"
    End Function

    Public Function GetConnectionString() As String
        Return _cnn.ConnectionString
    End Function

#End Region

End Class