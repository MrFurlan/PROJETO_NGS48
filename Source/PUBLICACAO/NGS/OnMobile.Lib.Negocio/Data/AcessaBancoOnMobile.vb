Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Text

Public Class AcessaBancoOnMobile
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
        '_cnn.ConnectionString = ConfigurationManager.ConnectionStrings("baxi").ConnectionString
        '_cnn.ConnectionString = "Data Source=MRFURLAN; Initial Catalog=Baxi; User Id=sa; Password=pwd_ngs123"
        _cnn.ConnectionString = "Data Source=SRVBAXI; Initial Catalog=Baxi; User Id=sa; Password=BaxiH24q#p11Zxt"
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
                Throw New Exception("Registro esta sendo utilizado em outras tabelas! " & msg)
            Else
                Throw New Exception(ex.Message)
            End If
        Catch ex As Exception
            If _transaction IsNot Nothing Then
                _transaction.Rollback()
            End If
            Throw New Exception(ex.Message)
        Finally
            'strm.Close()
            'strm.Dispose()
            _cmd.Connection.Close()
        End Try
        Return True
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
            If ex.Message.Contains("A propriedade ConnectionString não foi inicializada") Then
                Throw New Exception(ex.Message)
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
            If _transaction IsNot Nothing Then
                _transaction.Rollback()
            End If
            Throw New Exception(ex.Message)
        Finally
            _cmd.Connection.Close()
        End Try
    End Function
#End Region
End Class
