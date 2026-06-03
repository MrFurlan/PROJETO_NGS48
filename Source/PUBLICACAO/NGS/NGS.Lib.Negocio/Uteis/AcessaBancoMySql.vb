
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
Imports MySql.Data.MySqlClient


Public Class AcessaBancoMySql

#Region "Fields"
    Private _cmd As New MySqlCommand
    Private _cnn As MySqlConnection
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
        _cnn = New MySqlConnection()
        Select Case pBanco
            Case 0 : _cnn.ConnectionString = "server=ajuda.ngssolucoes.com.br;User Id=ngs_ajuda;Persist Security Info=True;database=ngs_ajuda;pwd=tbqcp@1fg;port=3306;SslMode=None"
                'Case 0 : _cnn.ConnectionString = "server=ajudangs.mysql.uhserver.com;User Id=ngs;Persist Security Info=True;database=ajudangs;password=tbqcp@1fg"
        End Select
        _cmd.Connection = _cnn
        _cmd.CommandType = CommandType.Text
        _cmd.CommandTimeout = 5000
    End Sub
#End Region

#Region "Methods"

    Public Function GravaBanco(ByVal sqlorproc As String, ByVal procedure As Boolean, ByVal param As Dictionary(Of String, Object)) As Boolean
        Dim _transaction As MySqlTransaction = Nothing
        Try
            _cmd.Connection.Open()
            If Not procedure Then
                _transaction = _cnn.BeginTransaction(IsolationLevel.ReadCommitted)
                _cmd.Transaction = _transaction
            Else : _cmd.CommandType = CommandType.StoredProcedure
            End If
            _cmd.CommandText = sqlorproc
            getParameters(sqlorproc, param, _cmd)
            _cmd.CommandTimeout = 5000
            _cmd.ExecuteNonQuery()

            If Not procedure Then _transaction.Commit()
        Catch ex As SqlException
            If Not procedure Then _transaction.Rollback()
            'Violação de primary key.
            If ex.Number.Equals(2627) Then
                HttpContext.Current.Session("ssMessage") = "Registro já cadastrado!"
                Throw New Exception("Registro já cadastrado!")
            ElseIf ex.Number.Equals(547) Then
                HttpContext.Current.Session("ssMessage") = "Registro esta sendo utilizado em outras tabelas!"
                Throw New Exception("Registro esta sendo utilizado em outras tabelas!")
            End If
            Throw New Exception(ex.Message)
        Catch ex As Exception
            If Not procedure Then _transaction.Rollback()
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message)
            Throw New Exception(ex.Message)
        Finally
            _cmd.Connection.Close()
        End Try
        Return True
    End Function

    Private Sub getParameters(ByVal proc As String, ByVal param As Dictionary(Of String, Object), ByRef cmd As MySqlCommand)
        If proc.Equals("sp_insertorupdate_processo") Then
            cmd.Parameters.AddWithValue("p_Processo_Id", param("Processo_Id"))
            cmd.Parameters.AddWithValue("p_descricao", param("Descricao"))
            cmd.Parameters.AddWithValue("p_manual", param("Manual"))
            cmd.Parameters.AddWithValue("p_dataAtualizacao", param("DataAtualizacao"))
            cmd.Parameters.AddWithValue("p_usuarioAlteracao", param("usuarioAlteracao"))
        End If
    End Sub

    Public Function ExecuteNonQuery(ByVal sql As String) As Boolean
        Dim rowsNumber As Integer = 0
        Dim _transaction As MySqlTransaction = Nothing
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
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message)
            Throw New Exception(ex.Message)
        Finally
            _cmd.Connection.Close()
        End Try
        Return rowsNumber > 0
    End Function

    Public Function ConsultaDataSet(ByVal sql As String, ByVal tableName As String, Optional TipoComando As System.Data.CommandType = CommandType.Text, Optional ParametrosSP As System.Data.SqlClient.SqlParameterCollection = Nothing) As DataSet
        Try
            Dim adp As New MySqlDataAdapter
            Dim ds As New DataSet
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
            HttpContext.Current.Session("ssMessage") = Funcoes.EliminarCaracteresEspeciais(ex.Message)
            Throw New Exception(ex.Message)
        Finally
            _cmd.Connection.Close()
        End Try
    End Function

    Public Function ConsultaReport(ByVal sql As String, ByVal tableName As String) As DataSet
        Dim adp As New MySqlDataAdapter
        Dim _transaction As MySqlTransaction = Nothing
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

        Return "Data Source=" & servidor & ";Initial Catalog=" & UsuarioServidor.BancoDeDados & ";User Id=" & ds.Tables(0).Rows(0)("UsuarioDoBanco") & ";Password=" & ds.Tables(0).Rows(0)("SenhaDoBanco") & ";"
    End Function

    Public Function GetConnectionString() As String
        Return _cnn.ConnectionString
    End Function

#End Region

End Class

