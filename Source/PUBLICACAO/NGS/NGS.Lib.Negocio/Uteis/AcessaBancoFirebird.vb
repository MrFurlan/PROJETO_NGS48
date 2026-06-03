Imports FirebirdSql.Data.FirebirdClient
Imports System.Web
Imports System.Data.SqlClient

Public Class AcessaBancoFirebird

#Region "Fields"
    Private _cmd As New FbCommand
    Private _cnn As FbConnection
#End Region

#Region "Construtor"
    Public Sub New()
        _cnn = New FbConnection
        _cnn.ConnectionString = "User=SYSDBA;Password=masterkey;Database=C:\Banco de Dados\CNRT-DADOS.FDB;DataSource=localhost;Port=3050;Dialect=3;Charset=NONE;Role=;Connection lifetime=15;Pooling=true;MinPoolSize=0;MaxPoolSize=50;Packet Size=8192;ServerType=0;"
        _cmd.Connection = _cnn
        _cmd.CommandType = CommandType.Text
        _cmd.CommandTimeout = 5000
    End Sub
#End Region


#Region "Methods"

    Public Function ConsultaDataSet(ByVal sql As String, ByVal tableName As String, Optional TipoComando As System.Data.CommandType = CommandType.Text, Optional ParametrosSP As System.Data.SqlClient.SqlParameterCollection = Nothing) As DataSet
        Try
            Dim adp As New FbDataAdapter
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

#End Region

End Class
