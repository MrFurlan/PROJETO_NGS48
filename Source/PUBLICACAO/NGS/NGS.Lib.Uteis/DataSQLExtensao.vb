Imports System.Data.SqlClient

Public Class DataSQLExtensao
    Public Shared Function CommandAsSqlText(ByVal command As SqlCommand) As String
        Dim query = command.CommandText
        Dim erro = String.Empty
        Try
            For Each p As SqlParameter In command.Parameters
                erro = String.Format("{0} - {1} - {2}", p.SqlDbType, p.ParameterName, p.Value.ToString())
                Select Case p.SqlDbType
                    Case SqlDbType.VarChar
                        query = Replace(query, p.ParameterName, ToSQLString(p.Value.ToString()))
                    Case SqlDbType.Int
                        query = Replace(query, p.ParameterName, ToSQLInteger(p.Value.ToString()))
                    Case SqlDbType.DateTime
                        query = Replace(query, p.ParameterName, ToSQLDateTime(p.Value.ToString()))
                    Case SqlDbType.Decimal
                        query = Replace(query, p.ParameterName, ToSQLDecimal(p.Value.ToString()))
                        Exit Select
                End Select

            Next
        Catch ex As Exception
            erro = String.Format("{0} - {1}", erro, ex.Message)
        End Try

        Return query

    End Function

    Protected Shared Function ToSQLString(value) As String
        Try
            Return String.Format("'{0}'", value)
        Catch ex As Exception
            Dim e = ex.Message
        End Try
    End Function

    Protected Shared Function ToSQLInteger(value) As String
        Try
            Return Integer.Parse(value)
        Catch ex As Exception
            Dim e = ex.Message
        End Try

    End Function

    Protected Shared Function ToSQLDateTime(value) As String
        Try
            Return String.Format("'{0}'", DateTime.Parse(value))
        Catch ex As Exception
            Dim e = ex.Message
        End Try
    End Function

    Protected Shared Function ToSQLDecimal(value) As String
        Try
            Return String.Format("{0}", Decimal.Parse(value))
        Catch ex As Exception
            Dim e = ex.Message
        End Try
    End Function
End Class
