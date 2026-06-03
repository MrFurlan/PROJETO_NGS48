Public Class ExceptionLog
    Implements IBaseEntity

    Private _ex As Exception
    Private _time As DateTime
    Private _mensagem As String

    Public Property Ex As Exception
        Get
            Return _ex
        End Get
        Set(value As Exception)
            _ex = value
        End Set
    End Property

    Public Property Time As Date
        Get
            Return _time
        End Get
        Set(value As Date)
            _time = value
        End Set
    End Property

    Public Function LogException()

        Dim sql As String = "INSERT INTO ExceptionLogs" &
        "(ExceptionMessage, StackTrace, LogDate) VALUES " &
        "('" & _ex.Message.Replace("'", """") & "','" & _ex.StackTrace & "','" & _time.ToString("yyyy-MM-dd HH:mm:ss") & "')"

        Dim Banco As New AcessaBanco

        If banco.GravaBanco(sql) Then
            _mensagem = "log de exceção foi incluido nos logs do banco, consultar tabela ExceptionLogs"
        Else
            _mensagem = "erro na gravação do log"
        End If

        Return _mensagem
    End Function
End Class