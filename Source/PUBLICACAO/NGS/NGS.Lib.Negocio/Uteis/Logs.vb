Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Logs
    Private _log_id As String
    Private _info As String
    Private _tipo As String

    Private _pathLog As String = "C:\NGS\Log"
    Private _LogFile As StreamWriter

    Public Property Log_id As String
        Get
            Return _log_id
        End Get
        Set(value As String)
            _log_id = value
        End Set
    End Property

    Public Property Tipo As String
        Get
            Return _tipo
        End Get
        Set(value As String)
            _tipo = value
        End Set
    End Property

    Private Function GravarArquivo()
        Try
            _LogFile = New StreamWriter(String.Format("{0}\{1}\{2}.data", _pathLog, _tipo, Log_id), True)
            _LogFile.WriteLine(_info)
            _LogFile.Flush()
            _LogFile.Close()
            Return True
        Catch Ex As Exception
            Return False
        End Try
    End Function

    Public Function Gravar()
        If GravarArquivo() Then
            Return True
        End If
        Return False
    End Function

    Public Sub GetLog(info As String)
        _log_id = UsuarioServidor.CodigoEmpresa
        _info = Now.ToString("yyyy/MM/dd HH:mm:ss") & " - " &
                             UsuarioServidor.NomeUsuario & " - " &
                             UsuarioServidor.CodigoEmpresa & " - " &
                             UsuarioServidor.NomeEmpresa & " - " &
                             UsuarioServidor.CidadeEmpresa & "/" & UsuarioServidor.EstadoEmpresa & " - " &
                             info
    End Sub
End Class