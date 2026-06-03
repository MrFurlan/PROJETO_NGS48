Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListBalancaEletronica
    Inherits List(Of BalancaEletronica)

#Region "Contrutor"
    Public Sub New()
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT Balanca_Servidor, Balanca_Ip, Balanca_TCPIP, Balanca_Tipo, Eletronica, ReadBuffer, " & vbCrLf & _
              "       ReceivedBytes, PortName, BaudRate, Parity, DataBits, StopBits, NumeroDeCopias " & vbCrLf & _
              "  FROM BalancaEletronica "

        ds = Banco.ConsultaDataSet(sql, "BalancaEletronica")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim BE As New BalancaEletronica
            BE.BalancaServidor = row("Balanca_Servidor")
            BE.BalancaIp = row("Balanca_Ip")
            BE.BalancaTCPIP = row("Balanca_TCPIP")
            BE.BalancaTipo = row("Balanca_Tipo")
            BE.Eletronica = row("Eletronica")
            BE.ReadBuffer = row("ReadBuffer")
            BE.ReceivedBytes = row("ReceivedBytes")
            BE.PortName = row("PortName")
            BE.BaudRate = row("BaudRate")
            BE.Parity = row("Parity")
            BE.DataBits = row("DataBits")
            BE.StopBits = row("StopBits")
            BE.NumeroDeCopias = row("NumeroDeCopias")
            Me.Add(BE)
        Next
    End Sub

    Public Sub New(ByVal Ip As String, ByVal Usuario As String)
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT balanca.Balanca_Servidor, Balanca_Ip, Balanca_TCPIP, Balanca_Tipo, Eletronica, ReadBuffer," &
              "ReceivedBytes, PortName, BaudRate, Parity, DataBits, StopBits, NumeroDeCopias " &
              "FROM BalancaEletronica AS balanca " &
              "INNER JOIN BalancasXUsuarios ON BalancasXUsuarios.Balanca_Servidor = balanca.Balanca_Servidor " &
              "WHERE BalancasXUsuarios.BalancaUsuario_IP = '" & Ip & "' " &
              "AND BalancasXUsuarios.Usuario_Id = '" & Usuario & "'"

        ds = Banco.ConsultaDataSet(sql, "BalancaEletronica")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim BE As New BalancaEletronica
            BE.BalancaServidor = row("Balanca_Servidor")
            BE.BalancaIp = row("Balanca_Ip")
            BE.BalancaTCPIP = row("Balanca_TCPIP")
            BE.BalancaTipo = row("Balanca_Tipo")
            BE.Eletronica = row("Eletronica")
            BE.ReadBuffer = row("ReadBuffer")
            BE.ReceivedBytes = row("ReceivedBytes")
            BE.PortName = row("PortName")
            BE.BaudRate = row("BaudRate")
            BE.Parity = row("Parity")
            BE.DataBits = row("DataBits")
            BE.StopBits = row("StopBits")
            BE.NumeroDeCopias = row("NumeroDeCopias")
            Me.Add(BE)
        Next
    End Sub

#End Region

End Class

<Serializable()> _
Public Class BalancaEletronica
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal NomeBalanca As String)
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT Balanca_Servidor, Balanca_Ip, Balanca_TCPIP, Balanca_Tipo, Eletronica, ReadBuffer, " & vbCrLf & _
              "       ReceivedBytes, PortName, BaudRate, Parity, DataBits, StopBits, NumeroDeCopias " & vbCrLf & _
              "  FROM BalancaEletronica " & vbCrLf & _
              " Where Balanca_Servidor = '" & NomeBalanca & "'"

        ds = Banco.ConsultaDataSet(sql, "BalancaEletronica")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        _BalancaServidor = row("Balanca_Servidor")
        _BalancaIp = row("Balanca_Ip")
        _BalancaTCPIP = row("Balanca_TCPIP")
        _BalancaTipo = row("Balanca_Tipo")
        _Eletronica = row("Eletronica")
        _ReadBuffer = row("ReadBuffer")
        _ReceivedBytes = row("ReceivedBytes")
        _PortName = row("PortName")
        _BaudRate = row("BaudRate")
        _Parity = row("Parity")
        _DataBits = row("DataBits")
        _StopBits = row("StopBits")
        _NumeroDeCopias = row("NumeroDeCopias")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _BalancaServidor As String = ""
    Private _BalancaIp As String = ""
    Private _BalancaTCPIP As Integer
    Private _BalancaTipo As String = ""
    Private _Eletronica As String = ""
    Private _ReadBuffer As Integer
    Private _ReceivedBytes As Integer
    Private _PortName As String = ""
    Private _BaudRate As Integer
    Private _Parity As Integer
    Private _DataBits As Integer
    Private _StopBits As Integer
    Private _NumeroDeCopias As Integer
#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property BalancaServidor() As String
        Get
            Return _BalancaServidor
        End Get
        Set(ByVal value As String)
            _BalancaServidor = value
        End Set
    End Property

    Public Property BalancaIp() As String
        Get
            Return _BalancaIp
        End Get
        Set(ByVal value As String)
            _BalancaIp = value
        End Set
    End Property

    Public Property BalancaTCPIP() As Integer
        Get
            Return _BalancaTCPIP
        End Get
        Set(ByVal value As Integer)
            _BalancaTCPIP = value
        End Set
    End Property

    Public Property BalancaTipo() As String
        Get
            Return _BalancaTipo
        End Get
        Set(ByVal value As String)
            _BalancaTipo = value
        End Set
    End Property

    Public Property Eletronica() As String
        Get
            Return _Eletronica
        End Get
        Set(ByVal value As String)
            _Eletronica = value
        End Set
    End Property

    Public Property ReadBuffer() As Integer
        Get
            Return _ReadBuffer
        End Get
        Set(ByVal value As Integer)
            _ReadBuffer = value
        End Set
    End Property

    Public Property ReceivedBytes() As Integer
        Get
            Return _ReceivedBytes
        End Get
        Set(ByVal value As Integer)
            _ReceivedBytes = value
        End Set
    End Property

    Public Property PortName() As String
        Get
            Return _PortName
        End Get
        Set(ByVal value As String)
            _PortName = value
        End Set
    End Property

    Public Property BaudRate() As Integer
        Get
            Return _BaudRate
        End Get
        Set(ByVal value As Integer)
            _BaudRate = value
        End Set
    End Property

    Public Property Parity() As Integer
        Get
            Return _Parity
        End Get
        Set(ByVal value As Integer)
            _Parity = value
        End Set
    End Property

    Public Property DataBits() As Integer
        Get
            Return _DataBits
        End Get
        Set(ByVal value As Integer)
            _DataBits = value
        End Set
    End Property

    Public Property StopBits() As Integer
        Get
            Return _StopBits
        End Get
        Set(ByVal value As Integer)
            _StopBits = value
        End Set
    End Property

    Public Property NumeroDeCopias() As Integer
        Get
            Return _NumeroDeCopias
        End Get
        Set(ByVal value As Integer)
            _NumeroDeCopias = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"

                Sql = "INSERT INTO BalancaEletronica (Balanca_Servidor, Balanca_Ip, Balanca_TCPIP, Balanca_Tipo, " & vbCrLf & _
                      "       Eletronica, ReadBuffer, ReceivedBytes, PortName, BaudRate, Parity, DataBits, StopBits, NumeroDeCopias)" & vbCrLf & _
                      "VALUES " & vbCrLf & _
                      "('" & _BalancaServidor & "','" & _BalancaIp & "', " & _BalancaTCPIP & ", '" & _BalancaTipo & "', " & vbCrLf & _
                      "'" & _Eletronica & "', " & _ReadBuffer & ", " & _ReceivedBytes & ", '" & _PortName & "', " & vbCrLf & _
                      "" & _BaudRate & ", " & _Parity & ", " & _DataBits & ", " & _StopBits & ", " & _NumeroDeCopias & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update BalancaEletronica set " & vbCrLf & _
                      "  Eletronica     = '" & _Eletronica & "'" & vbCrLf & _
                      " ,ReadBuffer     = " & _ReadBuffer & vbCrLf & _
                      " ,ReceivedBytes  = " & _ReceivedBytes & vbCrLf & _
                      " ,PortName       = '" & _PortName & "'" & vbCrLf & _
                      " ,BaudRate       = " & _BaudRate & vbCrLf & _
                      " ,Parity         = " & _Parity & vbCrLf & _
                      " ,DataBits       = " & _DataBits & vbCrLf & _
                      " ,StopBits       = " & _StopBits & vbCrLf & _
                      " ,NumeroDeCopias = " & _NumeroDeCopias & vbCrLf & _
                      " Where Balanca_Servidor = '" & _BalancaServidor & "'" & vbCrLf & _
                      "   and Balanca_Ip       = '" & _BalancaIp & "'" & vbCrLf & _
                      "   and Balanca_TCPIP    = " & _BalancaTCPIP & vbCrLf & _
                      "   and Balanca_Tipo     = '" & _BalancaTipo & "'"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE BalancaEletronica " & vbCrLf & _
                      " Where Balanca_Servidor = '" & _BalancaServidor & "'" & vbCrLf & _
                      "   and Balanca_Ip       = '" & _BalancaIp & "'" & vbCrLf & _
                      "   and Balanca_TCPIP    = " & _BalancaTCPIP & vbCrLf & _
                      "   and Balanca_Tipo     = '" & _BalancaTipo & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class