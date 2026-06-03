Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListBalancasXUsuarios
    Inherits List(Of BalancasXUsuarios)

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Servidor As String)
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT Balanca_Servidor, BalancaUsuario_Ip, Usuario_ID " & vbCrLf & _
              "  FROM BalancasXUsuarios " & vbCrLf

        If Servidor.Length > 0 Then
            sql &= "WHERE Balanca_Servidor = '" & Servidor & "'"
        End If

        ds = Banco.ConsultaDataSet(sql, "BalancasXUsuarios")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim BxU As New BalancasXUsuarios
            BxU.CodigoBalancaServidor = row("Balanca_Servidor")
            BxU.BalancaUsuarioIp = row("BalancaUsuario_Ip")
            BxU.NomeUsuario = row("Usuario_ID")
            Me.Add(BxU)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class BalancasXUsuarios
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal NomeBalanca As String, ByVal IpUsuario As String, ByVal NomeUsuario As String)
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT Balanca_Servidor, BalancaUsuario_Ip, Usuario_ID " & vbCrLf & _
              "  FROM BalancasXUsuarios " & vbCrLf & _
              " Where Balanca_Servidor  = '" & NomeBalanca & "'" & vbCrLf & _
              "   and BalancaUsuario_Ip = '" & IpUsuario & "'" & vbCrLf & _
              "   and Usuario_ID        = '" & NomeUsuario & "'"

        ds = Banco.ConsultaDataSet(sql, "BalancasXUsuarios")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        _CodigoBalancaServidor = row("Balanca_Servidor")
        _BalancaUsuarioIp = row("BalancaUsuario_Ip")
        _NomeUsuario = row("Usuario_ID")

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoBalancaServidor As String = ""
    Private _BalancaServidor As BalancaEletronica
    Private _BalancaUsuarioIp As String = ""
    Private _NomeUsuario As String = ""
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

    Public Property CodigoBalancaServidor() As String
        Get
            Return _CodigoBalancaServidor
        End Get
        Set(ByVal value As String)
            _CodigoBalancaServidor = value
        End Set
    End Property

    Public Property BalancaServidor() As BalancaEletronica
        Get
            If _BalancaServidor Is Nothing And _CodigoBalancaServidor.Length > 0 Then _BalancaServidor = New BalancaEletronica(_CodigoBalancaServidor)
            Return _BalancaServidor
        End Get
        Set(ByVal value As BalancaEletronica)
            _BalancaServidor = value
        End Set
    End Property

    Public Property BalancaUsuarioIp() As String
        Get
            Return _BalancaUsuarioIp
        End Get
        Set(ByVal value As String)
            _BalancaUsuarioIp = value
        End Set
    End Property

    Public Property NomeUsuario() As String
        Get
            Return _NomeUsuario
        End Get
        Set(ByVal value As String)
            _NomeUsuario = value
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

                Sql = "INSERT INTO BalancasXUsuarios (Balanca_Servidor, BalancaUsuario_Ip, Usuario_ID)" & vbCrLf & _
                      "VALUES " & vbCrLf & _
                      "('" & _CodigoBalancaServidor & "', '" & _BalancaUsuarioIp & "', '" & _NomeUsuario & "')"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE BalancasXUsuarios " & vbCrLf & _
                      " Where Balanca_Servidor  = '" & _CodigoBalancaServidor & "'" & vbCrLf & _
                      "   and BalancaUsuario_Ip = '" & _BalancaUsuarioIp & "'" & vbCrLf & _
                      "   and Usuario_ID        = '" & _NomeUsuario & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class