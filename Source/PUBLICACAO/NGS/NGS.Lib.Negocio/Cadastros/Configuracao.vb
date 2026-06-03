Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Drawing
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class Configuracao
    Implements IBaseEntity

#Region "Propriedades"
    Private _IUD As String
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Private _Codigo As Integer
    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Private _Email As String
    Public Property Email() As String
        Get
            Return _Email
        End Get
        Set(ByVal value As String)
            _Email = value
        End Set
    End Property

    Private _Host As String
    Public Property Host() As String
        Get
            Return _Host
        End Get
        Set(ByVal value As String)
            _Host = value
        End Set
    End Property

    Private _Senha As String
    Public Property Senha() As String
        Get
            Return _Senha
        End Get
        Set(ByVal value As String)
            _Senha = value
        End Set
    End Property

    Private _Usuario As String
    Public Property Usuario() As String
        Get
            Return _Usuario
        End Get
        Set(ByVal value As String)
            _Usuario = value
        End Set
    End Property

    Private _Porta As String
    Public Property Porta() As String
        Get
            Return _Porta
        End Get
        Set(ByVal value As String)
            _Porta = value
        End Set
    End Property

    Private _Ssl As Boolean
    Public Property Ssl() As Boolean
        Get
            Return _Ssl
        End Get
        Set(ByVal value As Boolean)
            _Ssl = value
        End Set
    End Property

    Private _Credencial As Boolean
    Public Property Credenciail() As Boolean
        Get
            Return _Credencial
        End Get
        Set(ByVal value As Boolean)
            _Credencial = value
        End Set
    End Property

    Private _Cheque As String
    Public Property Cheque() As String
        Get
            Return _Cheque
        End Get
        Set(ByVal value As String)
            _Cheque = value
        End Set
    End Property

    Private _Balanca As String
    Public Property Balanca() As String
        Get
            Return _Balanca
        End Get
        Set(ByVal value As String)
            _Balanca = value
        End Set
    End Property

    Private _FilesServer As String
    Public Property FilesServer() As String
        Get
            Return _FilesServer
        End Get
        Set(ByVal value As String)
            _FilesServer = value
        End Set
    End Property
#End Region

#Region "Construtor"
    Public Sub New()
        Dim sql = "SELECT TOP 1 * FROM Configuracao"
        Dim db As New AcessaBanco()
        Dim ds As DataSet = db.ConsultaDataSet(sql, "Configuracoes")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            Me.Codigo = row("Configuracao_Id")
            Me.Email = row("Email")
            Me.Host = row("Host")
            Me.Senha = row("Senha")
            Me.Usuario = row("Usuario")
            Me.Porta = row("Porta")
            Me.Cheque = row("Cheque")
            Me.Balanca = row("Balanca")
            Me.FilesServer = row("FilesServer")
            Me.Ssl = CBool(row("Ssl"))
            Me.Credenciail = CBool(row("Credencial"))
        End If
    End Sub
#End Region

#Region "Métodos"
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
                Sql = " INSERT INTO Configuracao (Email, Host, Senha, Usuario, Porta, Ssl, Credencial, Cheque, Balanca, FilesServer) " & vbCrLf & _
                      " VALUES ('" & _Email & "','" & _Host & "','" & _Senha & "','" & _Usuario & "','" & _Porta & "','" & _Ssl & "', " & vbCrLf & _
                      "'" & _Credencial & "','" & _Cheque & "','" & _Balanca & "','" & _FilesServer & "')"
                Sqls.Add(Sql)

            Case "U"
                Sql = " UPDATE Configuracao SET" & vbCrLf & _
                      "     Email            =   '" & _Email & "'" & vbCrLf & _
                      "    ,Host             =   '" & _Host & "'" & vbCrLf & _
                      "    ,Senha            =   '" & _Senha & "'" & vbCrLf & _
                      "    ,Usuario          =   '" & _Usuario & "'" & vbCrLf & _
                      "    ,Porta            =   '" & _Porta & "'" & vbCrLf & _
                      "    ,Ssl              =   '" & _Ssl & "'" & vbCrLf & _
                      "    ,Credencial       =   '" & _Credencial & "'" & vbCrLf & _
                      "    ,Cheque           =   '" & _Cheque & "'" & vbCrLf & _
                      "    ,Balanca          =   '" & _Balanca & "'" & vbCrLf & _
                      "    ,FilesServer      =   '" & _FilesServer & "'" & vbCrLf & _
                      "  WHERE Configuracao_Id  =  '" & _Codigo & "'"
                Sqls.Add(Sql)

            Case "D"
                Sql = " DELETE Configuracao" & vbCrLf & _
                      "  WHERE Configuracao_Id = '" & _Codigo & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
