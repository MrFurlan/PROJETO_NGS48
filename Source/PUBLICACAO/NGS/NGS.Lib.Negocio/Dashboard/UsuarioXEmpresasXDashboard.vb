Imports System.Collections.Generic


<Serializable()>
Public Class ListUsuarioXEmpresasXDashboard
    Inherits List(Of UsuarioXEmpresasXDashboard)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pUsuario_Id As String)

        Dim Sql As String
        Sql = " SELECT UXEXD.Usuario_Id, UXEXD.Empresa_Id, UXEXD.EndEmpresa_Id " & vbCrLf &
              "   FROM UsuarioXEmpresasXDashboard As UXEXD " & vbCrLf &
              "  WHERE UXEXD.Usuario_Id ='" & pUsuario_Id & "'; "

        Dim Banco As New AcessaBanco
        Dim DsFaturas As DataSet = Banco.ConsultaDataSet(Sql, "UXEXD")

        For Each dr As DataRow In DsFaturas.Tables(0).Rows
            Dim usuario = New UsuarioXEmpresasXDashboard
            usuario.Usuario_Id = dr("Usuario_Id")
            usuario.Empresa_Id = dr("Empresa_Id")
            usuario.EndEmpresa_Id = dr("EndEmpresa_Id")

            Me.Add(usuario)

        Next

    End Sub

End Class

<Serializable()>
Public Class UsuarioXEmpresasXDashboard
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New()

    End Sub

#End Region

#Region "Fields"

    Private _IUD As String
    Private _Usuario_Id As String
    Private _Empresa_Id As String
    Private _EndEmpresa_Id As Integer

#End Region

#Region "Propriedades"

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Usuario_Id As String
        Get
            Return _Usuario_Id
        End Get
        Set(value As String)
            _Usuario_Id = value
        End Set
    End Property

    Public Property Empresa_Id As String
        Get
            Return _Empresa_Id
        End Get
        Set(value As String)
            _Empresa_Id = value
        End Set
    End Property

    Public Property EndEmpresa_Id As Integer
        Get
            Return _EndEmpresa_Id
        End Get
        Set(value As Integer)
            _EndEmpresa_Id = value
        End Set
    End Property

#End Region

#Region "Métodos"

    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Sqls.Clear()
        SalvarSql(Sqls)
        If Banco.GravaBanco(Sqls) Then
            Me.IUD = Nothing
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSQL As String

        Select Case Me.IUD

            Case "I"
                strSQL = "INSERT INTO UsuarioXEmpresasXDashboard (Usuario_Id, Empresa_Id, EndEmpresa_Id)" & vbCrLf &
                         "VALUES ('" & Me.Usuario_Id & "','" & Me.Empresa_Id & "'," & Me.EndEmpresa_Id & "); "
                Sqls.Add(strSQL)

            Case "D"
                strSQL = "  DELETE UsuarioXEmpresasXDashboard" & vbCrLf &
                         "  WHERE Usuario_Id        = '" & Me.Usuario_Id & "'" & vbCrLf &
                         "      AND Empresa_Id      = '" & Me.Empresa_Id & "'" & vbCrLf &
                         "      AND EndEmpresa_Id   = " & Me.EndEmpresa_Id & ";"
                Sqls.Add(strSQL)

        End Select
    End Sub

#End Region

End Class