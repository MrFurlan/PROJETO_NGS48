Imports System.Collections.Generic


<Serializable()>
Public Class ListUsuarioXGruposDeEstoqueXDashBoard
    Inherits List(Of UsuarioXGruposDeEstoqueXDashBoard)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pUsuario_Id As String)

        Dim Sql As String
        Sql = " SELECT UXGXD.Usuario_Id, UXGXD.Grupo_Id, GRU.Descricao " & vbCrLf &
              "   FROM UsuarioXGruposDeEstoqueXDashBoard As UXGXD " & vbCrLf &
              "   INNER JOIN GruposDeEstoques As GRU " & vbCrLf &
              "     ON UXGXD.Grupo_Id   = GRU.Grupo_Id " & vbCrLf &
              "  WHERE UXGXD.Usuario_Id ='" & pUsuario_Id & "'; "

        Dim Banco As New AcessaBanco
        Dim DsFaturas As DataSet = Banco.ConsultaDataSet(Sql, "UXGXD")

        For Each dr As DataRow In DsFaturas.Tables(0).Rows
            Dim usuario = New UsuarioXGruposDeEstoqueXDashBoard
            usuario.Usuario_Id = dr("Usuario_Id")
            usuario.Grupo_Id = dr("Grupo_Id")
            usuario.Descricao = dr("Descricao")

            Me.Add(usuario)

        Next

    End Sub

End Class

<Serializable()>
Public Class UsuarioXGruposDeEstoqueXDashBoard
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New()

    End Sub

#End Region

#Region "Fields"

    Private _IUD As String
    Private _Usuario_Id As String
    Private _Grupo_Id As String
    Private _Descricao As String

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

    Public Property Grupo_Id As String
        Get
            Return _Grupo_Id
        End Get
        Set(value As String)
            _Grupo_Id = value
        End Set
    End Property

    Public Property Descricao As String
        Get
            Return _Descricao
        End Get
        Set(value As String)
            _Descricao = value
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
                strSQL = "INSERT INTO UsuarioXGruposDeEstoqueXDashBoard (Usuario_Id, Grupo_Id, Descricao)" & vbCrLf &
                         "VALUES ('" & Me.Usuario_Id & "','" & Me.Grupo_Id & "','" & Me.Descricao & "'); "
                Sqls.Add(strSQL)

            Case "D"
                strSQL = "  DELETE UsuarioXGruposDeEstoqueXDashBoard" & vbCrLf &
                         "  WHERE Usuario_Id        = '" & Me.Usuario_Id & "'" & vbCrLf &
                         "      AND Grupo_Id        = '" & Me.Grupo_Id & "';"

                Sqls.Add(strSQL)

        End Select
    End Sub

#End Region

End Class