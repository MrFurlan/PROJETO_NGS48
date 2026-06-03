Imports System.Collections.Generic


<Serializable()>
Public Class ListUsuarioXProdutosXDashBoard
    Inherits List(Of UsuarioXProdutosXDashBoard)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pUsuario_Id As String)

        Dim Sql As String
        Sql = " SELECT UXPXD.Usuario_Id, UXPXD.Grupo_Id, GRU.Descricao, UXPXD.Produto_Id, PRO.Nome " & vbCrLf &
              "   FROM UsuarioXProdutosXDashBoard As UXPXD " & vbCrLf &
              "   INNER JOIN GruposDeEstoques As GRU " & vbCrLf &
              "     ON UXPXD.Grupo_Id       = GRU.Grupo_Id " & vbCrLf &
              "   INNER JOIN Produtos As PRO " & vbCrLf &
              "     ON UXPXD.Produto_Id     = PRO.Produto_Id " & vbCrLf &
              "  WHERE UXPXD.Usuario_Id     ='" & pUsuario_Id & "'; "

        Dim Banco As New AcessaBanco
        Dim DsFaturas As DataSet = Banco.ConsultaDataSet(Sql, "UXPXD")

        For Each dr As DataRow In DsFaturas.Tables(0).Rows
            Dim usuario = New UsuarioXProdutosXDashBoard
            usuario.Usuario_Id = dr("Usuario_Id")
            usuario.Grupo_Id = dr("Grupo_Id")
            usuario.Descricao = dr("Descricao")
            usuario.Produto_Id = dr("Produto_Id")
            usuario.Nome = dr("Nome")

            Me.Add(usuario)

        Next

    End Sub

End Class

<Serializable()>
Public Class UsuarioXProdutosXDashBoard
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
    Private _Produto_Id As String
    Private _Nome As String

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

    Public Property Produto_Id As String
        Get
            Return _Produto_Id
        End Get
        Set(value As String)
            _Produto_Id = value
        End Set
    End Property

    Public Property Nome As String
        Get
            Return _Nome
        End Get
        Set(value As String)
            _Nome = value
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
                strSQL = "INSERT INTO UsuarioXProdutosXDashBoard (Usuario_Id, Grupo_Id, Produto_Id)" & vbCrLf &
                         "VALUES ('" & Me.Usuario_Id & "','" & Me.Grupo_Id & "','" & Me.Produto_Id & "'); "
                Sqls.Add(strSQL)

            Case "D"
                strSQL = "  DELETE UsuarioXProdutosXDashBoard" & vbCrLf &
                         "  WHERE Usuario_Id        = '" & Me.Usuario_Id & "'" & vbCrLf &
                         "      AND Grupo_Id        = '" & Me.Grupo_Id & "'" & vbCrLf &
                         "      AND Produto_Id      = '" & Me.Produto_Id & ";"
                Sqls.Add(strSQL)

        End Select
    End Sub

#End Region

End Class