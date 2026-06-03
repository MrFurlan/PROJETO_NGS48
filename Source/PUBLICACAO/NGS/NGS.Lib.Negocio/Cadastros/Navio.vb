Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Web
Imports NGS.Lib.Uteis

<Serializable>
Public Class ListNavios
    Inherits List(Of Navio)
    Implements IBaseEntity

#Region "Constructor"
    Public Sub New()

    End Sub

    Public Sub New(Optional ByVal Carregar As Boolean = False, Optional ByVal pWhere As String = "")
        If Carregar Then
            Dim objBanco As New AcessaBanco
            Dim sql As String = " SELECT Codigo_Id, Descricao, Ativo, UsuarioInclusao, UsuarioInclusaoData, " & vbCrLf &
                                " ISNULL(UsuarioAlteracao, 0) as UsuarioAlteracao, ISNULL(UsuarioAlteracaoData, 0) as UsuarioAlteracaoData " & vbCrLf &
                                " FROM Navios"

            If pWhere.Trim.Length = 0 Then
                sql &= pWhere
            End If
            sql &= "  ORDER BY Codigo_Id "

            Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "Navios")

            If (ds IsNot Nothing AndAlso ds.Tables.Count > 0) Then
                For Each Row As DataRow In ds.Tables(0).Rows
                    Dim objNavio As New Navio
                    objNavio.Codigo = Row("Codigo_Id")
                    objNavio.Descricao = Row("Descricao")
                    objNavio.Ativo = Row("Ativo")
                    objNavio.UsuarioInclusao = Row("UsuarioInclusao")
                    objNavio.UsuarioInclusaoData = Row("UsuarioInclusaoData")
                    objNavio.UsuarioAlteracao = Row("UsuarioAlteracao")
                    objNavio.UsuarioAlteracaoData = Row("UsuarioAlteracaoData")
                    Me.Add(objNavio)
                Next
            End If
        End If
    End Sub
#End Region
End Class

<Serializable>
Public Class Navio
    Implements IBaseEntity

#Region "Constructor"
    Public Sub New()
    End Sub

    Public Sub New(Codigo As String)
        Dim Banco As New AcessaBanco
        Dim sql As String = " SELECT Codigo_Id, Descricao, Ativo, " & vbCrLf &
                            "UsuarioInclusao, UsuarioInclusaoData," & vbCrLf &
                            "ISNULL(UsuarioAlteracao, '') UsuarioAlteracao, ISNULL(UsuarioAlteracaoData, '') UsuarioAlteracaoData" & vbCrLf &
                            "FROM Navios" & vbCrLf &
                            "WHERE Codigo_Id = '" & Codigo & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Navios")
        If ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then Exit Sub
        Dim Row As DataRow = ds.Tables(0).Rows(0)

        _Codigo = Row("Codigo_Id")
        _Descricao = Row("Descricao")
        _Ativo = Row("Ativo")
        _UsuarioInclusao = Row("UsuarioInclusao")
        _UsuarioInclusaoData = Row("UsuarioInclusaoData")
        _UsuarioAlteracao = Row("UsuarioAlteracao")
        _UsuarioAlteracaoData = Row("UsuarioAlteracaoData")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As String
    Private _Descricao As String
    Private _Ativo As Boolean
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As String
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As String
#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property
    Public Property Codigo As String
        Get
            Return _Codigo
        End Get
        Set(value As String)
            _Codigo = value
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

    Public Property Ativo As Boolean
        Get
            Return _Ativo
        End Get
        Set(value As Boolean)
            _Ativo = value
        End Set
    End Property

    Public Property UsuarioInclusao As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(value As String)
            _UsuarioInclusao = value
        End Set
    End Property
    Public Property UsuarioInclusaoData As String
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(value As String)
            _UsuarioInclusaoData = value
        End Set
    End Property
    Public Property UsuarioAlteracao As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(value As String)
            _UsuarioAlteracao = value
        End Set
    End Property
    Public Property UsuarioAlteracaoData As String
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(value As String)
            _UsuarioAlteracaoData = value
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
        Dim sql As String = ""

        Select Case Me.IUD
            Case "I"
                sql = " INSERT INTO Navios (Descricao, " & vbCrLf &
                    " Ativo, " & vbCrLf &
                    " UsuarioInclusao, " & vbCrLf &
                    " UsuarioInclusaoData)" & vbCrLf &
                    " Values ('" & _Descricao & "'," & vbCrLf &
                    "'" & IIf(_Ativo, 1, 0) & "'," & vbCrLf &
                    "'" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf &
                    "'" & Now().ToString("yyyy-MM-dd") & "')"
                Sqls.Add(sql)
            Case "U"
                sql = " UPDATE Navios " & vbCrLf &
                      "     SET Descricao = '" & _Descricao & "', Ativo = " & IIf(_Ativo, 1, 0) & ", UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf &
                      "         UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      " WHERE Codigo_Id = " & _Codigo
                Sqls.Add(sql)
            Case "D"
                sql = " UPDATE Navios SET" & vbCrLf &
                sql = "    SET Ativo = " & IIf(_Ativo, 1, 0) & ", UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf &
                      "        UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      "  WHERE Codigo_Id = " & _Codigo
                Sqls.Add(sql)
        End Select
    End Sub
#End Region
End Class