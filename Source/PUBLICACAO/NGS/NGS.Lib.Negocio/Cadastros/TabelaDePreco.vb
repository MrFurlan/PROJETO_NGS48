Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis
Imports System.Web

<Serializable()>
Public Class ListTabelaDePreco
    Inherits List(Of TabelaDePreco)

#Region "Construtor"
    Public Sub New(Optional ByVal Codigo As Integer = 0)
        Dim sql As String = "Select Codigo_Id, Descricao, Ativo from TabelaDePrecos Where Ativo = 1"
        Dim banco As New AcessaBanco
        Dim ds As DataSet

        ds = banco.ConsultaDataSet(sql, "TP")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim tp As New TabelaDePreco
            tp.Codigo = row("Codigo_Id")
            tp.Descricao = row("Descricao")
            tp.Ativo = row("Ativo")
            Me.Add(tp)
        Next
    End Sub
#End Region

End Class

<Serializable()>
Public Class TabelaDePreco
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim sql As String = "Select Codigo_Id, Descricao, Ativo from TabelaDePrecos where Codigo_id = " & pCodigo
        Dim banco As New AcessaBanco
        Dim ds As DataSet

        ds = banco.ConsultaDataSet(sql, "TP")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        _Codigo = row("Codigo_id")
        _Descricao = row("Descricao")
        _Ativo = row("Ativo")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer
    Private _Descricao As String
    Private _Ativo As Boolean
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As DateTime

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

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property Ativo() As Boolean
        Get
            Return _Ativo
        End Get
        Set(ByVal value As Boolean)
            _Ativo = value
        End Set
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property UsuarioInclusaoData() As DateTime
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioInclusaoData = value
        End Set
    End Property

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property UsuarioAlteracaoData() As DateTime
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As DateTime)
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
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into TabelaDePrecos(Descricao, Ativo, UsuarioInclusao, " & " UsuarioInclusaoData)" & vbCrLf &
                      " Values ('" & Me.Descricao & "','" & IIf(Me.Ativo, 1, 0) & "','" & HttpContext.Current.Session("ssNomeUsuario") & "', " & "'" & Now().ToString("yyyy-MM-dd") & "')"
                Sqls.Add(sql)
            Case "U"
                sql = " Update TabelaDePrecos " & vbCrLf &
                      "   Set Descricao = '" & Me.Descricao & "', UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf &
                      "       UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      "	Where Codigo_Id = " & Me.Codigo
                Sqls.Add(sql)
            Case "D"
                sql = " Update TabelaDePrecos " & vbCrLf &
                      "   Set Ativo = 0" & vbCrLf & "', UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf &
                      "	Where Codigo_Id = " & Me.Codigo
                Sqls.Add(sql)
        End Select

    End Sub
#End Region

End Class