Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis
Imports System.Web

<Serializable()> _
Public Class ListEPI
    Inherits List(Of EPI)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean)
        If CarregarDados Then
            Dim sql As String = "Select Codigo_Id, Descricao, isnull(Ativo,0) as Ativo From EPI"
            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "EPI")
            For Each row As DataRow In ds.Tables(0).Rows
                Dim Cr As New EPI
                Cr.Codigo = row("Codigo_Id")
                Cr.Descricao = row("Descricao")
                Cr.Ativo = row("Ativo")

                Me.Add(Cr)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class EPI

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet

        ds = Banco.ConsultaDataSet("Select Codigo_Id, Descricao, isnull(Ativo,0) as Ativo From EPI Where Codigo_Id = " & pCodigo, "EPI")

        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Codigo_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
            _Ativo = ds.Tables(0).Rows(0)("Ativo")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _Descricao As String = ""
    Private _Ativo As Boolean = False

    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As DateTime
    Private _UsuarioExclusao As String = ""
    Private _DataExclusao As DateTime
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

    Public Property DataInclusao() As DateTime
        Get
            Return _DataInclusao
        End Get
        Set(ByVal value As DateTime)
            _DataInclusao = value
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

    Public Property DataAlteracao() As DateTime
        Get
            Return _DataAlteracao
        End Get
        Set(ByVal value As DateTime)
            _DataAlteracao = value
        End Set
    End Property

    Public Property UsuarioExclusao() As String
        Get
            Return _UsuarioExclusao
        End Get
        Set(ByVal value As String)
            _UsuarioExclusao = value
        End Set
    End Property

    Public Property DataExclusao() As DateTime
        Get
            Return _DataExclusao
        End Get
        Set(ByVal value As DateTime)
            _DataExclusao = value
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
                Sql = " INSERT INTO EPI (Descricao, Ativo, UsuarioInclusao, UsuarioInclusaoData) " & vbCrLf & _
                      " VALUES ('" & _Descricao & "','" & _Ativo.ToString() & "','" & HttpContext.Current.Session("ssNomeUsuario") & "',getdate())"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE EPI SET" & vbCrLf & _
                      "   Descricao            ='" & _Descricao & "'" & vbCrLf & _
                      "  ,Ativo                ='" & _Ativo.ToString() & "'" & vbCrLf & _
                      "	 ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                      "	 ,UsuarioAlteracaoData = getdate() " & vbCrLf & _
                      "  WHERE Codigo_Id =" & _Codigo
                Sqls.Add(Sql)
            Case "D"
                Sql = " UPDATE EPI SET" & vbCrLf & _
                      "   Ativo               = 0" & vbCrLf & _
                      "	 ,UsuarioExclusao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                      "	 ,UsuarioExclusaoData = getdate() " & vbCrLf & _
                      "  WHERE Codigo_Id =" & _Codigo
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class