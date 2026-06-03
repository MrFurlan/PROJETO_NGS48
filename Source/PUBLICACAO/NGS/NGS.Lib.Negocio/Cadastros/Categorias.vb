Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCategorias
    Inherits List(Of Categorias)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean)
        If CarregarDados Then
            Dim sql As String = "Select Categoria_Id, Descricao, Funrural, BasePisCofins from Categorias "

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "Categorias")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Ca As New Categorias
                Ca.Codigo = row("Categoria_Id")
                Ca.Descricao = row("Descricao")
                Ca.Funrural = row("Funrural")
                Ca.BasePisCofins = row("BasePisCofins")
                Me.Add(Ca)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class Categorias
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Categoria_Id, Descricao, Funrural, BasePisCofins from Categorias where Categoria_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "Categorias")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Categoria_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
            _Funrural = ds.Tables(0).Rows(0)("Funrural")
            _BasePisCofins = ds.Tables(0).Rows(0)("BasePisCofins")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _Descricao As String = ""
    Private _Funrural As Decimal = 0
    Private _BasePisCofins As Integer = 0
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

    Public Property Funrural() As Decimal
        Get
            Return _Funrural
        End Get
        Set(ByVal value As Decimal)
            _Funrural = value
        End Set
    End Property

    Public Property BasePisCofins() As Integer
        Get
            Return _BasePisCofins
        End Get
        Set(ByVal value As Integer)
            _BasePisCofins = value
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
                Sql = " INSERT INTO Categorias(Categoria_Id, Descricao, Funrural, BasePisCofins) " & vbCrLf & _
                      " VALUES (" & _Codigo & ",'" & _Descricao & "'," & Str(_Funrural) & "," & _BasePisCofins & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Categorias SET" & vbCrLf & _
                      "    Descricao        = '" & _Descricao & "'" & vbCrLf & _
                      "    Funrural         = " & Str(_Funrural) & vbCrLf & _
                      "    BasePisCofins    = " & _BasePisCofins & vbCrLf & _
                      "  WHERE Categoria_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Categorias" & vbCrLf & _
                      "  WHERE Categoria_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
