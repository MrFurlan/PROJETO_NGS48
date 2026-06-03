Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListGeneroDoProdutoXSub
    Inherits List(Of GeneroDoProdutoXSub)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean)
        If CarregarDados Then
            Dim sql As String = "Select Codigo_Id, SubCodigo_Id, Descricao, CodigoPamcard from GeneroDoProdutoXSub Order By Codigo_Id, SubCodigo_Id "

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "GeneroDoProdutoXSub")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim ST As New GeneroDoProdutoXSub
                ST.Codigo = row("Codigo_Id")
                ST.SubCodigo = row("SubCodigo_Id")
                ST.Descricao = row("Descricao")
                ST.CodigoPamcard = row("CodigoPamcard")
                Me.Add(ST)
            Next
        End If
    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim sql As String = "Select Codigo_Id, SubCodigo_Id, Descricao, CodigoPamcard " & vbCrLf & _
                            "From GeneroDoProdutoXSub " & vbCrLf & _
                            "Where Codigo_id = " & pCodigo & vbCrLf & _
                            " Order SubCodigo_Id "

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "GeneroDoProdutoXSub")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ST As New GeneroDoProdutoXSub
            ST.Codigo = row("Codigo_Id")
            ST.SubCodigo = row("SubCodigo_Id")
            ST.Descricao = row("Descricao")
            ST.CodigoPamcard = row("CodigoPamcard")
            Me.Add(ST)
        Next
    End Sub

End Class

<Serializable()> _
Public Class GeneroDoProdutoXSub
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String, ByVal pSubCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Codigo_Id, SubCodigo_Id, Descricao, CodigoPamcard " & vbCrLf & _
                            "From GeneroDoProdutoXSub " & vbCrLf & _
                            "Where Codigo_Id = " & pCodigo & vbCrLf & _
                            "  and SubCodigo_Id = " & pSubCodigo & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "GeneroDoProdutoXSub")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Codigo_Id")
            _SubCodigo = ds.Tables(0).Rows(0)("SubCodigo_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
            _CodigoPamcard = ds.Tables(0).Rows(0)("CodigoPamcard")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _SubCodigo As Integer = 0
    Private _Descricao As String = ""
    Private _CodigoPamcard As String = ""
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

    Public Property SubCodigo() As Integer
        Get
            Return _SubCodigo
        End Get
        Set(ByVal value As Integer)
            _SubCodigo = value
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

    Public Property CodigoPamcard() As String
        Get
            Return _CodigoPamcard
        End Get
        Set(ByVal value As String)
            _CodigoPamcard = value
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
                Sql = " INSERT INTO GeneroDoProdutoXSub(Codigo_Id, SubCodigo_Id, Descricao, CodigoPamcard) " & vbCrLf & _
                      " VALUES (" & _Codigo & "," & _SubCodigo & ",'" & _Descricao & "','" & _CodigoPamcard & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE GeneroDoProdutoXSub SET" & vbCrLf & _
                      "    Descricao     = '" & _Descricao & "'," & vbCrLf & _
                      "    CodigoPamcard = '" & _Descricao & "'" & vbCrLf & _
                      "  WHERE Codigo_Id    = " & _Codigo & vbCrLf & _
                      "    AND SubCodigo_Id = " & _SubCodigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE GeneroDoProdutoXSub " & vbCrLf & _
                      "  WHERE Codigo_Id    = " & _Codigo & vbCrLf & _
                      "    AND SubCodigo_Id = " & _SubCodigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class