Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListSituacaoTributariaIPI
    Inherits List(Of SituacaoTributariaIPI)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean)
        If CarregarDados Then
            Dim sql As String = "Select SituacaoTributariaIPI_Id, Descricao from SituacaoTributariaIPI Order By SituacaoTributariaIPI_Id "

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "SituacaoTributariaIPI")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim ST As New SituacaoTributariaIPI
                ST.Codigo = row("SituacaoTributariaIPI_Id")
                ST.Descricao = row("Descricao")
                Me.Add(ST)
            Next
        End If
    End Sub
End Class

<Serializable()> _
Public Class SituacaoTributariaIPI
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select SituacaoTributariaIPI_Id, Descricao from SituacaoTributariaIPI where SituacaoTributariaIPI_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "SituacaoTributariaIPI")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("SituacaoTributariaIPI_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _Descricao As String = ""
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

    Public ReadOnly Property DescricaoObj() As String
        Get
            Return Format(_Codigo, "000") & " - " & _Descricao
        End Get
    End Property

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Dim retorno As Boolean

        Sqls.Clear()

        SalvarSql(Sqls)

        If IUD = "D" Then
            If Not JaUsada() Then
                retorno = Banco.GravaBanco(Sqls)
            Else
                retorno = False
            End If
        Else
            retorno = Banco.GravaBanco(Sqls)
        End If

        Return retorno
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = " INSERT INTO SituacaoTributariaIPI(SituacaoTributariaIPI_Id, Descricao) " & vbCrLf & _
                      " VALUES (" & _Codigo & ",'" & _Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE SituacaoTributariaIPI SET" & vbCrLf & _
                      "    Descricao     = '" & _Descricao & "'" & vbCrLf & _
                      "  WHERE SituacaoTributariaIPI_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE SituacaoTributariaIPI" & vbCrLf & _
                      "  WHERE SituacaoTributariaIPI_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function JaUsada() As Boolean
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim retorno As Boolean = False
        Dim sql As String = ""

        sql = "SELECT Top 1 STIPI from OperacaoXEstado WHERE STIPI = " & _Codigo
        ds = Banco.ConsultaDataSet(sql, "SituacaoTributariaIPI")
        retorno = ds.Tables(0).Rows.Count > 0

        Return retorno
    End Function

#End Region

End Class