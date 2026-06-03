Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListSituacaoTributaria
    Inherits List(Of SituacaoTributaria)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean)
        If CarregarDados Then
            Dim sql As String = "Select SituacaoTributaria_Id, Descricao from SituacaoTributaria Order By SituacaoTributaria_Id "

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "SituacaoTributaria")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim ST As New SituacaoTributaria
                ST.Codigo = row("SituacaoTributaria_Id")
                ST.Descricao = row("Descricao")
                Me.Add(ST)
            Next
        End If
    End Sub
End Class

<Serializable()> _
Public Class SituacaoTributaria
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select SituacaoTributaria_Id, Descricao from SituacaoTributaria where SituacaoTributaria_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "SituacaoTributaria")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("SituacaoTributaria_Id")
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

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = " INSERT INTO SituacaoTributaria(SituacaoTributaria_Id, Descricao) " & vbCrLf & _
                      " VALUES (" & _Codigo & ",'" & _Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE SituacaoTributaria SET" & vbCrLf & _
                      "    Descricao     = '" & _Descricao & "'" & vbCrLf & _
                      "  WHERE SituacaoTributaria_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE SituacaoTributaria" & vbCrLf & _
                      "  WHERE SituacaoTributaria_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class