Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListTipoDeAplicacao
    Inherits List(Of TipoDeAplicacao)

#Region "Construtor"
    Public Sub New()
        Dim sql As String = "Select Codigo_Id, Descricao from TiposDeAplicacoes"
        Dim banco As New AcessaBanco
        Dim ds As DataSet

        ds = banco.ConsultaDataSet(sql, "TD")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim TD As New TipoDeAplicacao
            TD.Codigo = row("Codigo_Id")
            TD.Descricao = row("Descricao")
            Me.Add(TD)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class TipoDeAplicacao
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim sql As String = "Select Codigo_Id, Descricao from TiposDeAplicacoes where Codigo_id = " & pCodigo
        Dim banco As New AcessaBanco
        Dim ds As DataSet

        ds = banco.ConsultaDataSet(sql, "TD")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        _Codigo = row("Codigo_id")
        _Descricao = row("Descricao")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer
    Private _Descricao As String
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
                sql = " Insert Into TiposDeAplicacoes(Codigo_Id, Descricao)" & vbCrLf & _
                      " Values (" & Me.Codigo & ",'" & Me.Descricao & "')"
                Sqls.Add(sql)
            Case "U"
                sql = " Update TiposDeAplicacoes set " & vbCrLf & _
                      "   Descricao ='" & Me.Descricao & "'" & vbCrLf & _
                      "	Where Codigo_Id =" & Me.Codigo
                Sqls.Add(sql)
            Case "D"
                sql = " Delete TiposDeAplicacoes " & vbCrLf & _
                      "  Where Codigo_Id =" & Me.Codigo
                Sqls.Add(sql)
        End Select

    End Sub
#End Region

End Class