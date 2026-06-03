Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListTipoDeDocumento
    Inherits List(Of TipoDeDocumento)

#Region "Construtor"
    Public Sub New()
        Dim sql As String = "Select Codigo_Id, Descricao, Historico from TipoDeDocumento"
        Dim banco As New AcessaBanco
        Dim ds As DataSet

        ds = banco.ConsultaDataSet(sql, "TD")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim TD As New TipoDeDocumento
            TD.Codigo = row("Codigo_Id")
            TD.Descricao = row("Descricao")
            TD.Historico = row("Historico")
            Me.Add(TD)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class TipoDeDocumento
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim sql As String = "Select Codigo_Id, Descricao, Historico from TipoDeDocumento where Codigo_id = " & pCodigo
        Dim db As New AcessaBanco
        Dim ds As DataSet

        ds = db.ConsultaDataSet(sql, "TipoDeDocumento")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            _Codigo = row("Codigo_Id")
            _Descricao = row("Descricao")
            _Historico = row("Historico")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer
    Private _Descricao As String
    Private _Historico As String
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

    Public Property Historico() As String
        Get
            Return _Historico
        End Get
        Set(ByVal value As String)
            _Historico = value
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
                sql = " Insert Into TipoDeDocumento(Codigo_Id, Descricao, Historico)" & vbCrLf & _
                      " Values (" & _Codigo & ",'" & _Descricao & "','" & _Historico & "')"
                Sqls.Add(sql)
            Case "U"
                sql = " Update TipoDeDocumento set " & vbCrLf & _
                      "   Descricao ='" & _Descricao & "'" & vbCrLf & _
                      "  ,Historico ='" & _Historico & "'" & vbCrLf & _
                      "	Where Codigo_Id =" & _Codigo
                Sqls.Add(sql)
            Case "D"
                sql = " Delete TipoDeDocumento" & vbCrLf & _
                      "  Where Codigo_Id =" & _Codigo
                Sqls.Add(sql)
        End Select

    End Sub
#End Region

End Class