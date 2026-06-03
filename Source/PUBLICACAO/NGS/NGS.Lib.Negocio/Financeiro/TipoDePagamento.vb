Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListTipoDePagamento
    Inherits List(Of TipoDePagamento)

    Public Sub New(Optional ByVal carregar As Boolean = False)
        If Not carregar Then Exit Sub

        Dim Banco As New AcessaBanco
        Dim Sql As String
        Sql = "SELECT TipoDePagamento_Id, Descricao, isnull(EnviaAoBanco,'N') as EnviaAoBanco" & vbCrLf & _
              "  FROM TiposDePagamentos" & vbCrLf
        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "TipoPgto")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim TP As New TipoDePagamento
            TP.TipoDePagamento = row("TipoDePagamento_Id")
            TP.Descicao = row("Descricao")
            TP.EnviaAoBanco = row("EnviaAoBanco") = "S"
            Me.Add(TP)
        Next
    End Sub

End Class

<Serializable()> _
Public Class TipoDePagamento
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pTipoPagto As Integer)
        Dim Banco As New AcessaBanco
        Dim Sql As String
        Sql = "SELECT TipoDePagamento_Id, Descricao, isnull(EnviaAoBanco,'N') as EnviaAoBanco" & vbCrLf & _
              "  FROM TiposDePagamentos" & vbCrLf & _
              " Where TipoDePagamento_Id =" & pTipoPagto
        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "TipoPgto")

        For Each row As DataRow In ds.Tables(0).Rows
            _TipoDePagamento = row("TipoDePagamento_Id")
            _Descricao = row("Descricao")
            _EnviaAoBanco = row("EnviaAoBanco") = "S"
        Next
    End Sub
#End Region

#Region "Fields"
    Private _TipoDePagamento As Integer
    Private _Descricao As String
    Private _EnviaAoBanco As Boolean
#End Region

#Region "Property"
    Public Property TipoDePagamento() As Integer
        Get
            Return _TipoDePagamento
        End Get
        Set(ByVal value As Integer)
            _TipoDePagamento = value
        End Set
    End Property

    Public Property Descicao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property EnviaAoBanco() As Boolean
        Get
            Return _EnviaAoBanco
        End Get
        Set(ByVal value As Boolean)
            _EnviaAoBanco = value
        End Set
    End Property
#End Region

End Class

