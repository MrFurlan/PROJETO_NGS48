Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListModeloCertidao
    Inherits List(Of ModeloCertidao)

#Region "Construtor"
    Public Sub New(Optional ByVal pCarregar As Boolean = False)
        If Not pCarregar Then Exit Sub

        Dim Banco As New AcessaBanco
        Dim sql As String
        sql = "SELECT ModeloCertidao_Id, Descricao" & vbCrLf & _
              "  FROM CertidaoNegativaModelo" & vbCrLf
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "CNModelo")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CN As New ModeloCertidao
            CN.Codigo = row("ModeloCertidao_Id")
            CN.Descricao = row("Descricao")
            Me.Add(CN)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ModeloCertidao
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo)
        Dim Banco As New AcessaBanco
        Dim sql As String
        sql = "SELECT ModeloCertidao_Id, Descricao" & vbCrLf & _
              "  FROM CertidaoNegativaModelo" & vbCrLf & _
              " Where ModeloCertidao_Id = " & pCodigo
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "CNModelo")

        For Each row As DataRow In ds.Tables(0).Rows
            _Codigo = row("ModeloCertidao_Id")
            _Descricao = row("Descricao")
        Next

    End Sub
#End Region

#Region "Fields"
    Private _Codigo As Integer
    Private _Descricao As String
#End Region

#Region "Property"
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

End Class

