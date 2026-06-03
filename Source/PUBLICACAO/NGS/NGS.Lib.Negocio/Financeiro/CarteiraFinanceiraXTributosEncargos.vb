Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCarteiraFinanceiraXTributoEncargo
    Inherits List(Of CarteiraFinanceiraXTributoEncargo)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal CodigoCarteira As String)
        Dim sql As String
        sql = "SELECT Carteira_Id, Tributo_ID" & vbCrLf & _
              "  FROM CarteirasXTributos" & vbCrLf & _
              " Where Carteira_Id = '" & CodigoCarteira & "'"

        Dim ds As DataSet
        Dim Banco As New AcessaBanco

        ds = Banco.ConsultaDataSet(sql, "CxT")

        For Each Row As DataRow In ds.Tables(0).Rows
            Dim CxT As New CarteiraFinanceiraXTributoEncargo
            CxT.CodigoCarteira = Row("Carteira_Id")
            CxT.CodigoTributo = Row("Tributo_ID")
            Me.Add(CxT)
        Next

    End Sub
#End Region

End Class

<Serializable()> _
Public Class CarteiraFinanceiraXTributoEncargo
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoCarteira As String
    Private _Carteira As CarteiraFinanceira
    Private _CodigoTributo As String
    Private _Tributo As Encargo
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

    Public Property CodigoCarteira() As String
        Get
            Return _CodigoCarteira
        End Get
        Set(ByVal value As String)
            _CodigoCarteira = value
            _Carteira = Nothing
        End Set
    End Property

    Public ReadOnly Property Carteira() As CarteiraFinanceira
        Get
            If _Carteira Is Nothing And _CodigoCarteira.Length > 0 Then _Carteira = New CarteiraFinanceira(_CodigoCarteira)
            Return _Carteira
        End Get
    End Property

    Public Property CodigoTributo() As String
        Get
            Return _CodigoTributo
        End Get
        Set(ByVal value As String)
            _CodigoTributo = value
            _Tributo = Nothing
        End Set
    End Property

    Public ReadOnly Property Tributo() As Encargo
        Get
            If _Tributo Is Nothing And _CodigoTributo.Length > 0 Then _Tributo = New Encargo(_CodigoTributo)
            Return _Tributo
        End Get
    End Property
#End Region

End Class