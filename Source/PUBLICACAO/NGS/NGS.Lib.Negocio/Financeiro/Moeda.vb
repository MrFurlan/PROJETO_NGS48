Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'*******************************************************************************************************
'***********************  List Classe Moeda ************************************************************
'*******************************************************************************************************
<Serializable()> _
Public Class Moedas
    Inherits List(Of Moeda)

    Public Sub New()
        Dim objBanco As New AcessaBanco()

        Try
            Dim sql As String
            sql = "SELECT Moeda_Id, Descricao, Classificacao, Cifrao as Simbolo " & vbCrLf & _
                  "  FROM Moedas"

            Dim dsMoedas As DataSet = objBanco.ConsultaDataSet(sql, "Moedas")

            For Each drMoeda As DataRow In dsMoedas.Tables(0).Rows
                Dim objMoeda As New Moeda

                objMoeda.Codigo = drMoeda("Moeda_Id")
                objMoeda.Descricao = drMoeda("Descricao")
                objMoeda.Classificacao = Conversoes.ConverterTipoMoeda(drMoeda("Classificacao").ToString())
                objMoeda.Simbolo = drMoeda("Simbolo")

                Me.Add(objMoeda)
            Next
        Catch ex As Exception

        Finally
            objBanco = Nothing
        End Try
    End Sub

End Class

'*******************************************************************************************************
'***********************  Classe Base Moeda ************************************************************
'*******************************************************************************************************
<Serializable()> _
Public Class Moeda
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Dim objBanco As New AcessaBanco()

        Try
            Dim sql As String
            sql = "SELECT Moeda_Id, Descricao, Classificacao, Cifrao as Simbolo " & vbCrLf & _
                  "  FROM Moedas " & vbCrLf & _
                  " WHERE Moeda_Id = " & Codigo.ToString()

            Dim dsMoedas As DataSet = objBanco.ConsultaDataSet(sql, "Moedas")

            If dsMoedas.Tables(0).Rows.Count > 0 Then
                Dim drMoeda As DataRow = dsMoedas.Tables(0).Rows(0)

                Me.Codigo = drMoeda("Moeda_Id")
                Me.Descricao = drMoeda("Descricao")
                Me.Classificacao = Conversoes.ConverterTipoMoeda(drMoeda("Classificacao").ToString())
                Me.Simbolo = drMoeda("Simbolo")
            End If

        Catch ex As Exception

        Finally
            objBanco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    Private _Codigo As Integer
    Private _Descricao As String
    Private _Classificacao As eTiposMoeda
    Private _Simbolo As String
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

    Public Property Classificacao() As eTiposMoeda
        Get
            Return _Classificacao
        End Get
        Set(ByVal value As eTiposMoeda)
            _Classificacao = value
        End Set
    End Property

    Public Property Simbolo() As String
        Get
            Return _Simbolo
        End Get
        Set(ByVal value As String)
            _Simbolo = value
        End Set
    End Property
#End Region

End Class

