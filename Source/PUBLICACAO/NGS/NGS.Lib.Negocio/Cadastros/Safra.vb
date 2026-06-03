Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListSafra
    Inherits List(Of Safra)

    Public Sub New(Optional ByVal Carregar As Boolean = False)
        If Not Carregar Then Exit Sub
        Dim objBanco As New AcessaBanco()

        Dim strSQL As String = "SELECT Safra_Id, Produto, isnull(InicioDeSafra,Vencimento) as InicioDeSafra, Vencimento, isnull(Taxa,0) AS Taxa, isnull(Observacao,'') AS Observacao " & _
                               "  FROM Safras " & _
                               " ORDER BY Safra_Id"

        Dim dsSafras As DataSet = objBanco.ConsultaDataSet(strSQL, "Safras")

        For Each drSafra As DataRow In dsSafras.Tables(0).Rows
            Dim objSafra As New Safra

            objSafra.Codigo = drSafra("Safra_Id")
            objSafra.CodigoProduto = drSafra("Produto")
            objSafra.InicioDeSafra = drSafra("InicioDeSafra")
            objSafra.Vencimento = drSafra("Vencimento")
            objSafra.Taxa = drSafra("Taxa")
            objSafra.Observacao = drSafra("Observacao")

            Me.Add(objSafra)
        Next
    End Sub

End Class

<Serializable()> _
Public Class Safra
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As String)
        Dim objBanco As New AcessaBanco()

        Dim strSQL As String = " SELECT Safra_Id, Produto,isnull(InicioDeSafra,Vencimento) as InicioDeSafra, Vencimento, Taxa, Observacao " & _
                               "   FROM Safras " & _
                               "  WHERE Safra_Id = '" & Codigo.ToString() & "'"

        Dim dsSafras As DataSet = objBanco.ConsultaDataSet(strSQL, "Safras")

        For Each drSafra As DataRow In dsSafras.Tables(0).Rows
            _Codigo = drSafra("Safra_Id")
            _CodigoProduto = drSafra("Produto")
            _InicioDeSafra = drSafra("InicioDeSafra")
            _Vencimento = drSafra("Vencimento")
            _Taxa = IIf(IsDBNull(drSafra("Taxa")), 0, drSafra("Taxa"))
            _Observacao = IIf(IsDBNull(drSafra("Observacao")), String.Empty, drSafra("Observacao"))
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Codigo As String
    Private _CodigoProduto As String
    Private _InicioDeSafra As DateTime
    Private _Vencimento As DateTime
    Private _Taxa As Double
    Private _Observacao As String
#End Region

#Region "Property"
    Public Property Codigo() As String
        Get
            Return _Codigo
        End Get
        Set(ByVal value As String)
            _Codigo = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property InicioDeSafra() As DateTime
        Get
            Return _InicioDeSafra
        End Get
        Set(ByVal value As DateTime)
            _InicioDeSafra = value
        End Set
    End Property

    Public Property Vencimento() As DateTime
        Get
            Return _Vencimento
        End Get
        Set(ByVal value As DateTime)
            _Vencimento = value
        End Set
    End Property

    Public Property Taxa() As Double
        Get
            Return _Taxa
        End Get
        Set(ByVal value As Double)
            _Taxa = value
        End Set
    End Property

    Public Property Observacao() As String
        Get
            Return _Observacao
        End Get
        Set(ByVal value As String)
            _Observacao = value
        End Set
    End Property
#End Region

End Class