
'******************************************************************************************************************************************************
'**************************************************    LISTA DO FINANCEIRO DO PEDIDO   ****************************************************************
'******************************************************************************************************************************************************
<Serializable()> _
Public Class ListPedidoxFinanceiro
    Inherits List(Of PedidoxFinanceiro)

#Region "Construtor"
    Public Sub New(pPedido As Pedido)
        _Pedido = pPedido

        Dim sql As String
        sql = "SELECT Origem, Tipo, Qualificador, Situacao, Registro, Vencimento, ParcelaOficial, ParcelaMoeda" & vbCrLf & _
              "  FROM VW_PedidoxFinanceiro" & vbCrLf & _
              " WHERE Empresa    ='" & pPedido.CodigoEmpresa & "'" & vbCrLf & _
              "   AND EndEmpresa = " & pPedido.EnderecoEmpresa & vbCrLf & _
              "   AND Pedido     = " & pPedido.Codigo

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "PF")

        For Each row In ds.Tables(0).Rows
            Dim PF As New PedidoxFinanceiro(pPedido)
            PF.Origem = row("Origem")
            PF.Tipo = row("Tipo")
            PF.Qualificador = row("Qualificador")
            PF.Situacao = row("Situacao")
            PF.Registro = row("Registro")
            PF.Vencimento = row("Vencimento")
            PF.ValorOficial = row("ParcelaOficial")
            PF.ValorMoeda = row("ParcelaMoeda")
            Me.Add(PF)
        Next

    End Sub
#End Region

#Region "Fields"
    Private _Pedido As Pedido
#End Region

#Region "Property"
    Public ReadOnly Property Pedido As Pedido
        Get
            Return _Pedido
        End Get
    End Property
#End Region
End Class


'******************************************************************************************************************************************************
'*************************************************    CLASSE BASE FINANCEIRO DO PEDIDO   **************************************************************
'******************************************************************************************************************************************************
<Serializable()> _
Public Class PedidoxFinanceiro

#Region "Construtor"
    Public Sub New(pPedido As Pedido)
        _Pedido = pPedido
    End Sub
#End Region

#Region "Fields"
    Private _Pedido As Pedido
    Private _Origem As String
    Private _Tipo As String
    Private _Qualificador As String
    Private _Situacao As String
    Private _Registro As String
    Private _Vencimento As Date
    Private _ValorOficial As Decimal
    Private _ValorMoeda As Decimal
#End Region

#Region "Property"
    Public ReadOnly Property Pedido As Pedido
        Get
            Return _Pedido
        End Get
    End Property

    Public Property Origem As String
        Get
            Return _Origem
        End Get
        Set(value As String)
            _Origem = value
        End Set
    End Property

    Public Property Tipo As String
        Get
            Return _Tipo
        End Get
        Set(value As String)
            _Tipo = value
        End Set
    End Property

    Public Property Qualificador As String
        Get
            Return _Qualificador
        End Get
        Set(value As String)
            _Qualificador = value
        End Set
    End Property

    Public Property Situacao As String
        Get
            Return _Situacao
        End Get
        Set(value As String)
            _Situacao = value
        End Set
    End Property

    Public Property Registro As String
        Get
            Return _Registro
        End Get
        Set(value As String)
            _Registro = value
        End Set
    End Property

    Public Property Vencimento As Date
        Get
            Return _Vencimento
        End Get
        Set(value As Date)
            _Vencimento = value
        End Set
    End Property

    Public Property ValorOficial As Decimal
        Get
            Return _ValorOficial
        End Get
        Set(value As Decimal)
            _ValorOficial = value
        End Set
    End Property

    Public Property ValorMoeda As Decimal
        Get
            Return _ValorMoeda
        End Get
        Set(value As Decimal)
            _ValorMoeda = value
        End Set
    End Property
#End Region

End Class
