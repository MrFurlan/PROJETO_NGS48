Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.VisualBasic.Financial
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ResumoFinanceiro
#Region "Construtor"
    Public Sub New(ByVal pCodigoEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pCodigoPedido As Integer)
        _CodigoEmpresa = pCodigoEmpresa
        _CodigoEndEmpresa = pEndEmpresa
        _CodigoPedido = pCodigoPedido
        Carga()
    End Sub

    Public Sub New(ByVal pPedido As Pedido)
        _CodigoEmpresa = pPedido.CodigoEmpresa
        _CodigoEndEmpresa = pPedido.EnderecoEmpresa
        _CodigoPedido = pPedido.Codigo
        Carga()
    End Sub

    Private Sub Carga()
        Dim sql As String
        sql = "SELECT UnidadeDeNegocio," & vbCrLf & _
              "       Empresa," & vbCrLf & _
              "       EndEmpresa," & vbCrLf & _
              "       Pedido," & vbCrLf & _
              "       Cifrao," & vbCrLf & _
              "       Troca," & vbCrLf & _
              "       ContaContabilProduto," & vbCrLf & _
              "       isnull(ContaContabilAdiantamento,'') as ContaContabilAdiantamento," & vbCrLf & _
              "       ValorPedido," & vbCrLf & _
              "       ValorTitulosEmPrevisao," & vbCrLf & _
              "       ValorTitulosEmProvisao," & vbCrLf & _
              "       ValorTitulosBaixado," & vbCrLf & _
              "       ValorTitulosCompensado," & vbCrLf & _
              "       ValorAdiantamentoOriginal," & vbCrLf & _
              "       ValorAdiantamento," & vbCrLf & _
              "       ValorAdiantamentoCompensado," & vbCrLf & _
              "       ValorAdiantamentoPagoDireto," & vbCrLf & _
              "       SaldoAdiantamento" & vbCrLf & _
              "  FROM ResumoFinanceiro" & vbCrLf & _
              " Where Empresa    ='" & _CodigoEmpresa & "'" & vbCrLf & _
              "   And EndEmpresa = " & _CodigoEndEmpresa & vbCrLf & _
              "   And Pedido     = " & _CodigoPedido & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "ResumoFinanceiro")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow
        row = ds.Tables(0).Rows(0)

        Me.CodigoUnidadeDeNegocio = row("UnidadeDeNegocio")

        Me.Troca = row("Troca")
        Me.ContaContabilProduto = row("ContaContabilProduto")
        Me.ContaContabilAdiantamento = row("ContaContabilAdiantamento")

        Me.ValorPedido = row("ValorPedido")
        Me.ValorTitulosEmPrevisao = row("ValorTitulosEmPrevisao")
        Me.ValorTitulosEmProvisao = row("ValorTitulosEmProvisao")
        Me.ValorTitulosBaixado = row("ValorTitulosBaixado")
        Me.ValorTitulosCompensado = row("ValorTitulosCompensado")

        Me.ValorAdiantamentoOriginal = row("ValorAdiantamentoOriginal")
        Me.ValorAdiantamento = row("ValorAdiantamento")
        Me.ValorAdiantamentoCompensado = row("ValorAdiantamentoCompensado")
        Me.ValorAdiantamentoPagoDireto = row("ValorAdiantamentoPagoDireto")
    End Sub
#End Region

#Region "fields"
    Private _CodigoUnidadeDeNegocio As String
    Private _CodigoEmpresa As String
    Private _CodigoEndEmpresa As Integer
    Private _CodigoPedido As Integer
    Private _Pedido As Pedido
    Private _Troca As Boolean = False

    Private _ContaContabilProduto As String
    Private _ContaContabilAdiantamento As String

    Private _ValorPedido As Decimal

    Private _ValorTitulosEmPrevisao As Decimal
    Private _ValorTitulosEmProvisao As Decimal
    Private _ValorTitulosBaixado As Decimal
    Private _ValorTitulosCompensado As Decimal

    Private _ValorAdiantamentoOriginal As Decimal
    Private _ValorAdiantamento As Decimal
    Private _ValorAdiantamentoCompensado As Decimal
    Private _ValorAdiantamentoPagoDireto As Decimal
    'Private _SaldoAdiantamento As Decimal

    Private _ResumoTroca As ResumoFinanceiro
#End Region

#Region "Property"
    Public Property CodigoUnidadeDeNegocio As String
        Get
            Return _CodigoUnidadeDeNegocio
        End Get
        Set(value As String)
            _CodigoUnidadeDeNegocio = value
        End Set
    End Property
    Public ReadOnly Property CodigoEmpresa As String
        Get
            Return _CodigoEmpresa
        End Get
    End Property
    Public ReadOnly Property CodigoEndEmpresa As Integer
        Get
            Return _CodigoEndEmpresa
        End Get
    End Property
    Public ReadOnly Property CodigoPedido As Integer
        Get
            Return _CodigoPedido
        End Get
    End Property
    Public ReadOnly Property Pedido As Pedido
        Get
            If _Pedido Is Nothing And _CodigoPedido > 0 Then _Pedido = New Pedido(Me.CodigoEmpresa, Me.CodigoEndEmpresa, Me.CodigoPedido)
            Return _Pedido
        End Get
    End Property
    Public Property Troca As Boolean
        Get
            Return _Troca
        End Get
        Set(value As Boolean)
            _Troca = value
        End Set
    End Property

    Public Property ContaContabilProduto As String
        Get
            Return _ContaContabilProduto
        End Get
        Set(value As String)
            _ContaContabilProduto = value
        End Set
    End Property
    Public Property ContaContabilAdiantamento As String
        Get
            Return _ContaContabilAdiantamento
        End Get
        Set(value As String)
            _ContaContabilAdiantamento = value
        End Set
    End Property

    Public Property ValorPedido As Decimal
        Get
            Return _ValorPedido
        End Get
        Set(ByVal value As Decimal)
            _ValorPedido = value
        End Set
    End Property

    Public Property ValorTitulosEmPrevisao As Decimal
        Get
            Return _ValorTitulosEmPrevisao
        End Get
        Set(ByVal value As Decimal)
            _ValorTitulosEmPrevisao = value
        End Set
    End Property
    Public Property ValorTitulosEmProvisao As Decimal
        Get
            Return _ValorTitulosEmProvisao
        End Get
        Set(ByVal value As Decimal)
            _ValorTitulosEmProvisao = value
        End Set
    End Property
    Public Property ValorTitulosBaixado As Decimal
        Get
            Return _ValorTitulosBaixado
        End Get
        Set(ByVal value As Decimal)
            _ValorTitulosBaixado = value
        End Set
    End Property
    Public Property ValorTitulosCompensado As Decimal
        Get
            Return _ValorTitulosCompensado
        End Get
        Set(ByVal value As Decimal)
            _ValorTitulosCompensado = value
        End Set
    End Property

    Public Property ValorAdiantamentoOriginal As Decimal
        Get
            Return _ValorAdiantamentoOriginal
        End Get
        Set(ByVal value As Decimal)
            _ValorAdiantamentoOriginal = value
        End Set
    End Property
    Public Property ValorAdiantamento As Decimal
        Get
            Return _ValorAdiantamento
        End Get
        Set(ByVal value As Decimal)
            _ValorAdiantamento = value
        End Set
    End Property
    Public Property ValorAdiantamentoCompensado As Decimal
        Get
            Return _ValorAdiantamentoCompensado
        End Get
        Set(ByVal value As Decimal)
            _ValorAdiantamentoCompensado = value
        End Set
    End Property
    Public Property ValorAdiantamentoPagoDireto As Decimal
        Get
            Return _ValorAdiantamentoPagoDireto
        End Get
        Set(value As Decimal)
            _ValorAdiantamentoPagoDireto = value
        End Set
    End Property

    Public ReadOnly Property ValorAdiantamentoAmortizado As Decimal
        Get
            If Troca Then
                Return _ValorAdiantamentoCompensado
            Else
                If Me.ValorAdiantamento > 0 Then
                    Dim vlr As Decimal = Me.ValorPedido - Me.ValorTitulosEmPrevisao - Me.ValorTitulosEmProvisao - Me.ValorTitulosBaixado - (Me.ValorAdiantamento - Me.ValorAdiantamentoCompensado)
                    If vlr < 0 Then Return 0
                    If vlr > Me.ValorAdiantamentoCompensado Then Return Me.ValorAdiantamentoCompensado
                    Return vlr
                Else
                    Return 0
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property ValorPago As Decimal
        Get
            Return Me.ValorTitulosBaixado + SaldoAdiantamento + (ValorAdiantamentoOriginal - SaldoAdiantamento)
        End Get
    End Property

    Public ReadOnly Property SaldoAdiantamento As Decimal
        Get
            Return ValorAdiantamento - ValorAdiantamentoCompensado
        End Get
    End Property

    Public ReadOnly Property SaldoBaixa As Decimal
        Get
            Return _ValorTitulosEmPrevisao + _ValorTitulosEmProvisao + SaldoAdiantamento
        End Get
    End Property

    Public ReadOnly Property ResumoTroca As ResumoFinanceiro
        Get
            If _ResumoTroca Is Nothing AndAlso Not Me.Pedido Is Nothing AndAlso Me.Pedido.CodigoPedidoTroca > 0 Then
                _ResumoTroca = New ResumoFinanceiro(Me.Pedido.CodigoEmpresaTroca, Me.Pedido.EnderecoEmpresaTroca, Me.Pedido.CodigoPedidoTroca)
                _ResumoTroca.Carga()
            End If
            Return _ResumoTroca
        End Get
    End Property
#End Region

End Class
