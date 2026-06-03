Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

'******************************************************************************************************************************************************
'**************************************************    LISTA DE Titulos Virtuais Pedido   *************************************************************
'******************************************************************************************************************************************************
Public Class ListPedidoXTituloVirtual
    Inherits List(Of PedidoXTituloVirtual)

#Region "Construtor"
    Public Sub New(ByVal pPedido As Pedido)
        _Pedido = pPedido
        If Me.Pedido.Codigo = 0 Then Exit Sub

        Dim Banco As New AcessaBanco()
        Try
            Dim strSQL As String = ""
            strSQL = "SELECT cliente, endcliente, cr, moeda, classificacao, Datapedido, Parcela_id, Vencimento, Indice, ParcelaOficial, ParcelaMoeda" & vbCrLf & _
                     "  FROM VW_TituloVirtual" & vbCrLf & _
                     " WHERE Empresa_Id    ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                     "   AND EndEmpresa_Id = " & Me.Pedido.EnderecoEmpresa & vbCrLf & _
                     "   AND Pedido_Id     = " & Me.Pedido.Codigo & vbCrLf & _
                     " Order by Parcela_id" & vbCrLf

            Dim ds As DataSet
            ds = Banco.ConsultaDataSet(strSQL, "PXTV")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim objPxTV As New PedidoXTituloVirtual(Me.Pedido)
                objPxTV.CodigoCliente = row("cliente")
                objPxTV.EnderecoCliente = row("endcliente")
                objPxTV.TipoTitulo = row("cr")
                objPxTV.CodigoMoeda = row("moeda")
                objPxTV.Classificacao = row("classificacao")
                objPxTV.NrParcela = row("Parcela_id")
                objPxTV.Movimento = row("Datapedido")
                objPxTV.Vencimento = row("Vencimento")
                objPxTV.ValorOficial = row("ParcelaOficial")
                objPxTV.ValorMoeda = row("ParcelaMoeda")

                Me.Add(objPxTV)
            Next
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    Private _Pedido As Pedido
    Private _ApuracaoParcela As PedidoxParcelaApuracao
    Private _Msg As String = ""
#End Region

#Region "Property"
    Public ReadOnly Property Pedido() As Pedido
        Get
            Return _Pedido
        End Get
    End Property

    Public ReadOnly Property ApuracaoParcela As PedidoxParcelaApuracao
        Get
            If _ApuracaoParcela Is Nothing Then _ApuracaoParcela = New PedidoxParcelaApuracao(Me.Pedido)
            Return _ApuracaoParcela
        End Get
    End Property

    Public Property MSG() As String
        Get
            Return _Msg
        End Get
        Set(ByVal value As String)
            _Msg = value
        End Set
    End Property

    Public ReadOnly Property TotalOficial As Decimal
        Get
            Return (From x In Me Select x.ValorOficial).Sum()
        End Get
    End Property

    Public ReadOnly Property TotalMoeda As Decimal
        Get
            Return (From x In Me Select x.ValorMoeda).Sum()
        End Get
    End Property
#End Region

#Region "Methods"
    'Public Sub SalvarSql(ByVal Sqls As ArrayList)
    '    For Each Parc As PedidoXParcela In Me
    '        If Pedido.IUD = "I" Or Pedido.IUD = "D" Then Parc.IUD = Pedido.IUD
    '        Parc.SalvarSql(Sqls)
    '    Next
    'End Sub

    'Public Sub Parcelar()
    '    Dim ValorParcela As Decimal
    '    Dim VlrParcelaAjuste As Decimal

    '    If Me.Pedido.IUD = "I" Then
    '        Me.Clear()
    '        ValorParcela = Math.Round(Me.Pedido.Itens.Liquido / Me.Pedido.CondicaoPagamento.Parcelas, 2)
    '        VlrParcelaAjuste = Me.Pedido.Itens.Liquido - (ValorParcela * Me.Pedido.CondicaoPagamento.Parcelas)

    '        For i As Integer = 1 To Me.Pedido.CondicaoPagamento.Parcelas
    '            Dim Parc As New PedidoXParcela(Me.Pedido)
    '            Parc.CodigoParcela = i
    '            Parc.DataVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(Me.Pedido.CondicaoPagamento.Periodo(i - 1)))
    '            Parc.Valor = ValorParcela
    '            Me.Add(Parc)
    '        Next
    '        Me(Me.Count - 1).Valor += VlrParcelaAjuste
    '    Else
    '        ValorParcela = Me.Pedido.Itens.Liquido / Me.Pedido.CondicaoPagamento.Parcelas
    '        VlrParcelaAjuste = Me.Pedido.Itens.Liquido - (ValorParcela * Me.Pedido.CondicaoPagamento.Parcelas)

    '        Me.RemoveAll(Function(s) s.IUD = "I")
    '        Me.ForEach(Function(s)
    '                       If s.CodigoParcela > Me.Pedido.CondicaoPagamento.Parcelas Then
    '                           s.IUD = "D"
    '                       End If
    '                       Return True
    '                   End Function)

    '        Dim Parc As PedidoXParcela
    '        Dim t As Integer = 1
    '        For x As Integer = 1 To Me.Pedido.CondicaoPagamento.Parcelas
    '            Parc = Me.Where(Function(S) S.CodigoParcela = t).FirstOrDefault
    '            If Parc Is Nothing Then
    '                Parc = New PedidoXParcela(Me.Pedido)
    '                Parc.IUD = "I"
    '                Parc.CodigoParcela = x
    '            Else
    '                Parc.IUD = "U"
    '            End If

    '            Parc.DataVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(Me.Pedido.CondicaoPagamento.Periodo(x - 1)))
    '            Parc.Valor = ValorParcela
    '            If Parc.IUD = "I" Then Me.Add(Parc)
    '            t += 1
    '        Next
    '        Me(Me.Pedido.CondicaoPagamento.Parcelas - 1).Valor += VlrParcelaAjuste
    '    End If

    'End Sub

    'Public Sub ModificarParcela(ByVal pCodigoParcela As Integer, ByVal NovoVencimento As DateTime, ByVal NovoValor As Decimal)
    '    Dim objParcela As PedidoXParcela = Me.Where(Function(s) s.CodigoParcela = pCodigoParcela).First
    '    Dim Venc As Date = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, NovoVencimento)

    '    If Venc > objParcela.DataVencimento And objParcela.CodigoParcela < Me.Pedido.CondicaoPagamento.Parcelas Then
    '        Dim objParcelaProx As PedidoXParcela = Me.Where(Function(s) s.CodigoParcela = pCodigoParcela + 1).First

    '        If objParcelaProx IsNot Nothing Then
    '            If Venc > objParcelaProx.DataVencimento Then
    '                MSG = "Vencimento da Parcela não pode ser maior do que o vencimento da próxima parcela"
    '                Exit Sub
    '            End If
    '        End If
    '    End If


    '    Dim conf As Decimal = (From p In Me Where p.CodigoParcela < pCodigoParcela Group By p.Pedido.Codigo Into valorv = Sum(p.Valor) Select valorv).FirstOrDefault

    '    If conf + NovoValor > Me.Pedido.Itens.Liquido Then
    '        MSG = "Valor da Parcela somado as demais parcelas anteriores nao pode exceder o valor do pedido."
    '        Exit Sub
    '    End If

    '    objParcela.IUD = "U"
    '    objParcela.DataVencimento = Venc

    '    If objParcela.Valor <> NovoValor Then
    '        Dim VlrParcelaAjuste As Decimal
    '        Dim AjusteFinal As Decimal
    '        Dim NumPar As Integer = Me.Pedido.CondicaoPagamento.Parcelas - objParcela.CodigoParcela

    '        If NumPar = 0 Then
    '            MSG = "O Valor da ultima parcela nao pode ser alterado."
    '            Exit Sub
    '        End If

    '        If NovoValor > objParcela.Valor Then
    '            VlrParcelaAjuste = Math.Round((NovoValor - objParcela.Valor) / NumPar, 2)
    '            AjusteFinal = (NovoValor - objParcela.Valor) - (VlrParcelaAjuste * NumPar)

    '            For i As Integer = objParcela.CodigoParcela To Me.Pedido.CondicaoPagamento.Parcelas - 1
    '                Me(i).IUD = "U"
    '                Me(i).Valor -= VlrParcelaAjuste
    '            Next

    '            Me(Me.Pedido.CondicaoPagamento.Parcelas - 1).Valor += AjusteFinal
    '        Else
    '            VlrParcelaAjuste = Math.Round((objParcela.Valor - NovoValor) / NumPar, 2)
    '            AjusteFinal = (objParcela.Valor - NovoValor) - (VlrParcelaAjuste * NumPar)

    '            For i As Integer = objParcela.CodigoParcela To Me.Pedido.CondicaoPagamento.Parcelas - 1
    '                Me(i).IUD = "U"
    '                Me(i).Valor += VlrParcelaAjuste
    '            Next

    '            Me(Me.Pedido.CondicaoPagamento.Parcelas - 1).Valor += AjusteFinal
    '        End If
    '    End If
    '    objParcela.Valor = NovoValor

    'End Sub

#End Region

End Class

'******************************************************************************************************************************************************
'****************************************************    CLASSE BASE TituloVirtual   ******************************************************************
'******************************************************************************************************************************************************
Public Class PedidoXTituloVirtual

#Region "Construtor"
    Public Sub New(Optional pPedido As Pedido = Nothing)
        If pPedido IsNot Nothing Then
            Me.CodigoEmpresa = pPedido.CodigoEmpresa
            Me.EnderecoEmpresa = pPedido.EnderecoEmpresa
            Me.CodigoPedido = pPedido.Codigo
            _Pedido = pPedido
        End If
    End Sub
#End Region

#Region "Fields"
    Private _Pedido As Pedido
    Private _CodigoEmpresa As String
    Private _EnderecoEmpresa As Integer
    Private _CodigoPedido As Integer
    Private _CodigoCliente As String
    Private _EnderecoCliente As Integer
    Private _TipoTitulo As String
    Private _CodigoMoeda As Integer
    Private _Moeda As Moeda
    Private _Classificacao As String
    Private _NrParcela As Integer
    Private _Movimento As Date
    Private _Vencimento As Date
    Private _ValorOficial As Decimal
    Private _ValorMoeda As Decimal
#End Region

#Region "Property"
    Public ReadOnly Property Pedido As Pedido
        Get
            If _Pedido Is Nothing Then _Pedido = New Pedido(Me.CodigoEmpresa, Me.EnderecoEmpresa, Me.CodigoPedido)
            Return _Pedido
        End Get
    End Property
    Public Property CodigoEmpresa As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(value As String)
            _CodigoEmpresa = value
            _Pedido = Nothing
        End Set
    End Property
    Public Property EnderecoEmpresa As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(value As Integer)
            _EnderecoEmpresa = value
            _Pedido = Nothing
        End Set
    End Property
    Public Property CodigoPedido As Integer
        Get
            Return _CodigoPedido
        End Get
        Set(value As Integer)
            _CodigoPedido = value
            _Pedido = Nothing
        End Set
    End Property
    Public Property CodigoCliente As String
        Get
            Return _CodigoCliente
        End Get
        Set(value As String)
            _CodigoCliente = value
        End Set
    End Property
    Public Property EnderecoCliente As Integer
        Get
            Return _EnderecoCliente
        End Get
        Set(value As Integer)
            _EnderecoCliente = value
        End Set
    End Property
    Public Property TipoTitulo As String
        Get
            Return _TipoTitulo
        End Get
        Set(value As String)
            _TipoTitulo = value
        End Set
    End Property
    Public Property CodigoMoeda As Integer
        Get
            Return _CodigoMoeda
        End Get
        Set(value As Integer)
            _CodigoMoeda = value
        End Set
    End Property
    Public ReadOnly Property Moeda As Moeda
        Get
            If _Moeda Is Nothing Then _Moeda = New Moeda(Me.CodigoMoeda)
            Return _Moeda
        End Get
    End Property
    Public ReadOnly Property DescricaoMoeda
        Get
            Return Me.Moeda.Descricao
        End Get
    End Property
    Public Property Classificacao As String
        Get
            Return _Classificacao
        End Get
        Set(value As String)
            _Classificacao = value
        End Set
    End Property
    Public Property NrParcela As Integer
        Get
            Return _NrParcela
        End Get
        Set(value As Integer)
            _NrParcela = value
        End Set
    End Property
    Public Property Movimento As Date
        Get
            Return _Movimento
        End Get
        Set(value As Date)
            _Movimento = value
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

