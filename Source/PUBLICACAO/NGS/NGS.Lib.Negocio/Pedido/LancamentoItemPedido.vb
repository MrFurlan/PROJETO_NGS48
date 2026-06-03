Imports System.Web.UI.WebControls
Imports Microsoft.VisualBasic
Imports NGS.Lib.Uteis

'******************************************************************************************************************************************************************
'****************************************************  Lista Lancamentos Normal Complemento Estorno do Pedido  ****************************************************
'******************************************************************************************************************************************************************
Public Class ListLancamentoItemPedido
    Inherits List(Of LancamentoItemPedido)

#Region "Construtor"
    Public Sub New(pItemPedido As PedidoXItem)
        _ItemPedido = pItemPedido
        If pItemPedido.Pedido.Codigo = 0 Then Exit Sub

        Dim sql As String = ""
        sql = "SELECT pxi.Produto_Id," & vbCrLf &
              "       pxi.PedidoItem_Id," & vbCrLf &
              "       pxi.TipoDeLancamento," & vbCrLf &
              "       pxi.Movimento," & vbCrLf &
              "       isnull(pxi.DataEntrega, pxi.Movimento) as DataEntrega," & vbCrLf &
              "       isnull(pxi.QuantidadeComercializacao,pxi.quantidade) as QuantidadeComercializacao," & vbCrLf &
              "       pxi.Quantidade," & vbCrLf &
              "       pxi.UnitarioOficial," & vbCrLf &
              "       pxi.UnitarioMoeda," & vbCrLf &
              "       pxi.TotalOficial," & vbCrLf &
              "       pxi.TotalMoeda," & vbCrLf &
              "       isnull(pxi.UnitarioOficialCompra,0) as UnitarioOficialCompra," & vbCrLf &
              "       isnull(pxi.UnitarioMoedaCompra,0) as UnitarioMoedaCompra,  " & vbCrLf &
              "       isnull(pxi.UsuarioLiberacao,'') AS UsuarioLiberacao," & vbCrLf &
              "       isnull(pxi.UsuarioLiberacaoData,CAST(GETDATE() AS DATETIME)) AS UsuarioLiberacaoData" & vbCrLf &
              "  FROM PedidoXItemXLancamento pxi" & vbCrLf &
              " Inner Join Produtos p" & vbCrLf &
              "    on P.Produto_id = pxi.Produto_id" & vbCrLf &
              " WHERE pxi.Empresa_Id    ='" & pItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf &
              "   AND pxi.EndEmpresa_Id = " & pItemPedido.Pedido.EnderecoEmpresa & vbCrLf &
              "   AND pxi.Pedido_Id     = " & pItemPedido.Pedido.Codigo & vbCrLf &
              "   AND pxi.Produto_Id    = '" & pItemPedido.CodigoProduto & "'" & vbCrLf &
              " ORDER BY pxi.PedidoItem_Id" & vbCrLf

        Dim ds As DataSet
        Dim Banco As New AcessaBanco

        ds = Banco.ConsultaDataSet(sql, "LancamentosPedido")

        For Each row In ds.Tables(0).Rows
            Dim Lan As New LancamentoItemPedido(pItemPedido)

            Lan.CodigoPedidoItem = row("PedidoItem_Id")
            Lan.TipoLancamento = Conversoes.ConverterTipoLancamento(row("TipoDeLancamento"))
            Lan.Movimento = row("Movimento")
            Lan.DataEntrega = row("DataEntrega")
            Lan.QuantidadeComercializacao = row("QuantidadeComercializacao")
            Lan.QuantidadeFaturamento = row("Quantidade")
            Lan.UnitarioOficial = row("UnitarioOficial")
            Lan.UnitarioMoeda = row("UnitarioMoeda")
            Lan.TotalOficial = row("TotalOficial")
            Lan.TotalMoeda = row("TotalMoeda")
            Lan.UnitarioOficialCompra = row("UnitarioOficialCompra")
            Lan.UnitarioMoedaCompra = row("UnitarioMoedaCompra")
            Lan.CalculadoPedido = True
            Lan.UsuarioLiberacao = row("UsuarioLiberacao")
            Lan.UsuarioLiberacaoData = Convert.ToDateTime(row("UsuarioLiberacaoData"))
            Me.Add(Lan)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _ItemPedido As PedidoXItem
#End Region

#Region "Property"
    Public ReadOnly Property LancamentoNormal As LancamentoItemPedido
        Get
            Return Me.Where(Function(s) s.TipoLancamento = eTiposLancamentosPedidos.Normal).FirstOrDefault
        End Get
    End Property

    Public ReadOnly Property ItemPedido As PedidoXItem
        Get
            Return _ItemPedido
        End Get
    End Property

    Public ReadOnly Property QuantidadeTotalPrdFaturamento() As Decimal
        Get
            Return Me.Sum(Function(x)
                              Dim ret As Decimal
                              If (x.TipoLancamento = eTiposLancamentosPedidos.Estorno) Then
                                  ret += (x.QuantidadeFaturamento * CDec(-1))
                              Else
                                  ret += x.QuantidadeFaturamento
                              End If
                              Return ret
                          End Function)
        End Get
    End Property

    Public ReadOnly Property QuantidadeTotalPrdComercializacao() As Decimal
        Get
            Return Me.Sum(Function(x)
                              Dim ret As Decimal
                              If (x.TipoLancamento = eTiposLancamentosPedidos.Estorno) Then
                                  ret += x.QuantidadeComercializacao * CDec(-1)
                              Else
                                  ret += x.QuantidadeComercializacao
                              End If
                              Return ret
                          End Function)
        End Get
    End Property

    Public ReadOnly Property TotalOficialPrd() As Decimal
        Get
            Return Me.Sum(Function(x)
                              Dim ret As Decimal
                              If (x.TipoLancamento = eTiposLancamentosPedidos.Estorno) Then
                                  ret += x.TotalOficial * CDec(-1)
                              Else
                                  ret += x.TotalOficial
                              End If
                              Return ret
                          End Function)
        End Get
    End Property

    Public ReadOnly Property TotalMoedaPrd() As Decimal
        Get
            Return Me.Sum(Function(x)
                              Dim ret As Decimal
                              If (x.TipoLancamento = eTiposLancamentosPedidos.Estorno) Then
                                  ret += x.TotalMoeda * CDec(-1)
                              Else
                                  ret += x.TotalMoeda
                              End If
                              Return ret
                          End Function)
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function AdicionarLancamento(pLancamentoItemPedido As LancamentoItemPedido) As ArrayList
        'Retorno devolve na posicao 0 True ou False e na posicao 1 retorna uma MSG
        Dim Retorno As New ArrayList
        '************************************************
        '******  Lançamento Normal ou Complemento  ******
        '************************************************
        If pLancamentoItemPedido.TipoLancamento = eTiposLancamentosPedidos.Normal Or pLancamentoItemPedido.TipoLancamento = eTiposLancamentosPedidos.Complemento Then
            If pLancamentoItemPedido.ItemPedido.SaldoItem.Tipo = 1 Then
                pLancamentoItemPedido.ItemPedido.SaldoItem.QtdeProgramada += pLancamentoItemPedido.QuantidadeFaturamento
                pLancamentoItemPedido.ItemPedido.SaldoItem.QtdeContratadoFiscal += pLancamentoItemPedido.QuantidadeFaturamento
                pLancamentoItemPedido.ItemPedido.SaldoItem.QtdeProgramadaComercializacao += pLancamentoItemPedido.QuantidadeComercializacao

                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeGlobalFiscal += pLancamentoItemPedido.QuantidadeFaturamento
                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeDiretoFiscal += pLancamentoItemPedido.QuantidadeFaturamento

                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorOficialGlobalDireto += pLancamentoItemPedido.TotalOficial
                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorMoedaGlobalDireto += pLancamentoItemPedido.TotalMoeda
            End If
            Me.Add(pLancamentoItemPedido)
            Retorno.Add(True)
            Retorno.Add("")
            '**********************************************************
            '*********   ESTORNOS   ***********************************
            '**********************************************************
            '****   1 - Pedidos Entrega Futura ou Direta   ************
            '**********************************************************
        ElseIf pLancamentoItemPedido.TipoLancamento = eTiposLancamentosPedidos.Estorno AndAlso
            pLancamentoItemPedido.ItemPedido.Pedido.SubOperacao.PrecoFixo AndAlso
            pLancamentoItemPedido.ItemPedido.Pedido.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS AndAlso
            pLancamentoItemPedido.ItemPedido.Pedido.SubOperacao.Classe <> eClassesOperacoes.AFIXAR Then
            'SaldoQtdeGlobalFiscal e SaldoQtdeDiretoFiscal sempre representarao o mesmo valor pois o usuario pode escolher fazer uma venda direta ou uma futura com o saldo disponivel do pedido
            If pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeGlobalFiscal < pLancamentoItemPedido.QuantidadeFaturamento Then
                Retorno.Add(False)
                Retorno.Add("A Quantidade do Estorno " & pLancamentoItemPedido.QuantidadeFaturamento.ToString & " é Maior que o Saldo de Quantidade " & pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeGlobalFiscal.ToString & " do Produto")
            ElseIf Math.Round(IIf(pLancamentoItemPedido.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorOficialGlobalDireto, pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorMoedaGlobalDireto), 2, MidpointRounding.AwayFromZero) < IIf(pLancamentoItemPedido.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, pLancamentoItemPedido.TotalOficial, pLancamentoItemPedido.TotalMoeda) Then
                Retorno.Add(False)
                If pLancamentoItemPedido.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    Retorno.Add("O Valor do Estorno " & pLancamentoItemPedido.TotalOficial.ToString & " é Maior que o Saldo de Valor " & pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorOficialGlobalDireto.ToString & " do Produto")
                Else
                    Retorno.Add("O Valor do Estorno " & pLancamentoItemPedido.TotalMoeda.ToString & " é Maior que o Saldo de Valor " & pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorMoedaGlobalDireto.ToString & " do Produto")
                End If
            Else
                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeGlobalFiscal -= pLancamentoItemPedido.QuantidadeFaturamento
                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeComercializacao -= pLancamentoItemPedido.QuantidadeComercializacao

                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorOficialGlobalDireto -= pLancamentoItemPedido.TotalOficial
                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorMoedaGlobalDireto -= pLancamentoItemPedido.TotalMoeda

                Me.Add(pLancamentoItemPedido)
                Retorno.Add(True)
                Retorno.Add("")
            End If

            '************************************************
            '****   2 - Deposito   **************************
            '************************************************
        ElseIf pLancamentoItemPedido.TipoLancamento = eTiposLancamentosPedidos.Estorno AndAlso
               (pLancamentoItemPedido.ItemPedido.Pedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS Or pLancamentoItemPedido.ItemPedido.Pedido.SubOperacao.Classe = eClassesOperacoes.INDUSTRIALIZACAO) Then

            If pLancamentoItemPedido.ItemPedido.QuantidadePedidoFaturamento < pLancamentoItemPedido.QuantidadeFaturamento Then
                Retorno.Add(False)
                Retorno.Add("A Quantidade do Estorno é Maior que a Quantidade Estimada a do Produto no Pedido")
            ElseIf Math.Round(pLancamentoItemPedido.ItemPedido.PedidoValor, 2, MidpointRounding.AwayFromZero) < IIf(pLancamentoItemPedido.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, pLancamentoItemPedido.TotalOficial, pLancamentoItemPedido.TotalMoeda) Then
                Retorno.Add(False)
                Retorno.Add("O Valor do Estorno é Maior que o  Valor Estimado do Produto no Pedido")
            Else
                Me.Add(pLancamentoItemPedido)
                Retorno.Add(True)
                Retorno.Add("")
            End If
            '************************************************
            '****   3 - A FIxar   ***************************
            '************************************************
        ElseIf pLancamentoItemPedido.TipoLancamento = eTiposLancamentosPedidos.Estorno AndAlso
               pLancamentoItemPedido.ItemPedido.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then

            If pLancamentoItemPedido.ItemPedido.QuantidadePedidoFaturamento < pLancamentoItemPedido.QuantidadeFaturamento Then
                Retorno.Add(False)
                Retorno.Add("A Quantidade do Estorno é Maior que a Quantidade Estimada a do Produto no Pedido")
            ElseIf Math.Round(pLancamentoItemPedido.ItemPedido.PedidoValor, 2, MidpointRounding.AwayFromZero) < IIf(pLancamentoItemPedido.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, pLancamentoItemPedido.TotalOficial, pLancamentoItemPedido.TotalMoeda) Then
                Retorno.Add(False)
                Retorno.Add("O Valor do Estorno é Maior que o  Valor Estimado do Produto no Pedido")
            ElseIf pLancamentoItemPedido.ItemPedido.QuantidadePedidoFaturamento - pLancamentoItemPedido.ItemPedido.SaldoItem.QtdeEntregueFiscalAFixar < pLancamentoItemPedido.QuantidadeFaturamento Then
                Retorno.Add(False)
                Retorno.Add("A Quantidade do Estorno fara com que a quantidade ja fixada fique maior que a quantidade do pedido")
            Else
                Me.Add(pLancamentoItemPedido)
                Retorno.Add(True)
                Retorno.Add("")
            End If

        ElseIf pLancamentoItemPedido.TipoLancamento = eTiposLancamentosPedidos.Estorno AndAlso
               (pLancamentoItemPedido.ItemPedido.Pedido.SubOperacao.Classe = eClassesOperacoes.OUTRAS OrElse pLancamentoItemPedido.ItemPedido.Pedido.SubOperacao.Classe = eClassesOperacoes.COMPRASGERAIS) Then
            'SaldoQtdeGlobalFiscal e SaldoQtdeDiretoFiscal sempre representarao o mesmo valor pois o usuario pode escolher fazer uma venda direta ou uma futura com o saldo disponivel do pedido
            If pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeGlobalFiscal < pLancamentoItemPedido.QuantidadeFaturamento Then
                Retorno.Add(False)
                Retorno.Add("A Quantidade do Estorno é Maior que o Saldo de Quantidade do Produto")
            ElseIf Math.Round(IIf(pLancamentoItemPedido.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorOficialGlobalDireto, pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorMoedaGlobalDireto), 2, MidpointRounding.AwayFromZero) < IIf(pLancamentoItemPedido.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, pLancamentoItemPedido.TotalOficial, pLancamentoItemPedido.TotalMoeda) Then
                Retorno.Add(False)
                Retorno.Add("O Valor do Estorno é Maior que o Saldo de Valor do Produto")
            Else
                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeGlobalFiscal -= pLancamentoItemPedido.QuantidadeFaturamento
                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeComercializacao -= pLancamentoItemPedido.QuantidadeComercializacao

                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorOficialGlobalDireto -= pLancamentoItemPedido.TotalOficial
                pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorMoedaGlobalDireto -= pLancamentoItemPedido.TotalMoeda

                Me.Add(pLancamentoItemPedido)
                Retorno.Add(True)
                Retorno.Add("")
            End If
        Else
            Throw New InvalidOperationException("Classe da Operação não está sendo tratada, entre em contato com o Suporte")
        End If


        Return Retorno
    End Function

    Public Function AdicionarLancamentoEstorno(pLancamentoItemPedido As LancamentoItemPedido) As ArrayList
        'CRIADO PARA LIBERAR ESTORNO VINDO DA TELA DO AJUSTAR PEDIDO - FURLAN - 29/05/2025

        'Retorno devolve na posicao 0 True ou False e na posicao 1 retorna uma MSG
        Dim Retorno As New ArrayList

        If Funcoes.VerificaPermissao("AjustarPedido", "LIBERAR") Then
            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeGlobalFiscal -= pLancamentoItemPedido.QuantidadeFaturamento
            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeComercializacao -= pLancamentoItemPedido.QuantidadeComercializacao

            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorOficialGlobalDireto -= pLancamentoItemPedido.TotalOficial
            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorMoedaGlobalDireto -= pLancamentoItemPedido.TotalMoeda

            Me.Add(pLancamentoItemPedido)
            Retorno.Add(True)
            Retorno.Add("")
        Else
        End If
        Return Retorno
    End Function

    Public Sub RemoverLancamento(pLancamentoItemPedido As LancamentoItemPedido)
        'Retorno devolve na posicao 0 True ou False e na posicao 1 retorna uma MSG
        Dim Retorno As New ArrayList
        If pLancamentoItemPedido.TipoLancamento = eTiposLancamentosPedidos.Complemento Then
            pLancamentoItemPedido.ItemPedido.SaldoItem.QtdeProgramada -= pLancamentoItemPedido.QuantidadeFaturamento
            pLancamentoItemPedido.ItemPedido.SaldoItem.QtdeContratadoFiscal -= pLancamentoItemPedido.QuantidadeFaturamento
            pLancamentoItemPedido.ItemPedido.SaldoItem.QtdeProgramadaComercializacao -= pLancamentoItemPedido.QuantidadeComercializacao

            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeGlobalFiscal -= pLancamentoItemPedido.QuantidadeFaturamento
            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeDiretoFiscal -= pLancamentoItemPedido.QuantidadeFaturamento

            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorOficialGlobalDireto -= pLancamentoItemPedido.TotalOficial
            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorMoedaGlobalDireto -= pLancamentoItemPedido.TotalMoeda
        ElseIf pLancamentoItemPedido.TipoLancamento = eTiposLancamentosPedidos.Estorno AndAlso _
               pLancamentoItemPedido.ItemPedido.Pedido.SubOperacao.PrecoFixo AndAlso _
               pLancamentoItemPedido.ItemPedido.Pedido.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS AndAlso _
               pLancamentoItemPedido.ItemPedido.Pedido.SubOperacao.Classe <> eClassesOperacoes.AFIXAR Then

            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeGlobalFiscal += pLancamentoItemPedido.QuantidadeFaturamento
            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoQtdeComercializacao += pLancamentoItemPedido.QuantidadeComercializacao

            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorOficialGlobalDireto += pLancamentoItemPedido.TotalOficial
            pLancamentoItemPedido.ItemPedido.SaldoItem.SaldoValorMoedaGlobalDireto += pLancamentoItemPedido.TotalMoeda
        End If

        Me.Remove(pLancamentoItemPedido)
    End Sub

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Lan As LancamentoItemPedido In Me
            If ItemPedido.IUD = "D" Or ItemPedido.IUD = "I" Then Lan.IUD = ItemPedido.IUD
            Lan.SalvarSql(Sqls)
        Next
    End Sub

#End Region

End Class

'******************************************************************************************************************************************************************
'*****************************************************************  Classe Base do Item do Pedido  ****************************************************************
'******************************************************************************************************************************************************************
Public Class LancamentoItemPedido
#Region "Contrututor"
    Public Sub New(pItemPedido As PedidoXItem)
        _ItemPedido = pItemPedido
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _ItemPedido As PedidoXItem
    Private _CodigoPedidoItem As Integer
    Private _TipoLancamento As eTiposLancamentosPedidos
    Private _Movimento As DateTime
    Private _DataEntrega As DateTime
    Private _QuantidadeComercializacao As Decimal
    Private _QuantidadeFaturamento As Decimal
    Private _UnitarioOficial As Decimal
    Private _UnitarioMoeda As Decimal
    Private _TotalOficial As Decimal
    Private _TotalMoeda As Decimal
    Private _UnitarioOficialCompra As Decimal
    Private _UnitarioMoedaCompra As Decimal
    Private _IndiceFixado As Decimal
    Private _CalculadoPedido As Boolean
    Private _UsuarioLiberacao As String
    Private _UsuarioLiberacaoData As DateTime
#End Region

#Region "Property"
    Public ReadOnly Property ItemPedido As PedidoXItem
        Get
            Return _ItemPedido
        End Get
    End Property

    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property

    Public Property CalculadoPedido() As Boolean
        Get
            Return _CalculadoPedido
        End Get
        Set(ByVal value As Boolean)
            _CalculadoPedido = value
        End Set
    End Property

    Public Property CodigoPedidoItem() As Integer
        Get
            Return _CodigoPedidoItem
        End Get
        Set(ByVal value As Integer)
            _CodigoPedidoItem = value
        End Set
    End Property

    Public Property TipoLancamento() As eTiposLancamentosPedidos
        Get
            Return _TipoLancamento
        End Get
        Set(ByVal value As eTiposLancamentosPedidos)
            _TipoLancamento = value
        End Set
    End Property

    Public Property Movimento() As DateTime
        Get
            Return _Movimento
        End Get
        Set(ByVal value As DateTime)
            _Movimento = value
        End Set
    End Property

    Public Property DataEntrega() As DateTime
        Get
            Return _DataEntrega
        End Get
        Set(ByVal value As DateTime)
            _DataEntrega = value
        End Set
    End Property


    '**************************************************************************************
    '***** Propriedade que segue a unidade de Comercializacao do produto ******************
    '**************************************************************************************

    Public Property QuantidadeComercializacao As Decimal
        Get
            Return _QuantidadeComercializacao
        End Get
        Set(value As Decimal)
            _QuantidadeComercializacao = value
        End Set
    End Property

    'Unitario Comercializacao Segue a Moeda Do Pedido
    Public ReadOnly Property UnitarioComercializacao As Decimal
        Get
            If ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Return _UnitarioOficial * ItemPedido.UnidadeComercializacaoFatorDeConversao
            Else
                Return _UnitarioMoeda * ItemPedido.UnidadeComercializacaoFatorDeConversao
            End If
        End Get
    End Property

    Public ReadOnly Property UnitarioOficialComercializacao() As Decimal
        Get
            Return _UnitarioOficial * ItemPedido.UnidadeComercializacaoFatorDeConversao
        End Get
    End Property

    Public ReadOnly Property UnitarioMoedaComercializacao() As Decimal
        Get
            Return _UnitarioMoeda * ItemPedido.UnidadeComercializacaoFatorDeConversao
        End Get
    End Property

    'Unitario Compra Comercializacao Segue a Moeda Do Pedido
    Public ReadOnly Property UnitarioCompraComercializacao As Decimal
        Get
            If ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Return _UnitarioOficialCompra * ItemPedido.UnidadeComercializacaoFatorDeConversao
            Else
                Return _UnitarioMoedaCompra * ItemPedido.UnidadeComercializacaoFatorDeConversao
            End If
        End Get
    End Property

    Public ReadOnly Property UnitarioOficialCompraComercializacao() As Decimal
        Get
            Return _UnitarioOficialCompra * ItemPedido.UnidadeComercializacaoFatorDeConversao
        End Get
    End Property

    Public ReadOnly Property UnitarioMoedaCompraComercializacao() As Decimal
        Get
            Return _UnitarioMoedaCompra * ItemPedido.UnidadeComercializacaoFatorDeConversao
        End Get
    End Property

    '**********************************************************************************************************
    '***** Propriedade que segue a unidade de medida de Faturamento do Cadastro do produto ********************
    '**********************************************************************************************************
    Public Property QuantidadeFaturamento() As Decimal
        Get
            Return _QuantidadeFaturamento
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeFaturamento = value
        End Set
    End Property

    Public ReadOnly Property UnitarioFaturamento As Decimal
        Get
            If ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Return _UnitarioOficial
            Else
                Return _UnitarioMoeda
            End If
        End Get
    End Property

    Public Property UnitarioOficial() As Decimal
        Get
            Return _UnitarioOficial
        End Get
        Set(ByVal value As Decimal)
            _UnitarioOficial = value
        End Set
    End Property

    Public Property UnitarioMoeda() As Decimal
        Get
            Return _UnitarioMoeda
        End Get
        Set(ByVal value As Decimal)
            _UnitarioMoeda = value
        End Set
    End Property

    'Segue a Moeda do Pedido
    Public ReadOnly Property UnitarioCompra As Decimal
        Get
            If ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Return Me.UnitarioOficialCompra
            Else
                Return Me.UnitarioMoedaCompra
            End If
        End Get
    End Property

    Public Property UnitarioOficialCompra() As Decimal
        Get
            Return _UnitarioOficialCompra
        End Get
        Set(ByVal value As Decimal)
            _UnitarioOficialCompra = value
        End Set
    End Property

    Public Property UnitarioMoedaCompra() As Decimal
        Get
            Return _UnitarioMoedaCompra
        End Get
        Set(ByVal value As Decimal)
            _UnitarioMoedaCompra = value
        End Set
    End Property

    '*******************************************************************************************
    '******************  Segue a Moeda e a Unidade de comercializacao do Pedido ****************
    '*******************************************************************************************
    Public ReadOnly Property QuantidadePedido As Decimal
        Get
            If ItemPedido.Produto.Unidade <> ItemPedido.CodigoUnidadeComercializacao Then
                Return Me.QuantidadeComercializacao
            Else
                Return Me.QuantidadeFaturamento
            End If
        End Get
    End Property

    Public ReadOnly Property Unitario As Decimal
        Get
            Dim Unt As Decimal
            If ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                If QuantidadeComercializacao = 0 Then
                    'Unt = ItemPedido.Lancamentos(0).UnitarioOficial * ItemPedido.Lancamentos(0).UnidadeComercializacaoFatorDeConversao
                    Unt = UnitarioOficial * ItemPedido.UnidadeComercializacaoFatorDeConversao
                Else
                    Unt = (TotalOficial / QuantidadeComercializacao)
                End If

                If Math.Round(QuantidadeComercializacao * Math.Round(Unt, 0, MidpointRounding.AwayFromZero), 2, MidpointRounding.AwayFromZero) = TotalOficial Then
                    If TotalOficial = 0 Then
                        Return Unt
                    Else
                        Return Math.Round(Unt, 0, MidpointRounding.AwayFromZero)
                    End If
                ElseIf QuantidadeComercializacao = 0 Then
                    Return Unt
                Else
                    Return (TotalOficial / QuantidadeComercializacao)
                End If
                'Return Me.UnitarioOficial * Me.UnidadeComercializacaoFatorDeConversao
            Else
                If QuantidadeComercializacao = 0 Then
                    'Unt = ItemPedido.Lancamentos(0).UnitarioMoeda * ItemPedido.Lancamentos(0).UnidadeComercializacaoFatorDeConversao
                    Unt = UnitarioMoeda * ItemPedido.UnidadeComercializacaoFatorDeConversao
                Else
                    Unt = (TotalMoeda / QuantidadeComercializacao)
                End If
                If Math.Round(QuantidadeComercializacao * Math.Round(Unt, 0, MidpointRounding.AwayFromZero), 2, MidpointRounding.AwayFromZero) = TotalMoeda Then
                    If TotalMoeda = 0 Then
                        Return Unt
                    Else
                        Return Math.Round(Unt, 0, MidpointRounding.AwayFromZero)
                    End If
                ElseIf QuantidadeComercializacao = 0 Then
                    Return Unt
                Else
                    Return (TotalMoeda / QuantidadeComercializacao)
                End If
                'Return Me.UnitarioMoeda * Me.UnidadeComercializacaoFatorDeConversao
            End If
        End Get
    End Property

    Public ReadOnly Property Total As Decimal
        Get
            If ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Return Me.TotalOficial
            Else
                Return Me.TotalMoeda
            End If
        End Get
    End Property

    '*******************************************************************************************
    '******************  TOTAIS ****************************************************************
    '*******************************************************************************************
    Public Property TotalOficial() As Decimal
        Get
            Return _TotalOficial
        End Get
        Set(ByVal value As Decimal)
            _TotalOficial = value
        End Set
    End Property

    Public Property TotalMoeda() As Decimal
        Get
            Return _TotalMoeda
        End Get
        Set(ByVal value As Decimal)
            _TotalMoeda = value
        End Set
    End Property

    '*******************************************************************************************
    '*******************************************************************************************
    '*******************************************************************************************

    Public Property UsuarioLiberacao() As String
        Get
            Return _UsuarioLiberacao
        End Get
        Set(ByVal value As String)
            _UsuarioLiberacao = value
        End Set
    End Property

    Public Property UsuarioLiberacaoData() As DateTime
        Get
            Return _UsuarioLiberacaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioLiberacaoData = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case Me.IUD
            Case "I", "U"
                sql = " MERGE PedidoXItemXLancamento AS Dest" & vbCrLf & _
                      " USING (Select '" & Me.ItemPedido.Pedido.CodigoEmpresa & "' as Empresa_Id," & Me.ItemPedido.Pedido.EnderecoEmpresa & " as EndEmpresa_Id," & Me.ItemPedido.Pedido.Codigo & " as Pedido_Id," & Me.ItemPedido.CodigoProduto & " as Produto_Id ," & Me.CodigoPedidoItem & " as PedidoItem_Id) AS Ori" & vbCrLf & _
                      "    ON Dest.Empresa_Id    = Ori.Empresa_Id" & vbCrLf & _
                      "   and Dest.EndEmpresa_Id = Ori.EndEmpresa_Id" & vbCrLf & _
                      "   and Dest.Pedido_Id     = Ori.Pedido_Id" & vbCrLf & _
                      "   and Dest.Produto_Id    = Ori.Produto_Id" & vbCrLf & _
                      "   and Dest.PedidoItem_Id = Ori.PedidoItem_Id" & vbCrLf & _
                      "  WHEN NOT MATCHED" & vbCrLf & _
                      "    THEN Insert (Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, PedidoItem_Id," & vbCrLf & _
                      "                 TipoDeLancamento, Movimento, DataEntrega," & vbCrLf & _
                      "	   			    QuantidadeComercializacao, Quantidade," & vbCrLf & _
                      "                 UnitarioOficial, UnitarioMoeda, TotalOficial, TotalMoeda," & vbCrLf & _
                      "                 UnitarioOficialCompra, UnitarioMoedaCompra," & vbCrLf & _
                      "				    UsuarioLiberacao, UsuarioLiberacaoData)" & vbCrLf & _
                      "		    values ('" & Me.ItemPedido.Pedido.CodigoEmpresa & "', " & Me.ItemPedido.Pedido.EnderecoEmpresa & "," & Me.ItemPedido.Pedido.Codigo & ",'" & Me.ItemPedido.CodigoProduto & "', " & Me.CodigoPedidoItem & "," & vbCrLf & _
                      "'" & Me.TipoLancamento.ToString().Substring(0, 1) & "','" & Me.Movimento.ToString("yyyy-MM-dd") & "','" & Me.DataEntrega.ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                      Str(Me.QuantidadeComercializacao) & "," & Str(Me.QuantidadeFaturamento) & "," & vbCrLf & _
                      Str(Me.UnitarioOficial) & "," & Str(Me.UnitarioMoeda) & "," & Str(Me.TotalOficial) & "," & Str(Me.TotalMoeda) & "," & vbCrLf & _
                      Str(Me.UnitarioOficialCompra) & "," & Str(Me.UnitarioMoedaCompra) & "," & vbCrLf & _
                      "'" & Me.UsuarioLiberacao & "', CURRENT_TIMESTAMP) " & vbCrLf & _
                      "  WHEN MATCHED " & vbCrLf & _
                      "    THEN Update set " & vbCrLf & _
                      "          QuantidadeComercializacao = " & Str(Me.QuantidadeComercializacao) & vbCrLf & _
                      "         ,Quantidade                = " & Str(Me.QuantidadeFaturamento) & vbCrLf & _
                      "         ,UnitarioOficial           = " & Str(Me.UnitarioOficial) & vbCrLf & _
                      "         ,UnitarioMoeda             = " & Str(Me.UnitarioMoeda) & vbCrLf & _
                      "         ,TotalOficial              = " & Str(Me.TotalOficial) & vbCrLf & _
                      "         ,TotalMoeda                = " & Str(Me.TotalMoeda) & vbCrLf & _
                      "         ,TipoDeLancamento          ='" & Me.TipoLancamento.ToString().Substring(0, 1) & "'" & vbCrLf & _
                      "         ,Movimento                 ='" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "         ,DataEntrega               ='" & Me.DataEntrega.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "         ,UnitarioOficialCompra     = " & Str(Me.UnitarioOficialCompra) & vbCrLf & _
                      "         ,UnitarioMoedaCompra       = " & Str(Me.UnitarioMoedaCompra) & vbCrLf & _
                      "         ,UsuarioLiberacao          ='" & Me.UsuarioLiberacao & "'" & vbCrLf & _
                      "         ,UsuarioLiberacaoData      = CURRENT_TIMESTAMP;" & vbCrLf
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE PedidoXItemXLancamento " & _
                      " WHERE Empresa_Id    ='" & Me.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & Me.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                      "   AND Pedido_Id     = " & Me.ItemPedido.Pedido.Codigo & vbCrLf & _
                      "   AND Produto_Id    ='" & Me.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                      "   AND PedidoItem_id = " & Me.CodigoPedidoItem & ";" & vbCrLf

                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class

