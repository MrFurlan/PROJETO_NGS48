Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListNotaFiscalXItem
    Inherits List(Of NotaFiscalXItem)
    Implements IBaseEntity

#Region "Fields"
    Private _NotaFiscal As NotaFiscal
#End Region

#Region "Contrutor"
    Public Sub New(ByVal objNotaFiscal As NotaFiscal, Optional ByVal Carregar As Boolean = True)
        Me.NotaFiscal = objNotaFiscal
        If Not Carregar Then Exit Sub
        Dim db As New AcessaBanco()

        Try
            Dim strSQL As String
            strSQL = "SELECT NFXI.Produto_Id, OE.CodigoFiscal, NFXI.Sequencia_id, isnull(NFXI.lote,'') as Lote, isnull(NFXI.classificacao,'') as classificacao,  isnull(NFXI.Pedido,0) as Pedido, isnull(NFXI.Deposito,'')as Deposito, isnull(NFXI.EndDeposito,0)as EndDeposito, isnull(NFXI.DepositoTerceiro,'')as DepositoTerceiro, isnull(NFXI.EndDepositoTerceiro,0) as EndDepositoTerceiro, " & vbCrLf &
                     "       NFXI.PesoFiscal, NFXI.QuantidadeFisica, NFXI.QuantidadeFiscal," & vbCrLf &
                     "       NFXI.Unitario, NFXI.Valor, isnull(NFXI.ValorLiquido,0) as ValorLiquido, isnull(NFXI.ValorMoeda,0) as ValorMoeda,  isnull(NFXI.ValorliquidoMoeda,0) as ValorliquidoMoeda, " & vbCrLf &
                     "       NFXI.PesoQuantidade, isnull(NFXI.Operacao," & Me.NotaFiscal.CodigoOperacao & ") as Operacao, " & vbCrLf &
                     "       isnull(NFXI.SubOperacao," & Me.NotaFiscal.CodigoSubOperacao & ") as SubOperacao, isnull(NFXI.ObservacoesDoProduto,'') as ObservacoesDoProduto, " & vbCrLf &
                     "       isnull(NFXI.Embalagem,0) as Embalagem, isnull(NFXI.TipoDeEmbalagem,'') as TipoDeEmbalagem, isnull(NFXI.CapacidadeEmbalagem,0) as CapacidadeEmbalagem, isnull(NFXI.QtdeDeEmbalagem,0) as QtdeDeEmbalagem, isnull(Emb.EmbalagemIndea,'') as EmbalagemIndea, Rateado," & vbCrLf &
                     "       isnull(Fabricante,'') as Fabricante, isnull(EndFabricante,0) as EndFabricante, isnull(DataDesembarqueImportacao,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) as DataDesembarqueImportacao, isnull(EstadoDesembarqueImportacao,'') as EstadoDesembarqueImportacao, isnull(LocalDesembarqueImportacao,'') as LocalDesembarqueImportacao, isnull(DataRegistroDI,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) as DataRegistroDI, " & vbCrLf &
                     "       isnull(NumeroDeclaracaoImportacao,'') as NumeroDeclaracaoImportacao, isnull(NumeroPecas,0) as NumeroPecas, isnull(Fixacao,0) as Fixacao, isnull(Plano.CentroDeCusto,'N') as CentroDeCusto," & vbCrLf &
                     "       case" & vbCrLf &
                     "         when ValorMoeda > 0" & vbCrLf &
                     "           then NFxi.Valor / NFxi.valormoeda" & vbCrLf &
                     "           else 0" & vbCrLf &
                     "       end IndiceProdutoNota," & vbCrLf &
                     "       NFXI.OperacaoXEstado," & vbCrLf &
                     "       isnull(pxi.Retencao,1) as Retencao, isnull(ProdutoXML,'') as ProdutoXML, isnull(DescricaoProdutoXML,'') as DescricaoProdutoXML, isnull(InfAdicionalProdutoXML,'') as InfAdicionalProdutoXML, " & vbCrLf &
                     "       isnull(NFxi.ProdutoParaCusto,'') as ProdutoParaCusto" & vbCrLf &
                     "  FROM NotasFiscaisXItens NFXI" & vbCrLf &
                     "  INNER JOIN SubOperacoes SubOp " & vbCrLf &
                     "    ON NFXI.SubOperacao = SubOp.SubOperacoes_Id AND NFXI.Operacao = SubOp.Operacao_Id " & vbCrLf &
                     "  Inner Join OperacaoXEstado OE" & vbCrLf &
                     "    ON OE.Codigo_id = NFXI.OperacaoXEstado" & vbCrLf &
                     "  LEFT JOIN PlanoDeContas Plano " & vbCrLf &
                     "    ON Plano.Conta_Id = SubOp.GrupoDeContas " & vbCrLf &
                     "  LEFT JOIN Embalagens Emb" & vbCrLf &
                     "    ON Emb.Embalagem_Id = NFXI.Embalagem" & vbCrLf &
                     "  Left Join PedidoXItem Pxi" & vbCrLf &
                     "    on pxi.Empresa_id       = NFxI.Empresa_Id" & vbCrLf &
                     "   and pxi.EndEmpresa_id    = NFxI.EndEmpresa_Id" & vbCrLf &
                     "   and pxi.Pedido_id        = NFxI.Pedido" & vbCrLf &
                     "   and pxi.Produto_id       = NFxI.Produto_id" & vbCrLf &
                     " WHERE NFXI.Empresa_Id      ='" & Me.NotaFiscal.CodigoEmpresa & "' " & vbCrLf &
                     "   AND NFXI.EndEmpresa_Id   = " & Me.NotaFiscal.EnderecoEmpresa.ToString() & " " & vbCrLf &
                     "   AND NFXI.Cliente_Id      ='" & Me.NotaFiscal.CodigoCliente & "' " & vbCrLf &
                     "   AND NFXI.EndCliente_Id   = " & Me.NotaFiscal.EnderecoCliente.ToString() & " " & vbCrLf &
                     "   AND NFXI.EntradaSaida_Id ='" & Me.NotaFiscal.EntradaSaida.ToString().Substring(0, 1) & "' " & vbCrLf &
                     "   AND NFXI.Serie_Id        ='" & Me.NotaFiscal.Serie & "'" & vbCrLf &
                     "   AND NFXI.Nota_Id         = " & Me.NotaFiscal.Codigo.ToString() & vbCrLf &
                     "   AND NFXI.Nota_Id         > 0 " & vbCrLf &
                     " ORDER BY NFXI.Sequencia_id; "

            Dim ds As DataSet = db.ConsultaDataSet(strSQL, "NotasFiscaisXItens")

            Me.NotaFiscal.CarregandoItens = True
            For Each row As DataRow In ds.Tables(0).Rows
                Dim nfxi As New NotaFiscalXItem(Me.NotaFiscal)
                nfxi.CodigoProduto = row("Produto_Id").ToString()
                nfxi.CodigoProdutoOld = row("Produto_Id").ToString()
                nfxi.Produto = New Produto(row("Produto_Id").ToString())
                nfxi.CodigoOperacaoEstado = row("OperacaoXEstado")
                nfxi.Lote = row("Lote")
                nfxi.Classificacao = row("Classificacao")
                nfxi.CodigoEmbalagem = row("Embalagem")
                nfxi.CodigoEmbalagemIndea = row("EmbalagemIndea")
                nfxi.CodigoTipoDeEmbalagem = row("TipoDeEmbalagem")
                nfxi.CapacidadeEmbalagem = row("CapacidadeEmbalagem")
                nfxi.QuantidadeDeEmbalagem = row("QtdeDeEmbalagem")
                nfxi.PesoQuantidade = row("PesoQuantidade").ToString()
                nfxi.CFOP = row("CodigoFiscal")
                nfxi.CFOPOld = row("CodigoFiscal")
                nfxi.Sequencia = CInt(row("Sequencia_Id"))
                nfxi.CodigoPedido = CInt(row("Pedido"))
                nfxi.CodigoDeposito = row("Deposito").ToString()
                nfxi.EnderecoDeposito = CInt(row("EndDeposito"))
                nfxi.CodigoDepositoTerceiro = row("DepositoTerceiro").ToString()
                nfxi.EnderecoDepositoTerceiro = CInt(row("EndDepositoTerceiro"))
                nfxi.PesoFiscal = CDbl(row("PesoFiscal"))
                nfxi.QuantidadeFisica = CDbl(row("QuantidadeFisica"))
                nfxi.QuantidadeFiscal = CDbl(row("QuantidadeFiscal"))

                nfxi.ValorTotal = CDbl(row("Valor"))
                nfxi.ValorLiquido = CDbl(row("ValorLiquido"))
                nfxi.Unitario = CDbl(row("Unitario"))

                nfxi.ValorTotalMoeda = CDbl(row("ValorMoeda"))
                nfxi.ValorLiquidoMoeda = CDbl(row("ValorLiquidoMoeda"))

                nfxi.Retencao = row("Retencao")

                If nfxi.QuantidadeFiscal = 0 Then
                    nfxi.UnitarioMoeda = 0
                Else
                    nfxi.UnitarioMoeda = Math.Round(nfxi.ValorTotalMoeda / nfxi.QuantidadeFiscal, 10, MidpointRounding.AwayFromZero)
                End If

                nfxi.IndiceProdutoNota = row("IndiceProdutoNota")

                nfxi.CodigoOperacao = CInt(row("Operacao"))
                nfxi.CodigoSubOperacao = CInt(row("SubOperacao"))

                If row("ObservacoesDoProduto").ToString.Contains("|") Then
                    Dim obs As String() = row("ObservacoesDoProduto").ToString.Split("|")
                    nfxi.ObservacoesDoProduto = obs(0)
                Else
                    nfxi.ObservacoesDoProduto = row("ObservacoesDoProduto")
                End If
                nfxi.Rateado = row("Rateado")
                nfxi.NumeroPecas = row("NumeroPecas")
                nfxi.CodigoFixacao = row("Fixacao")

                nfxi.ProdutoXML = row("ProdutoXML")
                nfxi.DescricaoProdutoXML = row("DescricaoProdutoXML")
                nfxi.InfAdicionalProdutoXML = row("InfAdicionalProdutoXML")
                nfxi.CodigoProdutoCusto = row("ProdutoParaCusto")

                nfxi.CarregandoEncargos = True
                nfxi.Encargos = New ListNotaFiscalXItemXEncargo(nfxi, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                nfxi.CarregandoEncargos = False
                nfxi.TemCentroDeCusto = Not String.IsNullOrWhiteSpace(row("CentroDeCusto")) AndAlso row("CentroDeCusto") = "S"

                nfxi.Lotes = New ListNotaFiscalXLote(nfxi)
                nfxi.Lotes.CarregarLotes()

                nfxi.NotasDevolucao = New ListNotaFiscalDevolucaoXNotaFiscal(nfxi)
                'Estava duplicando a Nota - Furlan - 03/11/2024
                'nfxi.NotasDevolucao.CarregarNotasUsadasNaDevolucao()

                Me.Add(nfxi)
            Next
            Me.NotaFiscal.CarregandoItens = False
        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub
#End Region

#Region "Property"

    Sub New()
        ' TODO: Complete member initialization 
    End Sub

    Public Property NotaFiscal() As NotaFiscal
        Get
            Return _NotaFiscal
        End Get
        Set(ByVal value As NotaFiscal)
            _NotaFiscal = value
        End Set
    End Property

    Public ReadOnly Property ValorBruto_Oficial As Decimal
        Get
            Return (From x In Me Select x.ValorTotal).Sum()
        End Get
    End Property

    Public ReadOnly Property ValorBruto_Moeda As Decimal
        Get
            Return (From x In Me Select x.ValorTotalMoeda).Sum()
        End Get
    End Property

    Public ReadOnly Property ValorLiquido_Oficial As Decimal
        Get
            Return (From x In Me Select x.ValorLiquido).Sum()
        End Get
    End Property

    Public ReadOnly Property ValorLiquido_Moeda As Decimal
        Get
            Return (From x In Me Select x.ValorLiquidoMoeda).Sum()
        End Get
    End Property
#End Region

#Region "Methods"
    Public Shared Function Existe(ByVal Pedido As Integer, ByVal Empresa As String, ByVal EndEmpresa As Integer) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT 1 " & vbCrLf &
                                   "  FROM NotasFiscaisXItens " & vbCrLf &
                                   " INNER JOIN NotasFiscais " & vbCrLf &
                                   "    ON NotasFiscaisXItens.Empresa_Id      = NotasFiscais.Empresa_Id " & vbCrLf &
                                   "   And NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id " & vbCrLf &
                                   "   And NotasFiscaisXItens.Cliente_Id      = NotasFiscais.Cliente_Id " & vbCrLf &
                                   "   And NotasFiscaisXItens.EndCliente_Id   = NotasFiscais.EndCliente_Id " & vbCrLf &
                                   "   And NotasFiscaisXItens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id " & vbCrLf &
                                   "   And NotasFiscaisXItens.Serie_Id        = NotasFiscais.Serie_Id " & vbCrLf &
                                   "   And NotasFiscaisXItens.Nota_Id         = NotasFiscais.Nota_Id " & vbCrLf &
                                   " WHERE (NotasFiscaisXItens.Pedido         = " & Pedido.ToString() & ") " & vbCrLf &
                                   "   And (NotasFiscaisXItens.Empresa_Id     = " & Empresa & ")" & vbCrLf &
                                   "   And (NotasFiscaisXItens.EndEmpresa_Id  = " & EndEmpresa & ")" & vbCrLf &
                                   "   And (NotasFiscais.Situacao = 1)"

            Dim dsItens As DataSet = objBanco.ConsultaDataSet(strSQL, "NotasFiscaisXItens")

            If dsItens.Tables(0).Rows.Count > 0 Then Return True Else Return False
        Catch ex As Exception
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim pSequencia As Integer = 0
        'Notas Fiscais Gerais
        If Me.NotaFiscal.NFG Then
            For Each item As NotaFiscalXItem In Me
                If Not (item.IUD = "D" AndAlso item.NotaFiscal.IUD = "I") Then
                    If NotaFiscal.IUD = "D" OrElse NotaFiscal.IUD = "I" Then item.IUD = NotaFiscal.IUD

                    If Not item.IUD = "D" Then
                        pSequencia += 1
                    End If

                    item.SalvarSql(Sqls, pSequencia)
                End If
            Next
        Else 'Outras Notas Fiscais'
            For Each item As NotaFiscalXItem In Me
                If NotaFiscal.IUD = "D" OrElse NotaFiscal.IUD = "I" Then item.IUD = NotaFiscal.IUD

                pSequencia += 1
                item.SalvarSql(Sqls, pSequencia)
            Next
        End If
    End Sub

    Public Function Clone() As ListNotaFiscalXItem
        Dim itens As ListNotaFiscalXItem = CType(Me.MemberwiseClone(), ListNotaFiscalXItem)
        Return itens
    End Function

#End Region

End Class

<Serializable()> _
Public Class NotaFiscalXItem
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""
    Private _NotaFiscal As NotaFiscal
    Private _CodigoProduto As String
    Private _CodigoProdutoOld As String
    Private _Produto As Produto
    Private _IndiceProdutoNota As Decimal
    Private _Retencao As Boolean = True

    Private _CFOP As Integer
    Private _CFOPOld As Integer
    Private _CodigoPedido As Integer
    Private _Pedido As Pedido

    Private _Sequencia As Integer = 0
    Private _Lote As String = ""
    Private _Classificacao As String = ""
    Private _ProdutoLoteClassificacao As LoteXClassificacao

    Private _codigoEmbalagem As Integer = 0
    Private _CodigoEmbalagemIndea As String = ""
    Private _Embalagem As Embalagem
    Private _CodigoTipoDeEmbalagem As String = ""
    Private _TipoDeEmbalagem As TipoDeEmbalagem
    Private _CapacidadeEmbalagem As Decimal = 0
    Private _QuantidadeDeEmbalagem As Integer = 0
    Private _EmbalagemProduto As ProdutoXEmbalagem

    Private _CodigoDeposito As String
    Private _EnderecoDeposito As Integer
    Private _Deposito As Cliente

    Private _CodigoDepositoTerceiro As String
    Private _EnderecoDepositoTerceiro As Integer
    Private _DepositoTerceiro As Cliente

    Private _PesoFiscal As Decimal = 0
    Private _QuantidadeFisica As Decimal = 0
    Private _QuantidadeFiscal As Decimal = 0

    Private _PesoBruto As Decimal
    Private _PesoLiquido As Decimal
    Private _Volumes As Integer
    Private _Numeracao As Integer

    Private _Unitario As Decimal
    Private _UnitarioMoeda As Decimal

    Private _ValorTotal As Decimal
    Private _ValorLiquido As Decimal

    Private _ValorTotalMoeda As Decimal
    Private _ValorLiquidoMoeda As Decimal

    Private _PesoQuantidade As String
    Private _CodigoOperacao As Integer
    Private _Operacao As Operacao
    Private _CodigoSubOperacao As Integer
    Private _SubOperacao As SubOperacao
    Private _Encargos As ListNotaFiscalXItemXEncargo

    Private _CodigoOperacaoEstado As Integer
    Private _OperacaoEstado As OperacaoXEstado
    Public CarregandoEncargos As Boolean = False

    Private _NotasDevolucao As ListNotaFiscalDevolucaoXNotaFiscal
    Private _NotasDevolucaoOriginal As ListNotaFiscalDevolucaoXNotaFiscal

    '*************************************
    '**** Saldos *************************
    '*************************************
    Private _SaldoValorOficial As Decimal
    Private _SaldoValorMoeda As Decimal
    Private _SaldoPedidoFiscal As Decimal
    Private _SaldoPedidoFisico As Decimal
    '*************************************

    Private _ObservacoesDoProduto As String
    Private _Rateado As Boolean = False
    Private _NumeroPecas As Integer

    Private _CodigoFixacao As Integer
    Private _Rateios As ListNotaFiscalxRateio
    Private _TemCentroDeCusto As Boolean

    Private _DescontosPesoDeChegada As ListNotaFiscalXDestinoXDescontos

    Private _Lotes As ListNotaFiscalXLote
    Private _ProdutoXML As String = ""
    Private _DescricaoProdutoXML As String = ""
    Private _InfAdicionalProdutoXML As String = ""
    Private _ProdutoXMLDeTerceiro As String = ""

    Private _CodigoProdutoCusto As String

    Private _NomeProdutoXML As String = ""
    Private _NCMProdutoXML As String = ""
    Private _UnidadeProdutoXML As String = ""
    Private _CentroDeCustoInformado As String
    Private _ValorFreteXML As Decimal
    Private _ValorDescontoXML As Decimal
    Private _UsarRegiao As Boolean = False

#End Region

#Region "Construtor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal NotaFiscal As NotaFiscal)
        Me.NotaFiscal = NotaFiscal
    End Sub
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

    Public Property NotaFiscal() As NotaFiscal
        Get
            Return _NotaFiscal
        End Get
        Set(ByVal value As NotaFiscal)
            _NotaFiscal = value
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

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing And Not Me.CodigoProduto Is Nothing Then _Produto = New Produto(Me.CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property
    Public Property IndiceProdutoNota() As Decimal
        Get
            Return _IndiceProdutoNota
        End Get
        Set(ByVal value As Decimal)
            _IndiceProdutoNota = value
        End Set
    End Property

    Public Property Retencao As Boolean
        Get
            If NotaFiscal.CodigoPedido = 0 OrElse NotaFiscal.NFG Then
                Return _Retencao
            Else
                If Pedido.Itens.Count = 0 Then
                    Return False
                Else
                    Return Pedido.Itens.Where(Function(s) s.CodigoProduto = Me.CodigoProduto).FirstOrDefault.Retencao
                End If
            End If
        End Get
        Set(value As Boolean)
            _Retencao = value
        End Set
    End Property

    Public Property CFOP() As Integer
        Get
            Return _CFOP
        End Get
        Set(ByVal value As Integer)
            _CFOP = value
        End Set
    End Property

    Public Property CFOPOld As Integer
        Get
            Return _CFOPOld
        End Get
        Set(value As Integer)
            _CFOPOld = value
        End Set
    End Property

    Public Property CodigoPedido() As Integer
        Get
            Return _CodigoPedido
        End Get
        Set(ByVal value As Integer)
            _CodigoPedido = value
        End Set
    End Property

    Public Property Pedido() As Pedido
        Get
            If _Pedido Is Nothing And Me.CodigoPedido > 0 And Me.NotaFiscal.CodigoEmpresa.Trim.Length > 0 Then
                _Pedido = New Negocio.Pedido(Me.NotaFiscal.CodigoEmpresa, NotaFiscal.EnderecoEmpresa, Me.CodigoPedido)
            End If

            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
        End Set
    End Property

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
        End Set
    End Property

    Public Property Lote() As String
        Get
            Return _Lote
        End Get
        Set(ByVal value As String)
            _Lote = value
            _ProdutoLoteClassificacao = Nothing
        End Set
    End Property

    Public Property Classificacao() As String
        Get
            Return _Classificacao
        End Get
        Set(ByVal value As String)
            _Classificacao = value
            _ProdutoLoteClassificacao = Nothing
        End Set
    End Property

    Public ReadOnly Property ProdutoLoteClassificacao() As LoteXClassificacao
        Get
            If _ProdutoLoteClassificacao Is Nothing And Lote.Length > 0 Then _ProdutoLoteClassificacao = New LoteXClassificacao(Me.CodigoProduto, Me.Lote, Me.Classificacao)
            Return _ProdutoLoteClassificacao
        End Get
    End Property

    Public Property CodigoDeposito() As String
        Get
            Return _CodigoDeposito
        End Get
        Set(ByVal value As String)
            _CodigoDeposito = value
        End Set
    End Property

    Public Property EnderecoDeposito() As Integer
        Get
            Return _EnderecoDeposito
        End Get
        Set(ByVal value As Integer)
            _EnderecoDeposito = value
        End Set
    End Property

    Public Property Deposito() As Cliente
        Get
            If _Deposito Is Nothing And Not Me.CodigoDeposito Is Nothing Then _Deposito = New Cliente(Me.CodigoDeposito, Me.EnderecoDeposito)
            Return _Deposito
        End Get
        Set(ByVal value As Cliente)
            _Deposito = value
        End Set
    End Property

    Public Property CodigoDepositoTerceiro() As String
        Get
            Return _CodigoDepositoTerceiro
        End Get
        Set(ByVal value As String)
            _CodigoDepositoTerceiro = value
        End Set
    End Property

    Public Property EnderecoDepositoTerceiro() As Integer
        Get
            Return _EnderecoDepositoTerceiro
        End Get
        Set(ByVal value As Integer)
            _EnderecoDepositoTerceiro = value
        End Set
    End Property

    Public Property DepositoTerceiro() As Cliente
        Get
            If _DepositoTerceiro Is Nothing And Not Me.CodigoDepositoTerceiro Is Nothing Then _DepositoTerceiro = New Cliente(Me.CodigoDepositoTerceiro, Me.EnderecoDepositoTerceiro)
            Return _DepositoTerceiro
        End Get
        Set(ByVal value As Cliente)
            _DepositoTerceiro = value
        End Set
    End Property

    Public Property PesoFiscal() As Decimal
        Get
            Return _PesoFiscal
        End Get
        Set(ByVal value As Decimal)
            _PesoFiscal = value
        End Set
    End Property

    Public Property QuantidadeFisica() As Decimal
        Get
            Return _QuantidadeFisica
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeFisica = value
        End Set
    End Property

    Public Property QuantidadeFiscal() As Decimal
        Get
            Return _QuantidadeFiscal
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeFiscal = value

            If value > 0 AndAlso Me.UnitarioMoeda > 0 Then
                Me.ValorTotalMoeda = Me.QuantidadeFiscal * Me.UnitarioMoeda
            End If

            If Not Produto Is Nothing Then

                Select Case Produto.PesoQuantidade
                    Case "P"
                        _PesoFiscal = _QuantidadeFiscal
                        If NotaFiscal.CodigoRomaneio > 0 Then
                            Me.PesoLiquido = NotaFiscal.Romaneio.PesoLiquido
                            Me.PesoBruto = NotaFiscal.Romaneio.PesoBruto
                            Me.Volumes = 1
                            Me.Numeracao = 1
                        Else
                            Me.PesoLiquido = Me.QuantidadeFiscal
                            Me.PesoBruto = Me.PesoLiquido
                            Me.Volumes = 1
                            Me.Numeracao = 1
                        End If
                    Case "Q"
                        'ACRESCENTEI PARA GRAVAR NO PESO FISCAL A QUANTIDADE EM KGS - 16/04/2021 - FURLAN

                        Dim fatorConversao As Decimal
                        Dim pesoEmbalagem As Decimal

                        If Not Me.NotaFiscal.Pedido Is Nothing Then
                            Dim undComercializao As String = String.Empty

                            For Each itemPed In Me.NotaFiscal.Pedido.Itens
                                If itemPed.CodigoProduto = Me.CodigoProduto Then
                                    undComercializao = itemPed.CodigoUnidadeComercializacao
                                    Exit For
                                End If
                            Next

                            If Me.Produto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = undComercializao).Count() > 0 Then
                                fatorConversao = Me.Produto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = undComercializao).First.FatorConversao
                                pesoEmbalagem = Me.Produto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = undComercializao).First.PesoDaEmbalagem
                            Else
                                fatorConversao = Produto.UnidadesDeComercializacao.FirstOrDefault().FatorConversao
                                pesoEmbalagem = Produto.UnidadesDeComercializacao.FirstOrDefault().PesoDaEmbalagem
                            End If

                            _PesoFiscal = _QuantidadeFiscal * fatorConversao
                            _PesoBruto = _QuantidadeFiscal * pesoEmbalagem

                        Else

                            fatorConversao = Produto.UnidadesDeComercializacao.FirstOrDefault().FatorConversao
                            pesoEmbalagem = Produto.UnidadesDeComercializacao.FirstOrDefault().PesoDaEmbalagem

                        End If

                        'ATE AQUI ------------------------------------------------------------------------


                        If CodigoEmbalagem = 0 Then
                            Me.PesoLiquido = Me.QuantidadeFiscal * fatorConversao
                            Me.PesoBruto = Me.QuantidadeFiscal * pesoEmbalagem
                            Me.Volumes = Me.QuantidadeFiscal
                            Me.Numeracao = Me.QuantidadeFiscal
                        Else
                            Me.PesoLiquido = EmbalagemProduto.PesoBruto * Me.QuantidadeDeEmbalagem
                            Me.PesoBruto = EmbalagemProduto.PesoLiquido * Me.QuantidadeDeEmbalagem
                            Me.Volumes = Me.QuantidadeDeEmbalagem
                            Me.Numeracao = Me.QuantidadeDeEmbalagem
                        End If
                End Select

            End If

        End Set
    End Property

    Public Property PesoBruto() As Decimal
        Get
            Return _PesoBruto
        End Get
        Set(ByVal value As Decimal)
            _PesoBruto = value
        End Set
    End Property

    Public Property PesoLiquido() As Decimal
        Get
            Return _PesoLiquido
        End Get
        Set(ByVal value As Decimal)
            _PesoLiquido = value
        End Set
    End Property

    Public Property Volumes() As Integer
        Get
            Return _Volumes
        End Get
        Set(ByVal value As Integer)
            _Volumes = value
        End Set
    End Property

    Public Property Numeracao() As Integer
        Get
            Return _Numeracao
        End Get
        Set(ByVal value As Integer)
            _Numeracao = value
        End Set
    End Property

    Public Property ObservacoesDoProduto() As String
        Get
            Return _ObservacoesDoProduto
        End Get
        Set(ByVal value As String)
            _ObservacoesDoProduto = value
        End Set
    End Property

    '*********************************************************
    '*********************  VALORES  *************************
    '*********************************************************
    ' OFICIAL
    Public Property Unitario() As Decimal
        Get
            Return _Unitario
        End Get
        Set(ByVal value As Decimal)
            _Unitario = value
        End Set
    End Property

    Public Property ValorTotal() As Decimal
        Get
            Return _ValorTotal
        End Get
        Set(ByVal value As Decimal)
            _ValorTotal = value
            If Not NotaFiscal.CarregandoItens Then
                If Me.QuantidadeFiscal = 0 Then
                    If Me.IndiceProdutoNota = 0 Then
                        _ValorTotalMoeda = 0
                    Else
                        _ValorTotalMoeda = value / Me.IndiceProdutoNota
                    End If
                Else
                    _ValorTotalMoeda = Me.QuantidadeFiscal * Me.UnitarioMoeda
                End If
                CarregandoEncargos = True
                Encargos = New ListNotaFiscalXItemXEncargo(Me, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                CarregandoEncargos = False
                NotaFiscal.AtualizaTotais()
            End If

        End Set
    End Property

    Public Property ValorLiquido() As Decimal
        Get
            Return _ValorLiquido
        End Get
        Set(ByVal value As Decimal)
            _ValorLiquido = IIf(value < 0, 0, value)
            If Not NotaFiscal.CarregandoItens Then
                'Apos update na base pode retirar esse if
                If Me.ValorTotalMoeda = 0 Then
                    _ValorLiquidoMoeda = 0
                Else
                    If value > 0 And Me.ValorTotalMoeda > 0 Then _ValorLiquidoMoeda = value / (Me.ValorTotal / Me.ValorTotalMoeda)
                End If
            End If
        End Set
    End Property

    ' MOEDA
    Public Property UnitarioMoeda() As Decimal
        Get
            Return _UnitarioMoeda
        End Get
        Set(ByVal value As Decimal)
            _UnitarioMoeda = value
        End Set
    End Property

    Public Property ValorTotalMoeda() As Decimal
        Get
            Return _ValorTotalMoeda
        End Get
        Set(ByVal value As Decimal)
            _ValorTotalMoeda = value
        End Set
    End Property

    Public Property ValorLiquidoMoeda() As Decimal
        Get
            Return _ValorLiquidoMoeda
        End Get
        Set(ByVal value As Decimal)
            _ValorLiquidoMoeda = value
        End Set
    End Property

    '*********************************************************
    '*********************************************************

    Public Property PesoQuantidade() As String
        Get
            Return _PesoQuantidade
        End Get
        Set(ByVal value As String)
            _PesoQuantidade = value
        End Set
    End Property

    Public Property CodigoOperacao() As Integer
        Get
            Return _CodigoOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoOperacao = value
        End Set
    End Property

    Public Property Operacao() As Operacao
        Get
            If _Operacao Is Nothing And Me.CodigoOperacao > 0 Then _Operacao = New Operacao(Me.CodigoOperacao)
            Return _Operacao
        End Get
        Set(ByVal value As Operacao)
            _Operacao = value
        End Set
    End Property

    Public Property CodigoSubOperacao() As Integer
        Get
            Return _CodigoSubOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSubOperacao = value
            _SubOperacao = Nothing
            'CarregandoEncargos = True
            'If Not NotaFiscal.CarregandoItens Then
            '    _Encargos = New ListNotaFiscalXItemXEncargo(Me, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
            'End If
            'CarregandoEncargos = False
        End Set
    End Property

    Public Property SubOperacao() As SubOperacao
        Get
            If _SubOperacao Is Nothing And Me.CodigoOperacao > 0 And Me.CodigoSubOperacao > 0 Then _SubOperacao = New SubOperacao(Me.CodigoOperacao, Me.CodigoSubOperacao)
            Return _SubOperacao
        End Get
        Set(ByVal value As SubOperacao)
            _SubOperacao = value
        End Set
    End Property

    '********************************************************************************
    '*************** SALDOS *********************************************************
    '********************************************************************************
    Public Property SaldoValorOficial() As Decimal
        Get
            Return _SaldoValorOficial
        End Get
        Set(ByVal value As Decimal)
            _SaldoValorOficial = value
        End Set
    End Property

    Public Property SaldoValorMoeda() As Decimal
        Get
            Return _SaldoValorMoeda
        End Get
        Set(ByVal value As Decimal)
            _SaldoValorMoeda = value
        End Set
    End Property

    Public Property SaldoPedidoFiscal() As Decimal
        Get
            Return _SaldoPedidoFiscal
        End Get
        Set(ByVal value As Decimal)
            _SaldoPedidoFiscal = value
        End Set
    End Property

    Public Property SaldoPedidoFisico() As Decimal
        Get
            Return _SaldoPedidoFisico
        End Get
        Set(ByVal value As Decimal)
            _SaldoPedidoFisico = value
        End Set
    End Property

    '********************************************************************************
    '********************************************************************************

    Public Property CodigoEmbalagemIndea() As String
        Get
            Return _CodigoEmbalagemIndea
        End Get
        Set(ByVal value As String)
            _CodigoEmbalagemIndea = value
        End Set
    End Property

    Public Property CodigoEmbalagem() As Integer
        Get
            Return _codigoEmbalagem
        End Get
        Set(ByVal value As Integer)
            _codigoEmbalagem = value
        End Set
    End Property

    Public ReadOnly Property Embalagem() As Embalagem
        Get
            If _Embalagem Is Nothing AndAlso _codigoEmbalagem > 0 Then _Embalagem = New Embalagem(_codigoEmbalagem)
            Return _Embalagem
        End Get
    End Property

    Public Property CodigoTipoDeEmbalagem() As String
        Get
            Return _CodigoTipoDeEmbalagem
        End Get
        Set(ByVal value As String)
            _CodigoTipoDeEmbalagem = value
        End Set
    End Property

    Public Property TipoDeEmbalagem() As TipoDeEmbalagem
        Get
            If _TipoDeEmbalagem Is Nothing AndAlso _CodigoTipoDeEmbalagem.Length > 0 Then _TipoDeEmbalagem = New TipoDeEmbalagem(_CodigoTipoDeEmbalagem)
            Return _TipoDeEmbalagem
        End Get
        Set(ByVal value As TipoDeEmbalagem)
            _TipoDeEmbalagem = value
        End Set
    End Property

    Public Property CapacidadeEmbalagem() As Decimal
        Get
            Return _CapacidadeEmbalagem
        End Get
        Set(ByVal value As Decimal)
            _CapacidadeEmbalagem = value
        End Set
    End Property

    Public Property QuantidadeDeEmbalagem() As Integer
        Get
            Return _QuantidadeDeEmbalagem
        End Get
        Set(ByVal value As Integer)
            _QuantidadeDeEmbalagem = value
        End Set
    End Property

    Public Property EmbalagemProduto() As ProdutoXEmbalagem
        Get
            If _EmbalagemProduto Is Nothing And _CodigoEmbalagemIndea.Length > 0 Then _EmbalagemProduto = New ProdutoXEmbalagem(_CodigoProduto, CodigoEmbalagem, _CodigoTipoDeEmbalagem, _CapacidadeEmbalagem)
            Return _EmbalagemProduto
        End Get
        Set(ByVal value As ProdutoXEmbalagem)
            _EmbalagemProduto = value
        End Set
    End Property

    Public Property NumeroPecas() As Integer
        Get
            Return _NumeroPecas
        End Get
        Set(ByVal value As Integer)
            _NumeroPecas = value
        End Set
    End Property

    Public Property CodigoFixacao() As Integer
        Get
            Return _CodigoFixacao
        End Get
        Set(ByVal value As Integer)
            _CodigoFixacao = value
        End Set
    End Property

    Public Property TemCentroDeCusto() As Boolean
        Get
            Return _TemCentroDeCusto
        End Get
        Set(ByVal value As Boolean)
            _TemCentroDeCusto = value
        End Set
    End Property

    Public Property Rateado() As Boolean
        Get
            Return _Rateado
        End Get
        Set(ByVal value As Boolean)
            _Rateado = value
        End Set
    End Property

    Public Property CodigoOperacaoEstado As Integer
        Get
            Return _CodigoOperacaoEstado
        End Get
        Set(value As Integer)
            _CodigoOperacaoEstado = value
            _OperacaoEstado = Nothing
        End Set
    End Property

    Public Property OperacaoEstado As OperacaoXEstado
        Get
            If Me.CodigoOperacaoEstado > 0 Then
                Dim Parametros As New OperacaoXEstado
                Parametros.Codigo = Me.CodigoOperacaoEstado
                _OperacaoEstado = New OperacaoXEstado(Parametros)
            End If
            Return _OperacaoEstado
        End Get
        Set(value As OperacaoXEstado)
            _OperacaoEstado = value
        End Set
    End Property

    '********************************************************************************
    '*************** LISTAS *********************************************************
    '********************************************************************************
    Public Property Encargos() As ListNotaFiscalXItemXEncargo
        Get
            If _Encargos Is Nothing Then
                CarregandoEncargos = True
                _Encargos = New ListNotaFiscalXItemXEncargo(Me, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                CarregandoEncargos = False
            End If

            Return _Encargos
        End Get
        Set(ByVal value As ListNotaFiscalXItemXEncargo)
            _Encargos = value
        End Set
    End Property

    Public Property NotasDevolucao() As ListNotaFiscalDevolucaoXNotaFiscal
        Get
            If _NotasDevolucao Is Nothing OrElse _NotasDevolucao.Count = 0 Then
                If Not NotaFiscal.SubOperacao.Devolucao Then
                    _NotasDevolucao = New ListNotaFiscalDevolucaoXNotaFiscal(Me)
                    If NotaFiscal.IUD = "U" Or NotaFiscal.IUD = "D" Then _NotasDevolucao.CarregarNotasUsadasNaDevolucao()
                    _NotasDevolucaoOriginal = New ListNotaFiscalDevolucaoXNotaFiscal(Me)
                Else
                    If NotaFiscal.IUD = "I" Then
                        _NotasDevolucao = New ListNotaFiscalDevolucaoXNotaFiscal(Me)
                        _NotasDevolucao.CarregarNotasParaSelecao(-1)
                    ElseIf NotaFiscal.IUD = "U" Or NotaFiscal.IUD = "D" Then
                        _NotasDevolucao = New ListNotaFiscalDevolucaoXNotaFiscal(Me)
                        _NotasDevolucao.CarregarNotasUsadasNaDevolucao()
                    End If
                End If
            End If
            Return _NotasDevolucao
        End Get
        Set(ByVal value As ListNotaFiscalDevolucaoXNotaFiscal)
            _NotasDevolucao = value
        End Set
    End Property

    'Usado apenas para Guardar a Lista das Devolu��es Carregadas - Furlan - 24-03-2023.
    Public Property NotasDevolucaoOriginal() As ListNotaFiscalDevolucaoXNotaFiscal
        Get
            If _NotasDevolucaoOriginal Is Nothing Then
                If Not NotaFiscal.SubOperacao.Devolucao Then
                    _NotasDevolucaoOriginal = New ListNotaFiscalDevolucaoXNotaFiscal(Me)
                Else
                    If NotaFiscal.IUD = "I" Then
                        _NotasDevolucaoOriginal = New ListNotaFiscalDevolucaoXNotaFiscal(Me)
                        _NotasDevolucaoOriginal.CarregarNotasParaSelecao(-1)
                    ElseIf NotaFiscal.IUD = "U" Or NotaFiscal.IUD = "D" Then
                        _NotasDevolucaoOriginal = New ListNotaFiscalDevolucaoXNotaFiscal(Me)
                        _NotasDevolucaoOriginal.CarregarNotasUsadasNaDevolucao()
                    End If
                End If
            End If
            Return _NotasDevolucaoOriginal
        End Get
        Set(ByVal value As ListNotaFiscalDevolucaoXNotaFiscal)
            _NotasDevolucaoOriginal = value
        End Set
    End Property

    Public Property Rateios() As ListNotaFiscalxRateio
        Get
            If _Rateios Is Nothing Then _Rateios = New ListNotaFiscalxRateio(Me)
            Return _Rateios
        End Get
        Set(ByVal value As ListNotaFiscalxRateio)
            _Rateios = value
        End Set
    End Property

    Public Property DescontosPesoDeChegada() As ListNotaFiscalXDestinoXDescontos
        Get
            If _DescontosPesoDeChegada Is Nothing Then _DescontosPesoDeChegada = New ListNotaFiscalXDestinoXDescontos(Me)
            Return _DescontosPesoDeChegada
        End Get
        Set(ByVal value As ListNotaFiscalXDestinoXDescontos)
            _DescontosPesoDeChegada = value
        End Set
    End Property

    Public Property Lotes() As ListNotaFiscalXLote
        Get
            If _Lotes Is Nothing Then _Lotes = New ListNotaFiscalXLote(Me)
            Return _Lotes
        End Get
        Set(ByVal value As ListNotaFiscalXLote)
            _Lotes = value
        End Set
    End Property

    Public Property ProdutoXML() As String
        Get
            Return _ProdutoXML
        End Get
        Set(ByVal value As String)
            _ProdutoXML = value
        End Set
    End Property

    Public Property DescricaoProdutoXML() As String
        Get
            Return _DescricaoProdutoXML
        End Get
        Set(ByVal value As String)
            _DescricaoProdutoXML = value
        End Set
    End Property

    Public Property InfAdicionalProdutoXML() As String
        Get
            Return _InfAdicionalProdutoXML
        End Get
        Set(ByVal value As String)
            _InfAdicionalProdutoXML = value
        End Set
    End Property

    Public Property ProdutoXMLDeTerceiro() As String
        Get
            If _ProdutoXML.Length > 0 Then _ProdutoXMLDeTerceiro = ProdutoXML + "-" + _DescricaoProdutoXML
            Return _ProdutoXMLDeTerceiro
        End Get
        Set(ByVal value As String)
            _ProdutoXMLDeTerceiro = value
        End Set
    End Property

    Public Property CodigoProdutoCusto() As String
        Get
            Return _CodigoProdutoCusto
        End Get
        Set(ByVal value As String)
            _CodigoProdutoCusto = value
        End Set
    End Property

    Public Property NomeProdutoXML As String
        Get
            Return _NomeProdutoXML
        End Get
        Set(value As String)
            _NomeProdutoXML = value
        End Set
    End Property

    Public Property NCMProdutoXML As String
        Get
            Return _NCMProdutoXML
        End Get
        Set(value As String)
            _NCMProdutoXML = value
        End Set
    End Property

    Public Property UnidadeProdutoXML As String
        Get
            Return _UnidadeProdutoXML
        End Get
        Set(value As String)
            _UnidadeProdutoXML = value
        End Set
    End Property

    Public Property CentroDeCustoInformado As String
        Get
            Return _CentroDeCustoInformado
        End Get
        Set(value As String)
            _CentroDeCustoInformado = value
        End Set
    End Property

    Public Property ValorFreteXML As Decimal
        Get
            Return _ValorFreteXML
        End Get
        Set(value As Decimal)
            _ValorFreteXML = value
        End Set
    End Property

    Public Property ValorDescontoXML As Decimal
        Get
            Return _ValorDescontoXML
        End Get
        Set(value As Decimal)
            _ValorDescontoXML = value
        End Set
    End Property

    Public Property CodigoProdutoOld As String
        Get
            Return _CodigoProdutoOld
        End Get
        Set(value As String)
            _CodigoProdutoOld = value
        End Set
    End Property

    Public Property UsarRegiao As Boolean
        Get
            Return _UsarRegiao
        End Get
        Set(value As Boolean)
            _UsarRegiao = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Sub SalvarSql(ByRef Sqls As ArrayList, Optional pSequencia As Integer = 1)
        Dim sql As String
        Select Case IUD
            Case "I"
                Me.Sequencia = pSequencia
                sql = "INSERT INTO NotasFiscaisXItens" & vbCrLf &
                      " (Empresa_Id, EndEmpresa_Id, " & vbCrLf &
                      "  Cliente_Id, EndCliente_Id, " & vbCrLf &
                      "  EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf &
                      "  Produto_Id, CFOP_Id,Sequencia_id, Lote, Classificacao," & vbCrLf &
                      "  Embalagem, TipoDeEmbalagem, CapacidadeEmbalagem, QtdeDeEmbalagem, Pedido, " & vbCrLf &
                      "  Deposito, EndDeposito, " & vbCrLf &
                      "  DepositoTerceiro, EndDepositoTerceiro, " & vbCrLf &
                      "  PesoFiscal, QuantidadeFisica, QuantidadeFiscal, Unitario," & vbCrLf &
                      "  Valor, ValorLiquido, ValorMoeda, ValorLiquidoMoeda, " & vbCrLf &
                      "  PesoQuantidade,Operacao, SubOperacao,ObservacoesDoProduto,Rateado,NumeroPecas,Fixacao, OperacaoXEstado," & vbCrLf &
                       " ProdutoXML, DescricaoProdutoXML,InfAdicionalProdutoXML, ProdutoParaCusto)" & vbCrLf &
                      "VALUES" & vbCrLf &
                      " ('" & NotaFiscal.CodigoEmpresa & "'," & NotaFiscal.EnderecoEmpresa & ", " & vbCrLf &
                      "'" & NotaFiscal.CodigoCliente & "', " & NotaFiscal.EnderecoCliente & ", " & vbCrLf &
                      "'" & NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & NotaFiscal.Serie & "', " & NotaFiscal.Codigo & ", " & vbCrLf &
                      "'" & Me.CodigoProduto & "', " & Me.CFOP & "," & Me.Sequencia & ",'" & Me.Lote & "','" & Me.Classificacao & "'," & vbCrLf &
                      Me.CodigoEmbalagem & ",'" & Me.CodigoTipoDeEmbalagem & "'," & Str(Me.CapacidadeEmbalagem) & "," & Me.QuantidadeDeEmbalagem & "," & vbCrLf &
                      NotaFiscal.CodigoPedido.ToSqlNULL & "," & vbCrLf &
                      "'" & NotaFiscal.CodigoDeposito & "'," & NotaFiscal.EnderecoDeposito & ", " & vbCrLf &
                      "'',0," & vbCrLf &
                      Str(Me.PesoFiscal) & ", " & Str(Me.QuantidadeFisica) & ", " & Str(Me.QuantidadeFiscal) & "," & Str(Me.Unitario) & "," & vbCrLf &
                      Str(Me.ValorTotal) & "," & Str(Me.ValorLiquido) & "," & Str(Me.ValorTotalMoeda) & "," & Str(Me._ValorLiquidoMoeda) & "," & vbCrLf &
                      "'" & Me.PesoQuantidade & "'," & Me.CodigoOperacao & ", " & Me.CodigoSubOperacao & ",'" & Me.ObservacoesDoProduto & IIf(String.IsNullOrWhiteSpace(Me.Encargos.MensagemImpostos), "", " | " & Me.Encargos.MensagemImpostos) & vbCrLf & "'," & IIf(Me.Rateado, "1", "0") & "," & Me.NumeroPecas & ", " & Me.CodigoFixacao & "," & Me.CodigoOperacaoEstado & "," & vbCrLf &
                      "'" & Me.ProdutoXML & "','" & Me.DescricaoProdutoXML & "','" & Me.InfAdicionalProdutoXML & "', '" & Me.CodigoProdutoCusto & "')" & vbCrLf


                Sqls.Add(sql)

                If Me.DescontosPesoDeChegada.Count() > 0 Then
                    Me.DescontosPesoDeChegada.SalvarSql(Sqls)
                End If

                If Me.Encargos IsNot Nothing AndAlso Me.Encargos.Count > 0 Then
                    Me.Encargos.SalvarSql(Sqls)
                End If

                If Me.NotasDevolucao IsNot Nothing AndAlso Me.NotasDevolucao.Count > 0 Then
                    Me.NotasDevolucao.SalvarSql(Sqls)
                End If

                If Me.Rateios IsNot Nothing AndAlso Me.Rateios.Count > 0 Then
                    Me.Rateios.SalvarSql(Sqls)
                End If

                If Me.Lotes IsNot Nothing AndAlso Me.Lotes.Count > 0 Then
                    Me.Lotes.SalvarSql(Sqls)
                End If

                If Not DescontosPesoDeChegada Is Nothing And DescontosPesoDeChegada.Count > 0 Then DescontosPesoDeChegada.SalvarSql(Sqls)

                If Me.Produto.Almoxarifado OrElse Me.Produto.PrecoDoProduto Then
                    sql = "if NOT Exists(SELECT Data_id FROM PRODUTOSxPRECOS WHERE TABELA_ID = 1 AND CLIENTE_ID = '" & NotaFiscal.CodigoCliente & "' AND ENDCLIENTE_ID = " & NotaFiscal.EnderecoCliente & " AND PRODUTO_ID = '" & Me.CodigoProduto & "' AND DATA_ID = '" & NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "')" & vbCrLf &
                               "begin" & vbCrLf &
                               "   INSERT INTO PRODUTOSxPRECOS (TABELA_ID, CLIENTE_ID, ENDCLIENTE_ID, PRODUTO_ID, DATA_ID, MOEDA, VALOR) VALUES (1,'" & NotaFiscal.CodigoCliente & "'," & NotaFiscal.EnderecoCliente & ",'" & Me.CodigoProduto & "','" & NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "', 1," & Str(Me.Unitario) & ")" & vbCrLf &
                               "end" & vbCrLf &
                               "else" & vbCrLf &
                               "begin" & vbCrLf &
                               "   UPDATE PRODUTOSxPRECOS SET VALOR = " & Str(Me.Unitario) & " WHERE TABELA_ID = 1 AND CLIENTE_ID = '" & NotaFiscal.CodigoCliente & "' AND ENDCLIENTE_ID = " & NotaFiscal.EnderecoCliente & " AND PRODUTO_ID = '" & Me.CodigoProduto & "' AND DATA_ID = '" & NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                               "end" & vbCrLf

                    Sqls.Add(sql)
                End If

            Case "U"

                IUD = "D"

                If Me.Rateios IsNot Nothing AndAlso Me.Rateios.Count > 0 Then
                    Me.Rateios.SalvarSql(Sqls)
                End If

                Encargos.SalvarSql(Sqls)

                IUD = "U"

                If Me.CodigoProdutoOld Is Nothing OrElse Me.CodigoProdutoOld.Length = 0 Then

                    sql = "Update NotasFiscaisXItens set " & vbCrLf &
                      "   Lote                 ='" & Me.Lote & "'" & vbCrLf &
                      "  ,Classificacao        ='" & Me.Classificacao & "'" & vbCrLf &
                      "  ,Embalagem            = " & Me.CodigoEmbalagem & vbCrLf &
                      "  ,TipoDeEmbalagem      ='" & Me.CodigoTipoDeEmbalagem & "'" & vbCrLf &
                      "  ,CapacidadeEmbalagem  = " & Str(Me.CapacidadeEmbalagem) & vbCrLf &
                      "  ,QtdeDeEmbalagem      = " & Me.QuantidadeDeEmbalagem & vbCrLf &
                      "  ,Pedido               =" & NotaFiscal.CodigoPedido.ToSqlNULL & vbCrLf &
                      "  ,Deposito             ='" & NotaFiscal.CodigoDeposito & "'" & vbCrLf &
                      "  ,EndDeposito          = " & NotaFiscal.EnderecoDeposito & vbCrLf &
                      "  ,DepositoTerceiro     =''" & vbCrLf &
                      "  ,EndDepositoTerceiro  = 0" & vbCrLf &
                      "  ,PesoFiscal           = " & Str(Me.PesoFiscal) & vbCrLf &
                      "  ,QuantidadeFisica     = " & Str(Me.QuantidadeFisica) & vbCrLf &
                      "  ,QuantidadeFiscal     = " & Str(Me.QuantidadeFiscal) & vbCrLf &
                      "  ,Unitario             = " & Str(Me.Unitario) & vbCrLf &
                      "  ,Valor                = " & Str(Me.ValorTotal) & vbCrLf &
                      "  ,ValorLiquido         = " & Str(Me.ValorLiquido) & vbCrLf &
                      "  ,ValorMoeda           = " & Str(Me.ValorTotalMoeda) & vbCrLf &
                      "  ,ValorLiquidoMoeda    = " & Str(Me.ValorLiquidoMoeda) & vbCrLf &
                      "  ,PesoQuantidade       ='" & Me.PesoQuantidade & "'" & vbCrLf &
                      "  ,Operacao             = " & Me.CodigoOperacao & vbCrLf &
                      "  ,SubOperacao          = " & Me.CodigoSubOperacao & vbCrLf &
                      "  ,ObservacoesDoProduto ='" & Me.ObservacoesDoProduto & IIf(String.IsNullOrWhiteSpace(Me.Encargos.MensagemImpostos), "", " | " & Me.Encargos.MensagemImpostos) & "'" & vbCrLf &
                      "  ,Rateado              = " & IIf(Me.Rateado, "1", "0") & vbCrLf &
                      "  ,NumeroPecas          = " & Me.NumeroPecas & vbCrLf &
                      "  ,Fixacao              = " & Me.CodigoFixacao & vbCrLf &
                      "  ,CFOP_Id              = " & Me.CFOP & vbCrLf &
                      "  ,OperacaoxEstado      = " & Me.CodigoOperacaoEstado & vbCrLf &
                      "  ,ProdutoXML           = '" & Me.ProdutoXML & "'" & vbCrLf &
                      "  ,DescricaoProdutoXML  = '" & Me.DescricaoProdutoXML & "'" & vbCrLf &
                      "  ,InfAdicionalProdutoXML = '" & Me.InfAdicionalProdutoXML & "'" & vbCrLf &
                      "  ,ProdutoParaCusto     ='" & Me.CodigoProdutoCusto & "'" & vbCrLf &
                      " Where Empresa_Id      ='" & NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id   = " & NotaFiscal.EnderecoEmpresa & vbCrLf &
                      "   and Cliente_Id      ='" & NotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id   = " & NotaFiscal.EnderecoCliente & vbCrLf &
                      "   and EntradaSaida_Id ='" & NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and Serie_Id        ='" & NotaFiscal.Serie & "'" & vbCrLf &
                      "   and Nota_Id         = " & NotaFiscal.Codigo & vbCrLf &
                      "   and Produto_Id      ='" & Me.CodigoProduto & "'" & vbCrLf &
                      "   and Sequencia_id    = " & Me.Sequencia
                    '"   and CFOP_Id         = " & _CFOP & vbCrLf & _
                    Sqls.Add(sql)

                Else

                    sql = "Update NotasFiscaisXItens set " & vbCrLf &
                      "   Produto_Id            = '" & Me.CodigoProduto & "'" & vbCrLf &
                      "   ,Lote                 ='" & Me.Lote & "'" & vbCrLf &
                      "  ,Classificacao         ='" & Me.Classificacao & "'" & vbCrLf &
                      "  ,Embalagem             = " & Me.CodigoEmbalagem & vbCrLf &
                      "  ,TipoDeEmbalagem       ='" & Me.CodigoTipoDeEmbalagem & "'" & vbCrLf &
                      "  ,CapacidadeEmbalagem   = " & Str(Me.CapacidadeEmbalagem) & vbCrLf &
                      "  ,QtdeDeEmbalagem       = " & Me.QuantidadeDeEmbalagem & vbCrLf &
                      "  ,Pedido                =" & NotaFiscal.CodigoPedido.ToSqlNULL & vbCrLf &
                      "  ,Deposito              ='" & NotaFiscal.CodigoDeposito & "'" & vbCrLf &
                      "  ,EndDeposito           = " & NotaFiscal.EnderecoDeposito & vbCrLf &
                      "  ,DepositoTerceiro      =''" & vbCrLf &
                      "  ,EndDepositoTerceiro   = 0" & vbCrLf &
                      "  ,PesoFiscal            = " & Str(Me.PesoFiscal) & vbCrLf &
                      "  ,QuantidadeFisica      = " & Str(Me.QuantidadeFisica) & vbCrLf &
                      "  ,QuantidadeFiscal      = " & Str(Me.QuantidadeFiscal) & vbCrLf &
                      "  ,Unitario              = " & Str(Me.Unitario) & vbCrLf &
                      "  ,Valor                 = " & Str(Me.ValorTotal) & vbCrLf &
                      "  ,ValorLiquido          = " & Str(Me.ValorLiquido) & vbCrLf &
                      "  ,ValorMoeda            = " & Str(Me.ValorTotalMoeda) & vbCrLf &
                      "  ,ValorLiquidoMoeda     = " & Str(Me.ValorLiquidoMoeda) & vbCrLf &
                      "  ,PesoQuantidade        ='" & Me.PesoQuantidade & "'" & vbCrLf &
                      "  ,Operacao              = " & Me.CodigoOperacao & vbCrLf &
                      "  ,SubOperacao           = " & Me.CodigoSubOperacao & vbCrLf &
                      "  ,ObservacoesDoProduto  ='" & Me.ObservacoesDoProduto & IIf(String.IsNullOrWhiteSpace(Me.Encargos.MensagemImpostos), "", " | " & Me.Encargos.MensagemImpostos) & "'" & vbCrLf &
                      "  ,Rateado               = " & IIf(Me.Rateado, "1", "0") & vbCrLf &
                      "  ,NumeroPecas           = " & Me.NumeroPecas & vbCrLf &
                      "  ,Fixacao               = " & Me.CodigoFixacao & vbCrLf &
                      "  ,CFOP_Id               = " & Me.CFOP & vbCrLf &
                      "  ,OperacaoxEstado       = " & Me.CodigoOperacaoEstado & vbCrLf &
                      "  ,ProdutoXML            = '" & Me.ProdutoXML & "'" & vbCrLf &
                      "  ,DescricaoProdutoXML   = '" & Me.DescricaoProdutoXML & "'" & vbCrLf &
                      "  ,InfAdicionalProdutoXML = '" & Me.InfAdicionalProdutoXML & "'" & vbCrLf &
                      "  ,ProdutoParaCusto      ='" & Me.CodigoProdutoCusto & "'" & vbCrLf &
                      " Where Empresa_Id        ='" & NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id     = " & NotaFiscal.EnderecoEmpresa & vbCrLf &
                      "   and Cliente_Id        ='" & NotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id     = " & NotaFiscal.EnderecoCliente & vbCrLf &
                      "   and EntradaSaida_Id   ='" & NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and Serie_Id          ='" & NotaFiscal.Serie & "'" & vbCrLf &
                      "   and Nota_Id           = " & NotaFiscal.Codigo & vbCrLf &
                      "   and Produto_Id        ='" & Me.CodigoProdutoOld & "'" & vbCrLf &
                      "   and Sequencia_id      = " & Me.Sequencia
                    '"   and CFOP_Id         = " & _CFOP & vbCrLf & _
                    Sqls.Add(sql)

                End If

                Encargos.SalvarSql(Sqls)

                If Me.Rateios IsNot Nothing AndAlso Me.Rateios.Count > 0 Then
                    IUD = "I"
                    Me.Rateios.SalvarSql(Sqls)
                    IUD = "U"
                End If

                If Me.Lotes IsNot Nothing AndAlso Me.Lotes.Count > 0 Then
                    Me.Lotes.SalvarSql(Sqls)
                End If

                If Not DescontosPesoDeChegada Is Nothing And DescontosPesoDeChegada.Count > 0 Then DescontosPesoDeChegada.SalvarSql(Sqls)

                If Me.Produto.Almoxarifado OrElse Me.Produto.PrecoDoProduto Then
                    sql = "if NOT Exists(SELECT Data_id FROM PRODUTOSxPRECOS WHERE PRODUTO_ID = '" & Me.CodigoProduto & "' AND DATA_ID = '" & NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "' AND MOEDA = 1)" & vbCrLf &
                                       "begin" & vbCrLf &
                                       "   INSERT INTO PRODUTOSxPRECOS (TABELA_ID, CLIENTE_ID, ENDCLIENTE_ID, PRODUTO_ID, DATA_ID, MOEDA, VALOR) VALUES (1, '" & NotaFiscal.Cliente.Codigo & "', " & NotaFiscal.Cliente.CodigoEndereco & ", '" & Me.CodigoProduto & "','" & NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "', 1, " & Str(Me.Unitario) & ")" & vbCrLf &
                                       "end" & vbCrLf &
                                       "else" & vbCrLf &
                                       "begin" & vbCrLf &
                                       "   UPDATE PRODUTOSxPRECOS SET VALOR = " & Str(Me.Unitario) & " WHERE PRODUTO_ID = '" & Me.CodigoProduto & "' AND DATA_ID = '" & NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "' AND MOEDA = 1" & vbCrLf &
                                       "end" & vbCrLf

                    Sqls.Add(sql)
                End If

            Case "D"

                If Not DescontosPesoDeChegada Is Nothing And DescontosPesoDeChegada.Count > 0 Then DescontosPesoDeChegada.SalvarSql(Sqls)

                If Me.Rateios IsNot Nothing AndAlso Me.Rateios.Count > 0 Then
                    Me.Rateios.SalvarSql(Sqls)
                End If

                If Me.Lotes IsNot Nothing AndAlso Me.Lotes.Count > 0 Then
                    Me.Lotes.SalvarSql(Sqls)
                End If

                If Me.NotasDevolucao IsNot Nothing AndAlso Me.NotasDevolucao.Count > 0 Then
                    Me.NotasDevolucao.SalvarSql(Sqls)
                End If

                If Me.Encargos IsNot Nothing AndAlso Me.Encargos.Count > 0 Then
                    Me.Encargos.SalvarSql(Sqls)
                End If

                If Me.CodigoProdutoOld Is Nothing OrElse Me.CodigoProdutoOld.Length = 0 Then

                    Me.CodigoProdutoOld = Me.CodigoProduto

                End If

                If Me.CFOPOld = 0 Then

                    Me.CFOPOld = Me.CFOP

                End If

                sql = " Delete NotasFiscaisXItens " & vbCrLf &
                      " Where  Empresa_Id     ='" & NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id   = " & NotaFiscal.EnderecoEmpresa & vbCrLf &
                      "   and Cliente_Id      ='" & NotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id   = " & NotaFiscal.EnderecoCliente & vbCrLf &
                      "   and EntradaSaida_Id ='" & NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and Nota_Id         = " & NotaFiscal.Codigo & vbCrLf &
                      "   and Serie_Id        ='" & NotaFiscal.Serie & "'" & vbCrLf &
                      "   and Produto_Id      ='" & Me.CodigoProdutoOld & "'" & vbCrLf &
                      "   and Sequencia_id    = " & Me.Sequencia
                Sqls.Add(sql)
            Case ""
                If Me.Rateios IsNot Nothing AndAlso Me.Rateios.Count > 0 Then
                    Me.Rateios.SalvarSql(Sqls)
                End If
                '    If Me.NotasDevolucao IsNot Nothing AndAlso Me.NotasDevolucao.Count > 0 Then
                '        Me.NotasDevolucao.SalvarSql(Sqls)
                '    End If
                '    If Me.Encargos IsNot Nothing AndAlso Me.Encargos.Count > 0 Then
                '        Me.Encargos.SalvarSql(Sqls)
                '    End If
        End Select
    End Sub

    Public Function Clone() As NotaFiscalXItem
        Dim item As NotaFiscalXItem = CType(Me.MemberwiseClone(), NotaFiscalXItem)
        Return item
    End Function

#End Region

End Class