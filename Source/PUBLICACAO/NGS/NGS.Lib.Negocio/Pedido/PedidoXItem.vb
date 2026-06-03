Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

'******************************************************************************************************************************************************
'********************************************    LISTA DE ITENS DO PEDIDO   *************************************************************************
'******************************************************************************************************************************************************
Public Class ListPedidoXItem
    Inherits List(Of PedidoXItem)

#Region "Contrutor"
    Public Sub New(ByVal pPedido As Pedido)
        _Pedido = pPedido
        If _Pedido.Codigo = 0 Then Exit Sub

        Dim Banco As New AcessaBanco()
        Try
            Dim strSQL As String = ""
            strSQL = "SELECT PxI.Produto_Id," & vbCrLf & _
                     "       P.Nome," & vbCrLf & _
                     "       P.Descricao," & vbCrLf & _
                     "       isnull(P.RegMinAgr,'') AS RegMinAgr," & vbCrLf & _
                     "       PxI.Classificacao," & vbCrLf & _
                     "       isnull(PxI.OperacaoXEstado,0) as OperacaoXEstado," & vbCrLf & _
                     "       isnull(PxI.Retencao,1) as Retencao," & vbCrLf & _
                     "       isnull(pxi.UnidadeComercializacao,p.unidade)   as UnidadeComercializacao" & vbCrLf & _
                     "  FROM PedidoXItem PxI" & vbCrLf & _
                     " INNER JOIN Produtos P" & vbCrLf & _
                     "    on PxI.Produto_id = P.Produto_id" & vbCrLf & _
                     " WHERE PxI.Empresa_Id       ='" & _Pedido.CodigoEmpresa & "'" & vbCrLf & _
                     "   AND PxI.EndEmpresa_Id    = " & _Pedido.EnderecoEmpresa & vbCrLf & _
                     "   AND PxI.Pedido_Id        = " & _Pedido.Codigo & vbCrLf & _
                     " Order by P.Nome" & vbCrLf

            Dim ds As DataSet
            ds = Banco.ConsultaDataSet(strSQL, "PXI")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim objItem As New PedidoXItem(_Pedido)
                objItem.CodigoProduto = row("Produto_Id")

                If _Pedido.Empresa.Empresa.UsarRegistroMinAgr Then
                    objItem.Descricao = row("Nome") & "-" & row("Descricao") & "(" & row("RegMinAgr")
                ElseIf _Pedido.Empresa.Empresa.UsarDescricaoProduto Then
                    objItem.Descricao = row("Nome") & "-" & row("Descricao")
                Else
                    objItem.Descricao = row("Nome")
                End If

                objItem.CodigoClassificacao = row("Classificacao")
                objItem.CodigoOperacaoXEstado = row("OperacaoXEstado")
                objItem.Retencao = row("Retencao")
                objItem.CodigoUnidadeComercializacao = row("UnidadeComercializacao")

                Me.Add(objItem)
            Next
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    'Public _IUD As String = ""
    Public _Pedido As Pedido
#End Region

#Region "Property"
    'Public Property IUD As String
    '    Get
    '        Return _IUD
    '    End Get
    '    Set(value As String)
    '        _IUD = value
    '    End Set
    'End Property

    Public ReadOnly Property Pedido As Pedido
        Get
            Return _Pedido
        End Get
    End Property

    Public ReadOnly Property QuantidadeTotal As Decimal
        Get
            Return Me.Sum(Function(s) s.Lancamentos.QuantidadeTotalPrdFaturamento)
        End Get
    End Property

    Public ReadOnly Property TotalOficial As Decimal
        Get
            Return Me.Sum(Function(s) s.Lancamentos.TotalOficialPrd)
        End Get
    End Property

    Public ReadOnly Property TotalMoeda As Decimal
        Get
            Return Me.Sum(Function(s) s.Lancamentos.TotalMoedaPrd)
        End Get
    End Property

    Public ReadOnly Property Total As Decimal
        Get
            If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Return TotalOficial
            Else
                Return TotalMoeda
            End If
        End Get
    End Property

    'Segue a moeda do pedido
    Public ReadOnly Property Liquido As Decimal
        Get
            If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Return LiquidoOficial
            Else
                Return LiquidoMoeda
            End If
        End Get
    End Property

    Public ReadOnly Property LiquidoOficial As Decimal
        Get
            Return Me.SelectMany(Function(s) s.Encargos).Where(Function(s) s.CodigoEncargo = "LIQUIDO").Sum(Function(k) k.ValorOficial)
        End Get
    End Property

    Public ReadOnly Property LiquidoMoeda As Decimal
        Get
            Return Me.SelectMany(Function(s) s.Encargos).Where(Function(s) s.CodigoEncargo = "LIQUIDO").Sum(Function(k) k.ValorMoeda)
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Item As PedidoXItem In Me
            If Pedido.IUD = "I" Or Pedido.IUD = "D" Then Item.IUD = Pedido.IUD
            Item.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

'******************************************************************************************************************************************************
'********************************************    CLASSE BASE ITEM DO PEDIDO   *************************************************************************
'******************************************************************************************************************************************************
Public Class PedidoXItem
#Region "Contrutor"
    Public Sub New(ByVal pPedido As Pedido)
        Me.Pedido = pPedido
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Pedido As Pedido
    Private _CodigoProduto As String
    Private _Produto As Produto
    Private _Descricao As String = ""
    Private _CodigoClassificacao As Integer
    Private _Classificacao As Classificacao
    Private _Retencao As Boolean = False
    Private _TemNota As Boolean

    Private _UnidadeComercializacao As String = ""

    Private _CodigoUnidadeComercializacao As String
    Private _UnidadeComercializacaoFatorDeConversao As Decimal

    '*********** Quantidades **************************
    Private _QuantidadePedido As Decimal

    '*********** LISTAS ********************************
    Private _Encargos As ListPedidoXEncargo
    Private _Fixacoes As ListFixacao
    Private _Lancamentos As ListLancamentoItemPedido

    '*********** SALDO *********************************
    Private _SaldoItem As SaldoPedido2015

    '**** Operacao x Encargos **************************
    Private _CodigoOperacaoXEstado As Integer
    Private _OperacaoxEstado As OperacaoXEstado

    Private _Acrescimo As Decimal
    Private _Desconto As Decimal
    Private _AjustarValorOficial As Boolean

#End Region

#Region "Propety"
    Public Property Pedido() As Pedido
        Get
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
        End Set
    End Property

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
            _Descricao = ""
        End Set
    End Property

    Public ReadOnly Property Produto() As Produto
        Get
            If _Produto Is Nothing And Not Me.CodigoProduto Is Nothing Then _Produto = New Produto(Me.CodigoProduto)
            Return _Produto
        End Get
    End Property

    Public Property Descricao() As String
        Get
            If _Descricao.Length = 0 Then _Descricao = Produto.Nome
            Return _Descricao
        End Get
        Set(value As String)
            _Descricao = value
        End Set
    End Property

    Public ReadOnly Property UnidadePedido As String
        Get
            If String.IsNullOrWhiteSpace(_UnidadeComercializacao) Then
                If Produto.UnidadesDeComercializacao.Count > 0 AndAlso Produto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = Me.CodigoUnidadeComercializacao).First.FatorConversao = 0 Then
                    _UnidadeComercializacao = Me.CodigoUnidadeComercializacao & " (X) " & Produto.Unidade
                Else
                    _UnidadeComercializacao = Me.CodigoUnidadeComercializacao & " (" & Me.UnidadeComercializacaoFatorDeConversao & ") " & Produto.Unidade
                End If
            End If
            Return _UnidadeComercializacao
        End Get
    End Property

    Public ReadOnly Property UnidadeFaturamento As String
        Get
            Return Produto.Unidade
        End Get
    End Property

    Public Property CodigoUnidadeComercializacao As String
        Get
            Return _CodigoUnidadeComercializacao
        End Get
        Set(value As String)
            _CodigoUnidadeComercializacao = value
            If Me.Produto Is Nothing Then
                _UnidadeComercializacaoFatorDeConversao = 1
            ElseIf Me.Produto Is Nothing OrElse Me.Produto.UnidadesDeComercializacao.Count = 0 Then
                _UnidadeComercializacaoFatorDeConversao = 1
            Else
                _UnidadeComercializacaoFatorDeConversao = Me.Produto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = _CodigoUnidadeComercializacao).First.FatorConversao
            End If
        End Set
    End Property

    Public ReadOnly Property UnidadeComercializacaoFatorDeConversao As Decimal
        Get
            Return IIf(_UnidadeComercializacaoFatorDeConversao = 0, 1, _UnidadeComercializacaoFatorDeConversao)
        End Get
    End Property

    Public Property Retencao As Boolean
        Get
            Return _Retencao
        End Get
        Set(value As Boolean)
            _Retencao = value
        End Set
    End Property

    Public Property CodigoClassificacao As Integer
        Get
            Return _CodigoClassificacao
        End Get
        Set(value As Integer)
            _CodigoClassificacao = value
            _Classificacao = Nothing
        End Set
    End Property

    Public ReadOnly Property Classificacao() As Classificacao
        Get
            If _Classificacao Is Nothing And Me.CodigoClassificacao > 0 Then _Classificacao = New Classificacao(Me.CodigoClassificacao)
            Return _Classificacao
        End Get
    End Property

    Public Property TemNota() As Boolean
        Get
            Return TemNotaLancada()
        End Get
        Set(ByVal value As Boolean)
            _TemNota = value
        End Set
    End Property

    '**************** Programado Pedido ***********************************
    Public ReadOnly Property QuantidadePedidoFaturamento() As Decimal
        Get
            If Lancamentos Is Nothing Then
                Return 0
            Else
                Return Lancamentos.QuantidadeTotalPrdFaturamento
            End If
        End Get
    End Property

    Public ReadOnly Property QuantidadePedidoComercializacao() As Decimal
        Get
            If Lancamentos Is Nothing Then
                Return 0
            Else
                Return Lancamentos.QuantidadeTotalPrdComercializacao
            End If
        End Get
    End Property

    Public ReadOnly Property UnitarioMedio() As Decimal
        Get
            ' Faz o unitario medio e testa se o unitario inteiro * a quantidade é igual ao total do produto ai ele devolve o inteiro se nao devolve o unitario calculado com decimais
            'ex valor 29583,33 / qtde 25000 * Base de calculo do produto 60 = unitario de 70,999992                   Unitario decimal
            '         unitario arredondado pra inteiro 71 * qtde 25000 / base produto 60 =  valor pedido 29583,33     se for igual devolve 71 unitario inteiro se nao devolve 70,999992 unitario decimal 
            Dim Unt As Decimal
            If Lancamentos Is Nothing Then
                Return 0
            Else
                If Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    If Lancamentos.QuantidadeTotalPrdComercializacao = 0 Then
                        Unt = Lancamentos(0).UnitarioOficial * Me.UnidadeComercializacaoFatorDeConversao
                    Else
                        Unt = (Lancamentos.TotalOficialPrd / Lancamentos.QuantidadeTotalPrdComercializacao)
                    End If

                    If Math.Round(Lancamentos.QuantidadeTotalPrdComercializacao * Math.Round(Unt, 0, MidpointRounding.AwayFromZero), 2, MidpointRounding.AwayFromZero) = Lancamentos.TotalOficialPrd Then
                        If Lancamentos.TotalOficialPrd = 0 Then
                            Return Unt
                        Else
                            Return Math.Round(Unt, 0, MidpointRounding.AwayFromZero)
                        End If
                    ElseIf Lancamentos.QuantidadeTotalPrdComercializacao = 0 Then
                        Return Unt
                    Else
                        Return (Lancamentos.TotalOficialPrd / Lancamentos.QuantidadeTotalPrdComercializacao)
                    End If
                Else
                    If Lancamentos.QuantidadeTotalPrdComercializacao = 0 Then
                        Unt = Lancamentos(0).UnitarioMoeda * Me.UnidadeComercializacaoFatorDeConversao
                    Else
                        Unt = (Lancamentos.TotalMoedaPrd / Lancamentos.QuantidadeTotalPrdComercializacao)
                    End If

                    If Math.Round(Lancamentos.QuantidadeTotalPrdComercializacao * Math.Round(Unt, 0, MidpointRounding.AwayFromZero), 2, MidpointRounding.AwayFromZero) = Lancamentos.TotalMoedaPrd Then
                        If Lancamentos.TotalMoedaPrd = 0 Then
                            Return Unt
                        Else
                            Return Math.Round(Unt, 0, MidpointRounding.AwayFromZero)
                        End If
                    ElseIf Lancamentos.QuantidadeTotalPrdComercializacao = 0 Then
                        Return Unt
                    Else
                        Return (Lancamentos.TotalMoedaPrd / Lancamentos.QuantidadeTotalPrdComercializacao)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property UnitarioMedioFaturamento() As Decimal
        Get
            ' Faz o unitario medio e testa se o unitario inteiro * a quantidade é igual ao total do produto ai ele devolve o inteiro se nao devolve o unitario calculado com decimais
            'ex valor 29583,33 / qtde 25000 * Base de calculo do produto 60 = unitario de 70,999992                   Unitario decimal
            '         unitario arredondado pra inteiro 71 * qtde 25000 / base produto 60 =  valor pedido 29583,33     se for igual devolve 71 unitario inteiro se nao devolve 70,999992 unitario decimal 
            Dim Unt As Decimal
            If Lancamentos Is Nothing Then
                Return 0
            Else
                If Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    Unt = Lancamentos(0).UnitarioOficial

                    If Math.Round(Lancamentos.QuantidadeTotalPrdFaturamento * Math.Round(Unt, 0, MidpointRounding.AwayFromZero), 2, MidpointRounding.AwayFromZero) = Lancamentos.TotalOficialPrd Then
                        If Lancamentos.TotalOficialPrd = 0 Then
                            Return Unt
                        Else
                            Return Math.Round(Unt, 0, MidpointRounding.AwayFromZero)
                        End If
                    ElseIf Lancamentos.QuantidadeTotalPrdFaturamento = 0 Then
                        Return Unt
                    Else
                        Return (Lancamentos.TotalOficialPrd / Lancamentos.QuantidadeTotalPrdFaturamento)
                    End If
                Else
                    Unt = Lancamentos(0).UnitarioMoeda

                    If Math.Round(Lancamentos.QuantidadeTotalPrdFaturamento * Math.Round(Unt, 0, MidpointRounding.AwayFromZero), 2, MidpointRounding.AwayFromZero) = Lancamentos.TotalMoedaPrd Then
                        If Lancamentos.TotalMoedaPrd = 0 Then
                            Return Unt
                        Else
                            Return Math.Round(Unt, 0, MidpointRounding.AwayFromZero)
                        End If
                    ElseIf Lancamentos.QuantidadeTotalPrdFaturamento = 0 Then
                        Return Unt
                    Else
                        Return (Lancamentos.TotalMoedaPrd / Lancamentos.QuantidadeTotalPrdFaturamento)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property UnitarioMedioComercializacao() As Decimal
        Get
            ' Faz o unitario medio e testa se o unitario inteiro * a quantidade é igual ao total do produto ai ele devolve o inteiro se nao devolve o unitario calculado com decimais
            'ex valor 29583,33 / qtde 25000 * Base de calculo do produto 60 = unitario de 70,999992                   Unitario decimal
            '         unitario arredondado pra inteiro 71 * qtde 25000 / base produto 60 =  valor pedido 29583,33     se for igual devolve 71 unitario inteiro se nao devolve 70,999992 unitario decimal 
            Dim Unt As Decimal
            If Lancamentos Is Nothing Then
                Return 0
            Else
                If Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    If Lancamentos.QuantidadeTotalPrdComercializacao = 0 Then
                        Unt = Lancamentos(0).UnitarioOficial
                    Else
                        Unt = (Lancamentos.TotalOficialPrd / Lancamentos.QuantidadeTotalPrdComercializacao)
                    End If

                    If Math.Round(Lancamentos.QuantidadeTotalPrdComercializacao * Math.Round(Unt, 0, MidpointRounding.AwayFromZero), 2, MidpointRounding.AwayFromZero) = Lancamentos.TotalOficialPrd Then
                        If Lancamentos.TotalOficialPrd = 0 Then
                            Return Unt
                        Else
                            Return Math.Round(Unt, 0, MidpointRounding.AwayFromZero)
                        End If
                    ElseIf Lancamentos.QuantidadeTotalPrdComercializacao = 0 Then
                        Return Unt
                    Else
                        Return (Lancamentos.TotalOficialPrd / Lancamentos.QuantidadeTotalPrdComercializacao)
                    End If
                Else
                    If Lancamentos.QuantidadeTotalPrdComercializacao = 0 Then
                        Unt = Lancamentos(0).UnitarioMoeda
                    Else
                        Unt = (Lancamentos.TotalOficialPrd / Lancamentos.QuantidadeTotalPrdComercializacao)
                    End If

                    If Math.Round(Lancamentos.QuantidadeTotalPrdComercializacao * Math.Round(Unt, 0, MidpointRounding.AwayFromZero), 2, MidpointRounding.AwayFromZero) = Lancamentos.TotalMoedaPrd Then
                        If Lancamentos.TotalMoedaPrd = 0 Then
                            Return Unt
                        Else
                            Return Math.Round(Unt, 0, MidpointRounding.AwayFromZero)
                        End If
                    ElseIf Lancamentos.QuantidadeTotalPrdComercializacao = 0 Then
                        Return Unt
                    Else
                        Return (Lancamentos.TotalMoedaPrd / Lancamentos.QuantidadeTotalPrdComercializacao)
                    End If
                End If
            End If
        End Get
    End Property

    Public ReadOnly Property PedidoValor() As Decimal
        Get
            If Lancamentos Is Nothing Then
                Return 0
            Else
                If Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    Return Lancamentos.TotalOficialPrd
                Else
                    Return Lancamentos.TotalMoedaPrd
                End If
            End If
        End Get
    End Property

    '***************** Saldo **********************************************
    Public Property SaldoItem As SaldoPedido2015
        Get
            If _SaldoItem Is Nothing Then
                Dim CifOficial As String = ""
                Dim CifPedido As String = ""
                If Pedido.IUD = "I" Then
                    CifOficial = (New Moedas).Where(Function(s) s.Classificacao = eTiposMoeda.Oficial).FirstOrDefault.Simbolo
                    CifPedido = Pedido.Moeda.Simbolo
                End If


                _SaldoItem = Pedido.SaldoItensPedido.Where(Function(s) s.CodigoProduto = Me.CodigoProduto).FirstOrDefault
                If _SaldoItem Is Nothing Then
                    _SaldoItem = New SaldoPedido2015()
                    _SaldoItem.CodigoProduto = Me.CodigoProduto
                    _SaldoItem.Tipo = IIf(Me.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR, 2, IIf(Me.Pedido.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS, 2, 1))
                    If Pedido.IUD = "I" Then
                        _SaldoItem.CifraoOficial = CifOficial
                        _SaldoItem.CifraoPedido = CifPedido
                    End If
                    Pedido.SaldoItensPedido.Add(_SaldoItem)
                End If
            End If
            Return _SaldoItem
        End Get
        Set(value As SaldoPedido2015)
            _SaldoItem = value
        End Set
    End Property

    '**************** Listas **********************************
    Public Property Encargos() As ListPedidoXEncargo
        Get
            If _Encargos Is Nothing Then _Encargos = New ListPedidoXEncargo(Me)
            Return _Encargos
        End Get
        Set(ByVal value As ListPedidoXEncargo)
            _Encargos = value
        End Set
    End Property

    Public Property Fixacoes() As ListFixacao
        Get
            If _Fixacoes Is Nothing Then _Fixacoes = New ListFixacao(Me)
            Return _Fixacoes
        End Get
        Set(ByVal value As ListFixacao)
            _Fixacoes = value
        End Set
    End Property

    Public Property Lancamentos() As ListLancamentoItemPedido
        Get
            If _Lancamentos Is Nothing Then _Lancamentos = New ListLancamentoItemPedido(Me)
            Return _Lancamentos
        End Get
        Set(ByVal value As ListLancamentoItemPedido)
            _Lancamentos = value
        End Set
    End Property

    '*************** Indica que Operacao x encargos foi usado ***********
    Public Property CodigoOperacaoXEstado As Integer
        Get
            Return _CodigoOperacaoXEstado
        End Get
        Set(value As Integer)
            _CodigoOperacaoXEstado = value
            _OperacaoxEstado = Nothing
        End Set
    End Property

    Public ReadOnly Property OperacaoxEstado As OperacaoXEstado
        Get
            If _OperacaoxEstado Is Nothing And Me.CodigoOperacaoXEstado > 0 Then
                Dim Parametros As New OperacaoXEstado
                Parametros.Codigo = Me.CodigoOperacaoXEstado
                _OperacaoxEstado = New OperacaoXEstado(Parametros)
            End If
            Return _OperacaoxEstado
        End Get
    End Property

    'Utilizado para manter os valores de acréscimos ou descontos quando são feitos novos lançamentos (Complemento, Estorno)'
    Public Property Acrescimo() As Decimal
        Get
            Return _Acrescimo
        End Get
        Set(ByVal value As Decimal)
            _Acrescimo = value
        End Set
    End Property

    Public Property Desconto() As Decimal
        Get
            Return _Desconto
        End Get
        Set(ByVal value As Decimal)
            _Desconto = value
        End Set
    End Property

    Public Property AjustarValorOficial As Boolean
        Get
            Return _AjustarValorOficial
        End Get
        Set(value As Boolean)
            _AjustarValorOficial = value
        End Set
    End Property

#End Region

#Region "Functions"
    Public Function TemNotaLancada() As Boolean
        If Me.CodigoProduto.Length > 0 Then
            Dim objBanco As New AcessaBanco()

            Dim strSQL As String = "SELECT TOP 1 1 " & vbCrLf & _
                                   "  FROM NotasFiscais N " & vbCrLf & _
                                   " INNER JOIN NotasFiscaisXItens nXi " & vbCrLf & _
                                   "    ON nXi.Empresa_id       = n.Empresa_id " & vbCrLf & _
                                   "   AND nXi.EndEmpresa_id   = n.EndEmpresa_id " & vbCrLf & _
                                   "   AND nXi.Cliente_id      = n.Cliente_id " & vbCrLf & _
                                   "   AND nXi.EndCliente_id   = n.EndCliente_id " & vbCrLf & _
                                   "   AND nXi.EntradaSaida_id = n.EntradaSaida_id " & vbCrLf & _
                                   "   AND nXi.Serie_id        = n.Serie_id " & vbCrLf & _
                                   "   AND nXi.Nota_id         = n.Nota_id " & vbCrLf & _
                                   " WHERE NOT n.Situacao       IN (2,9,10) " & vbCrLf & _
                                   "   AND nXi.Empresa_Id    = '" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                                   "   AND nXi.EndEmpresa_Id = " & Me.Pedido.EnderecoEmpresa & vbCrLf & _
                                   "   AND nXi.Pedido        = " & Me.Pedido.Codigo & vbCrLf & _
                                   "   AND nXi.Produto_Id    = '" & Me.CodigoProduto & "'" & vbCrLf
            If objBanco.ConsultaDataSet(strSQL, "NotasFiscaisXItens").Tables(0).Rows.Count > 0 Then
                Return True
            Else : Return False
            End If
        Else
            Return False
        End If
    End Function
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    'Tem que adquirir o numerador do pedido
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I", "U"
                Sql = " MERGE PedidoXItem AS Dest" & vbCrLf &
                      " USING (Select '" & Me.Pedido.CodigoEmpresa & "' as Empresa_Id," & Me.Pedido.EnderecoEmpresa & " as EndEmpresa_Id," & Pedido.Codigo & " as Pedido_Id," & Me.CodigoProduto & " as Produto_Id) AS Ori" & vbCrLf &
                      "    ON Dest.Empresa_Id    = Ori.Empresa_Id" & vbCrLf &
                      "   and Dest.EndEmpresa_Id = Ori.EndEmpresa_Id" & vbCrLf &
                      "   and Dest.Pedido_Id     = Ori.Pedido_Id" & vbCrLf &
                      "   and Dest.Produto_Id    = Ori.Produto_Id" & vbCrLf &
                      "  WHEN NOT MATCHED" & vbCrLf &
                      "    THEN Insert (Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id," & vbCrLf &
                      "                 Classificacao, UnidadeComercializacao, FatorConversao, OperacaoXEstado, Retencao)" & vbCrLf &
                      "		values ('" & Me.Pedido.CodigoEmpresa & "', " & Me.Pedido.EnderecoEmpresa & "," & Pedido.Codigo & ",'" & Me.CodigoProduto & "'," & vbCrLf &
                      Me.Classificacao.Codigo & ",'" & Me.CodigoUnidadeComercializacao & "'," & Str(Me.UnidadeComercializacaoFatorDeConversao) & ", " & CodigoOperacaoXEstado & "," & IIf(Me.Retencao, 1, 0) & ") " & vbCrLf &
                      "  WHEN MATCHED " & vbCrLf &
                      "    THEN Update set " & vbCrLf &
                      "          Classificacao             = " & Me.Classificacao.Codigo & vbCrLf &
                      "         ,UnidadeComercializacao    = '" & Me.CodigoUnidadeComercializacao & "'" & vbCrLf &
                      "         ,FatorConversao            = " & Str(Me.UnidadeComercializacaoFatorDeConversao) & " " & vbCrLf &
                      "         ,OperacaoXEstado           = " & Me.CodigoOperacaoXEstado & vbCrLf &
                      "         ,Retencao                  = '" & Me.Retencao.ToString & "';" & vbCrLf
                Sqls.Add(Sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                SalvarTabelasRelacionadasSql(Sqls)
                Sql = "DELETE PedidoXItem " & _
                      " WHERE Empresa_Id    ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & Me.Pedido.EnderecoEmpresa & vbCrLf & _
                      "   AND Pedido_Id     = " & Me.Pedido.Codigo & vbCrLf & _
                      "   AND Produto_Id    ='" & Me.CodigoProduto & "';" & vbCrLf
                Sqls.Add(Sql)
            Case Else
                SalvarTabelasRelacionadasSql(Sqls)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        Me.Lancamentos.SalvarSql(Sqls)
        Me.Encargos.SalvarSql(Sqls)
        Me.Fixacoes.SalvarSql(Sqls)
    End Sub

#End Region
End Class

