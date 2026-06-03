Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()>
Public Class PedidoXParcelas
    Inherits Negocio.MovimentacoesFinanceiras
    Implements IBaseEntity

#Region "Estruturas"

    Private Structure structManter
        Public Codigo As Integer
        Public Data As DateTime
        Public CarteiraDoTitulo As Integer
    End Structure

#End Region

#Region "Constantes"

    Public Const HISTORICO_PEDIDO As String = "{3} PARC. {0}/{1} REF. PEDIDO {2}"
    Public Const HISTORICO_PEDIDO_ADICIONAL As String = "{1} PARC. ADICIONAL REF. PEDIDO {0}"

#End Region

#Region "Variáveis Locais"

    Private _Pedido As Negocio.Pedido

#End Region

#Region "Construtores"
    Public Sub New()

    End Sub
    Public Sub New(ByVal Pedido As Negocio.Pedido)
        Me.Pedido = Pedido
    End Sub

#End Region

#Region "Propriedades"

    Public Property Pedido() As Negocio.Pedido
        Get
            Return _Pedido
        End Get
        Set(ByVal value As Negocio.Pedido)
            _Pedido = value

            If Not _Pedido.SubOperacao Is Nothing Then
                Me.EntradaSaida = _Pedido.SubOperacao.EntradaSaida
                'If _Pedido.Codigo > 0 Then Selecionar(_Pedido.CodigoEmpresa, _Pedido.EnderecoEmpresa, _Pedido.Codigo, _Pedido.Situacao)
                If _Pedido.Codigo > 0 Then Selecionar(_Pedido.CodigoEmpresa, _Pedido.EnderecoEmpresa, _Pedido.Codigo, 1, _Pedido)
            End If

        End Set
    End Property

#End Region

#Region "Métodos"

    Public Shared Function Existe(ByVal Pedido As Integer, ByVal Tipo As eEntradaSaida, ByVal Provisao As eProvisao, Optional ByVal Empresa As String = "", Optional ByVal EndEmpresa As Integer = 0) As Boolean
        Dim objBanco As New AcessaBanco()
        Try
            Dim strSQL As String = "SELECT Provisao " &
                                   "FROM ContasA" & IIf(Tipo = eEntradaSaida.Entrada, "Pagar", "Receber") & " " &
                                   "WHERE Pedido = " & Pedido.ToString() & " AND Situacao = 1 "

            If Provisao <> Nothing Then strSQL &= "AND Provisao = " & Convert.ToInt32(Provisao) & " "
            If Not String.IsNullOrWhiteSpace(Empresa) Then
                strSQL &= "AND Empresa= '" & Empresa & "' "
                strSQL &= "AND EndEmpresa= " & EndEmpresa & " "
            End If


            Dim dsMovimentacoes As DataSet = objBanco.ConsultaDataSet(strSQL, "Contas")

            If dsMovimentacoes.Tables(0).Rows.Count > 0 Then Return True Else Return False
        Catch ex As Exception
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

    Public Sub EstornoPedido(ByVal parcelas As PedidoXParcelas, ByVal totalEstorno As Decimal)
        If (parcelas.Any(Function(s) s.Provisao = s.Provisao = eProvisao.Provisao)) Then
            If Me.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                Exit Sub
            Else
                Throw New Exception("Não há Título em provisão no pedido para realizar o estorno.")
            End If
        End If

        If (Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Provisao And x.Situacao = eSituacao.Normal).Sum(Function(s) s.ValorDocumentoOficial) < totalEstorno) Then
            If Me.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                Exit Sub
            Else
                Throw New Exception("Não há valor suficiente em provisão no pedido para realizar o estorno.")
            End If
        End If

        For Each parc In parcelas.Where(Function(p) p.Provisao = eProvisao.Provisao And p.Situacao = eSituacao.Normal).OrderByDescending(Function(s) s.Codigo)
            Dim vlrParcela = IIf(Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, parc.ValorLiquidoOficial, parc.ValorLiquidoMoeda)

            If (vlrParcela > totalEstorno) Then
                Dim vlrParc = vlrParcela - totalEstorno
                parc.IUD = "U"
                AddValorMoeda(parc, vlrParc)
                'parc.ValorDocumentoOficial = vlrParc
                totalEstorno = 0
            Else
                Dim vlrParc = 0
                parc.IUD = "U"
                parc.Situacao = eSituacao.Excluido
                parc.ValorDocumentoOficial = 0
                totalEstorno = totalEstorno - parc.ValorLiquidoOficial
            End If

            If totalEstorno = 0 Then Exit For
        Next
    End Sub

    Public Sub CriaParcelaDeComplementoEmDolar(ByVal valorComplemento As Decimal)
        Dim dateVencimento As DateTime
        Dim strHistorico As String = String.Empty
        Dim codigoProvisao = 0

        For Each p In Me.Pedido.CondicaoPagamento.Periodo
            dateVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(p))

            strHistorico = String.Format(HISTORICO_PEDIDO, (Me.Count() + 1).ToString(), (Me.Count() + 1).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))
        Next

        Dim parcela = AtribuirDados(IIf(Me.Pedido.MomentoFinanceiro = 3, eProvisao.Provisao, eProvisao.Previsao), codigoProvisao,
                                     IIf(Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, Me.Pedido.Itens(0).Produto.CodigoCarteiraCompra, Me.Pedido.Itens(0).Produto.CodigoCarteiraVenda),
                                     Me.Pedido.Indexador.Codigo, Me.Pedido.Moeda.Codigo, Me.Pedido.DataPedido, dateVencimento,
                                     Me.Pedido.UnidadeNegocio, Me.Pedido.Empresa, Me.Pedido.Cliente, Me.Pedido.SubOperacao,
                                     valorComplemento, 0, strHistorico, Me.Pedido, Me.Pedido.ContaBancariaSelecionada, 0)
        parcela.IUD = "I"
        Me.Add(parcela)
    End Sub

    Public Sub CriarParcelamento(Optional ByVal ManterDatas As Boolean = False, Optional ByVal objPedVencimentos As Hashtable = Nothing)

        If Left(Me.Pedido.CodigoEmpresa, 8) = "24450490" OrElse Left(Me.Pedido.CodigoEmpresa, 8) = "44979506" OrElse Left(Me.Pedido.CodigoEmpresa, 8) = "41921621" Then
            CriarProvisao(ManterDatas, objPedVencimentos)
        Else
            CriarVariasProvisao(ManterDatas, objPedVencimentos)
        End If

    End Sub

    Public Sub CriarProvisao(Optional ByVal ManterDatas As Boolean = False, Optional ByVal objPedVencimentos As Hashtable = Nothing)

        Dim nParcelas As PedidoXParcelas = Me

        Dim totalDeParcelas As Integer = 0

        If Not objPedVencimentos Is Nothing Then totalDeParcelas = objPedVencimentos.Count

        Dim totalBaixa As Decimal
        Dim totalPrevisao As Decimal
        Dim totalProvisao As Decimal
        Dim total As Decimal

        If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
            totalBaixa = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Baixa And (x.Situacao = eSituacao.Normal Or x.Situacao = eSituacao.RemessaBancaria Or x.Situacao = eSituacao.EndossoTitulo)).Sum(Function(s) s.ValorDocumentoOficial)
            totalPrevisao = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Previsao And (x.Situacao = eSituacao.Normal Or x.Situacao = eSituacao.RemessaBancaria Or x.Situacao = eSituacao.EndossoTitulo)).Sum(Function(s) s.ValorDocumentoOficial)
            totalProvisao = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Provisao And x.Situacao = eSituacao.Normal).Sum(Function(s) s.ValorDocumentoOficial)
            total = Me.Pedido.Itens.Sum(Function(l) l.Encargos.Where(Function(x) x.IUD <> "D" And x.Encargo.Codigo = "LIQUIDO").Sum(Function(i) i.ValorOficial))
        Else
            totalBaixa = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Baixa And (x.Situacao = eSituacao.Normal Or x.Situacao = eSituacao.RemessaBancaria Or x.Situacao = eSituacao.EndossoTitulo)).Sum(Function(s) s.ValorDocumentoMoeda)
            totalPrevisao = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Previsao And (x.Situacao = eSituacao.Normal Or x.Situacao = eSituacao.RemessaBancaria Or x.Situacao = eSituacao.EndossoTitulo)).Sum(Function(s) s.ValorDocumentoMoeda)
            totalProvisao = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Provisao And x.Situacao = eSituacao.Normal).Sum(Function(s) s.ValorDocumentoMoeda)
            total = Me.Pedido.Itens.Sum(Function(l) l.Encargos.Where(Function(x) x.IUD <> "D" And x.Encargo.Codigo = "LIQUIDO").Sum(Function(i) i.ValorMoeda))
        End If

        'Dim totalNormalEComplemento = Me.Pedido.Itens.Sum(Function(l) l.Lancamentos.Where(Function(x) x.IUD = "I" And x.TipoLancamento = eTiposLancamentosPedidos.Normal).Sum(Function(i) i.Total)) +
        '                     Me.Pedido.Itens.Sum(Function(l) l.Lancamentos.Where(Function(x) x.IUD = "I" And x.TipoLancamento = eTiposLancamentosPedidos.Complemento).Sum(Function(i) i.Total))

        'Faz o estorno no pedido.
        Dim temEstorno = Me.Pedido.Itens.Any(Function(l) l.Lancamentos.Any(Function(x) x.IUD = "I" And x.TipoLancamento = eTiposLancamentosPedidos.Estorno))
        If (temEstorno) Then
            Dim totalEstorno = total - (totalBaixa + totalPrevisao + IIf(totalProvisao < 0, totalProvisao * (-1), totalProvisao))
            EstornoPedido(nParcelas, IIf(totalEstorno < 0, totalEstorno * (-1), totalEstorno))
            Exit Sub
        End If

        'Dim saldoFinanceiro = (totalBaixa + totalPrevisao + IIf(totalProvisao < 0, totalProvisao * (-1), totalProvisao)) -
        '    (Me.Pedido.Itens.Sum(Function(l) l.Lancamentos.Where(Function(x) x.IUD <> "I" And x.TipoLancamento <> eTiposLancamentosPedidos.Estorno).Sum(Function(i) i.Total)) -
        '     Me.Pedido.Itens.Sum(Function(l) l.Lancamentos.Where(Function(x) x.IUD <> "I" And x.TipoLancamento = eTiposLancamentosPedidos.Estorno).Sum(Function(i) i.Total)))

        'Dim saldoFinanceiro = (totalBaixa + totalPrevisao + IIf(totalProvisao < 0, totalProvisao * (-1), totalProvisao))
        Dim saldoFinanceiro = (totalBaixa + totalPrevisao)

        Dim saldoComplemento As Decimal = total - saldoFinanceiro

        'If saldoFinanceiro > 0 AndAlso (totalBaixa > 0 OrElse totalPrevisao > 0) Then
        '    saldoComplemento = total - saldoFinanceiro

        '    'saldoComplemento += totalProvisao

        '    If saldoComplemento = 0 AndAlso totalProvisao > 0 Then saldoComplemento = totalProvisao

        'Else
        '    saldoComplemento = total
        'End If

        'Dim saldoComplemento = totalNormalEComplemento - (totalEstorno + saldoFinanceiro)
        'saldoComplemento = total - totalProvisao

        'If saldoComplemento = 0 AndAlso totalProvisao = 0 Then
        If Not saldoComplemento > 0 Then
            'SALDO ZERO MAIS AINDA TEM TITULO EM PROVISÃO O MESMO DEVE SER EXCLUIDO - FURLAN - 21/01/2025
            If saldoComplemento = 0 AndAlso totalProvisao > 0 Then
                For Each parc In nParcelas
                    If parc.Provisao = eProvisao.Provisao Then
                        parc.Situacao = 3
                    End If
                Next
                Exit Sub
            Else
                'SE SALDO FOR NEGATIVO NÃO DEVE DEIXAR FAZER NADA
                Throw New Exception("O Pedido não tem saldo ou já foram pagas todas as parcelas deste pedido.")
                Exit Sub
            End If
        End If

        'If (Not Me.Pedido.IUD = "I" AndAlso Not Me.Any(Function(s) s.Provisao = eProvisao.Previsao Or s.Provisao = eProvisao.Provisao)) AndAlso (totalEstorno > 0 AndAlso Me.All(Function(s) s.Provisao = eProvisao.Baixa)) Then
        '    Throw New Exception("Já foram pagas todas as parcelas deste pedido.")
        'End If

        'Caso o pedido não tenha provisão: gera uma com o valor do complemento
        Dim valorParcela As Decimal
        Dim intParcela As Integer
        Dim dateVencimento As DateTime

        Dim condicaoPagamentoFixa As New CondicaoPagamento(1)

        If Not nParcelas.Any(Function(s) s.Provisao = eProvisao.Provisao And s.Situacao = eSituacao.Normal) Then

            intParcela = 0

            valorParcela = Math.Round(saldoComplemento / condicaoPagamentoFixa.Parcelas, 2, MidpointRounding.AwayFromZero)

            For Each p In condicaoPagamentoFixa.Periodo

                intParcela += 1

                dateVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(p))

                Dim strHistorico As String = String.Format(HISTORICO_PEDIDO, (Me.Count() + 1).ToString(), (Me.Count() + 1).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))

                Dim novaParc As New MovimentacaoFinanceira

                If intParcela = condicaoPagamentoFixa.Periodo.Count Then
                    AddValorMoeda(novaParc, saldoComplemento)
                Else
                    AddValorMoeda(novaParc, valorParcela)
                End If

                'Buscar sempre a provisão do pedido
                Dim titulosEmProvisao = New ListTitulo(" Pedido = " & Me.Pedido.Codigo & " And Empresa = '" & Me.Pedido.CodigoEmpresa & "' and EndEmpresa = " & Me.Pedido.EnderecoEmpresa & " And Provisao = " & eProvisao.Provisao)

                Dim codigoProvisao = 0
                If titulosEmProvisao.Any() Then
                    codigoProvisao = titulosEmProvisao.OrderByDescending(Function(t) t.UsuarioInclusaoData).FirstOrDefault().Codigo
                End If

                Dim parcela = AtribuirDados(IIf(Me.Pedido.MomentoFinanceiro = 3, eProvisao.Provisao, eProvisao.Previsao), codigoProvisao,
                                        IIf(Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, Me.Pedido.Itens(0).Produto.CodigoCarteiraCompra, Me.Pedido.Itens(0).Produto.CodigoCarteiraVenda),
                                        Me.Pedido.Indexador.Codigo, Me.Pedido.Moeda.Codigo, Me.Pedido.DataPedido, dateVencimento,
                                        Me.Pedido.UnidadeNegocio, Me.Pedido.Empresa, Me.Pedido.Cliente, Me.Pedido.SubOperacao,
                                        novaParc.ValorLiquidoOficial, novaParc.ValorLiquidoMoeda, strHistorico, Me.Pedido, Me.Pedido.ContaBancariaSelecionada, 0)
                parcela.IUD = "I"
                Me.Add(parcela)

                saldoComplemento -= valorParcela
            Next


            Exit Sub
        End If


        'altera a provisão
        If totalDeParcelas = 1 AndAlso condicaoPagamentoFixa.Parcelas = 1 Then
            For Each parc In nParcelas.Where(Function(p) p.Provisao = eProvisao.Provisao And p.Situacao = eSituacao.Normal).OrderByDescending(Function(s) s.Codigo)
                If (saldoComplemento < 0) Then
                    Dim vlrParc = parc.ValorLiquidoOficial + saldoComplemento
                    parc.IUD = "U"
                    If (vlrParc <= 0) Then
                        vlrParc = parc.ValorLiquidoOficial
                        AddValorMoeda(parc, 0)
                        parc.Situacao = eSituacao.Cancelado
                    Else
                        AddValorMoeda(parc, (parc.ValorLiquidoOficial + saldoComplemento))
                    End If

                    saldoComplemento = vlrParc + saldoComplemento
                Else
                    'Dim vlrParc = parc.ValorLiquidoOficial + saldoComplemento
                    Dim vlrParc = saldoComplemento
                    parc.IUD = "U"
                    AddValorMoeda(parc, vlrParc)
                    saldoComplemento = 0
                End If

                If saldoComplemento = 0 Then Exit For
            Next
        Else
            If totalDeParcelas = condicaoPagamentoFixa.Parcelas Then

                valorParcela = Math.Round(saldoComplemento / condicaoPagamentoFixa.Parcelas, 2, MidpointRounding.AwayFromZero)

                intParcela = 0

                For Each parc In nParcelas.Where(Function(p) p.Provisao = eProvisao.Provisao And p.Situacao = eSituacao.Normal).OrderBy(Function(s) s.Codigo)

                    parc.IUD = "U"
                    dateVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(condicaoPagamentoFixa.Periodo(intParcela)))

                    parc.Vencimento = dateVencimento
                    parc.DataProrrogacao = dateVencimento
                    parc.DataMoeda = dateVencimento
                    parc.Historico = String.Format(HISTORICO_PEDIDO, (intParcela + 1).ToString(), (condicaoPagamentoFixa.Periodo.Count).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))

                    AddValorMoeda(parc, valorParcela)
                    If (intParcela + 1) = condicaoPagamentoFixa.Periodo.Count Then
                        AddValorMoeda(parc, saldoComplemento)
                    Else
                        AddValorMoeda(parc, valorParcela)
                    End If

                    intParcela += 1

                    saldoComplemento -= valorParcela
                Next
            Else

                valorParcela = Math.Round(saldoComplemento / condicaoPagamentoFixa.Parcelas, 2, MidpointRounding.AwayFromZero)

                If totalDeParcelas > condicaoPagamentoFixa.Parcelas Then
                    intParcela = 0
                    For Each parc In nParcelas.Where(Function(p) p.Provisao = eProvisao.Provisao And p.Situacao = eSituacao.Normal).OrderBy(Function(s) s.Codigo)

                        parc.IUD = "U"

                        If intParcela < condicaoPagamentoFixa.Periodo.Count Then
                            dateVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(condicaoPagamentoFixa.Periodo(intParcela)))

                            parc.Vencimento = dateVencimento
                            parc.DataProrrogacao = dateVencimento
                            parc.DataMoeda = dateVencimento
                            parc.Historico = String.Format(HISTORICO_PEDIDO, (intParcela + 1).ToString(), (condicaoPagamentoFixa.Periodo.Count).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))

                            AddValorMoeda(parc, valorParcela)
                        End If

                        If (intParcela + 1) > condicaoPagamentoFixa.Periodo.Count OrElse valorParcela < 0 Then
                            parc.Situacao = eSituacao.Excluido
                        ElseIf (intParcela + 1) = condicaoPagamentoFixa.Periodo.Count Then
                            AddValorMoeda(parc, saldoComplemento)
                        Else
                            AddValorMoeda(parc, valorParcela)
                        End If

                        intParcela += 1

                        saldoComplemento -= valorParcela
                    Next
                Else
                    intParcela = 1
                    For Each p In condicaoPagamentoFixa.Periodo

                        If intParcela <= totalDeParcelas Then

                            nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString).IUD = "U"

                            dateVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(condicaoPagamentoFixa.Periodo(intParcela)))

                            nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString).Vencimento = dateVencimento
                            nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString).DataProrrogacao = dateVencimento
                            nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString).DataMoeda = dateVencimento
                            nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString).Historico = String.Format(HISTORICO_PEDIDO, (intParcela + 1).ToString(), (condicaoPagamentoFixa.Periodo.Count).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))

                            AddValorMoeda(nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString), valorParcela)
                        Else

                            dateVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(condicaoPagamentoFixa.Periodo(intParcela - 1)))

                            Dim strHistorico As String = String.Format(HISTORICO_PEDIDO, (intParcela + 1).ToString(), (condicaoPagamentoFixa.Periodo.Count).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))

                            Dim novaParc As New MovimentacaoFinanceira

                            If intParcela = condicaoPagamentoFixa.Periodo.Count Then
                                AddValorMoeda(novaParc, saldoComplemento)
                            Else
                                AddValorMoeda(novaParc, valorParcela)
                            End If

                            Dim parcela = AtribuirDados(IIf(Me.Pedido.MomentoFinanceiro = 3, eProvisao.Provisao, eProvisao.Previsao), 0,
                                     IIf(Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, Me.Pedido.Itens(0).Produto.CodigoCarteiraCompra, Me.Pedido.Itens(0).Produto.CodigoCarteiraVenda),
                                     Me.Pedido.Indexador.Codigo, Me.Pedido.Moeda.Codigo, Me.Pedido.DataPedido, dateVencimento,
                                     Me.Pedido.UnidadeNegocio, Me.Pedido.Empresa, Me.Pedido.Cliente, Me.Pedido.SubOperacao,
                                     novaParc.ValorLiquidoOficial, novaParc.ValorLiquidoMoeda, strHistorico, Me.Pedido, Me.Pedido.ContaBancariaSelecionada, 0)
                            parcela.IUD = "I"
                            nParcelas.Add(parcela)
                        End If


                        intParcela += 1

                        saldoComplemento -= valorParcela

                    Next
                End If
            End If
        End If
    End Sub

    Public Sub CriarVariasProvisao(Optional ByVal ManterDatas As Boolean = False, Optional ByVal objPedVencimentos As Hashtable = Nothing)

        Dim nParcelas As PedidoXParcelas = Me

        Dim totalDeParcelas As Integer = 0

        If Not objPedVencimentos Is Nothing Then totalDeParcelas = objPedVencimentos.Count

        Dim totalBaixa As Decimal
        Dim totalPrevisao As Decimal
        Dim totalProvisao As Decimal
        Dim total As Decimal

        Dim dataFixa As New DateTime
        Dim bDataFixa As Boolean = False

        bDataFixa = Me.Pedido.DataFixa(dataFixa)

        If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
            totalBaixa = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Baixa And (x.Situacao = eSituacao.Normal Or x.Situacao = eSituacao.RemessaBancaria Or x.Situacao = eSituacao.EndossoTitulo)).Sum(Function(s) s.ValorDocumentoOficial)
            totalPrevisao = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Previsao And (x.Situacao = eSituacao.Normal Or x.Situacao = eSituacao.RemessaBancaria Or x.Situacao = eSituacao.EndossoTitulo)).Sum(Function(s) s.ValorDocumentoOficial)
            totalProvisao = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Provisao And x.Situacao = eSituacao.Normal).Sum(Function(s) s.ValorDocumentoOficial)
            total = Me.Pedido.Itens.Sum(Function(l) l.Encargos.Where(Function(x) x.IUD <> "D" And x.Encargo.Codigo = "LIQUIDO").Sum(Function(i) i.ValorOficial))
        Else
            totalBaixa = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Baixa And (x.Situacao = eSituacao.Normal Or x.Situacao = eSituacao.RemessaBancaria Or x.Situacao = eSituacao.EndossoTitulo)).Sum(Function(s) s.ValorDocumentoMoeda)
            totalPrevisao = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Previsao And (x.Situacao = eSituacao.Normal Or x.Situacao = eSituacao.RemessaBancaria Or x.Situacao = eSituacao.EndossoTitulo)).Sum(Function(s) s.ValorDocumentoMoeda)
            totalProvisao = Me.Pedido.Vencimentos.Where(Function(x) x.Provisao = eProvisao.Provisao And x.Situacao = eSituacao.Normal).Sum(Function(s) s.ValorDocumentoMoeda)
            total = Me.Pedido.Itens.Sum(Function(l) l.Encargos.Where(Function(x) x.IUD <> "D" And x.Encargo.Codigo = "LIQUIDO").Sum(Function(i) i.ValorMoeda))
        End If

        'Dim totalNormalEComplemento = Me.Pedido.Itens.Sum(Function(l) l.Lancamentos.Where(Function(x) x.IUD = "I" And x.TipoLancamento = eTiposLancamentosPedidos.Normal).Sum(Function(i) i.Total)) +
        '                     Me.Pedido.Itens.Sum(Function(l) l.Lancamentos.Where(Function(x) x.IUD = "I" And x.TipoLancamento = eTiposLancamentosPedidos.Complemento).Sum(Function(i) i.Total))

        'Faz o estorno no pedido.
        Dim temEstorno = Me.Pedido.Itens.Any(Function(l) l.Lancamentos.Any(Function(x) x.IUD = "I" And x.TipoLancamento = eTiposLancamentosPedidos.Estorno))
        If (temEstorno) Then
            Dim totalEstorno = total - (totalBaixa + totalPrevisao + IIf(totalProvisao < 0, totalProvisao * (-1), totalProvisao))
            EstornoPedido(nParcelas, IIf(totalEstorno < 0, totalEstorno * (-1), totalEstorno))
            Exit Sub
        End If

        'Dim saldoFinanceiro = (totalBaixa + totalPrevisao + IIf(totalProvisao < 0, totalProvisao * (-1), totalProvisao)) -
        '    (Me.Pedido.Itens.Sum(Function(l) l.Lancamentos.Where(Function(x) x.IUD <> "I" And x.TipoLancamento <> eTiposLancamentosPedidos.Estorno).Sum(Function(i) i.Total)) -
        '     Me.Pedido.Itens.Sum(Function(l) l.Lancamentos.Where(Function(x) x.IUD <> "I" And x.TipoLancamento = eTiposLancamentosPedidos.Estorno).Sum(Function(i) i.Total)))

        'Dim saldoFinanceiro = (totalBaixa + totalPrevisao + IIf(totalProvisao < 0, totalProvisao * (-1), totalProvisao))
        Dim saldoFinanceiro = (totalBaixa + totalPrevisao)

        Dim saldoComplemento As Decimal = total - saldoFinanceiro

        'If saldoFinanceiro > 0 AndAlso (totalBaixa > 0 OrElse totalPrevisao > 0) Then
        '    saldoComplemento = total - saldoFinanceiro

        '    'saldoComplemento += totalProvisao

        '    If saldoComplemento = 0 AndAlso totalProvisao > 0 Then saldoComplemento = totalProvisao

        'Else
        '    saldoComplemento = total
        'End If

        'Dim saldoComplemento = totalNormalEComplemento - (totalEstorno + saldoFinanceiro)
        'saldoComplemento = total - totalProvisao

        'If saldoComplemento = 0 AndAlso totalProvisao = 0 Then
        If Not saldoComplemento > 0 Then
            'SALDO ZERO MAIS AINDA TEM TITULO EM PROVISÃO O MESMO DEVE SER EXCLUIDO - FURLAN - 21/01/2025
            If saldoComplemento <= 0 AndAlso totalProvisao > 0 Then
                For Each parc In nParcelas
                    If parc.Provisao = eProvisao.Provisao Then
                        parc.Situacao = 3
                    End If
                Next
                Exit Sub
            ElseIf saldoComplemento > 0 Then
                'SE SALDO FOR NEGATIVO NÃO DEVE DEIXAR FAZER NADA
                Throw New Exception("O Pedido não tem saldo ou já foram pagas todas as parcelas deste pedido.")
                Exit Sub
            Else
                Exit Sub
            End If
        End If

        'If (Not Me.Pedido.IUD = "I" AndAlso Not Me.Any(Function(s) s.Provisao = eProvisao.Previsao Or s.Provisao = eProvisao.Provisao)) AndAlso (totalEstorno > 0 AndAlso Me.All(Function(s) s.Provisao = eProvisao.Baixa)) Then
        '    Throw New Exception("Já foram pagas todas as parcelas deste pedido.")
        'End If

        'Caso o pedido não tenha provisão: gera uma com o valor do complemento
        Dim valorParcela As Decimal
        Dim intParcela As Integer
        Dim dateVencimento As DateTime

        If Not nParcelas.Any(Function(s) s.Provisao = eProvisao.Provisao And s.Situacao = eSituacao.Normal) Then

            intParcela = 0

            valorParcela = Math.Round(saldoComplemento / Me.Pedido.CondicaoPagamento.Parcelas, 2, MidpointRounding.AwayFromZero)

            For Each p In Me.Pedido.CondicaoPagamento.Periodo

                intParcela += 1

                If bDataFixa Then
                    dateVencimento = dataFixa
                Else
                    dateVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(p))
                End If

                Dim strHistorico As String = String.Format(HISTORICO_PEDIDO, (Me.Count() + 1).ToString(), (Me.Count() + 1).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))

                Dim novaParc As New MovimentacaoFinanceira

                If intParcela = Me.Pedido.CondicaoPagamento.Periodo.Count Then
                    AddValorMoeda(novaParc, saldoComplemento)
                Else
                    AddValorMoeda(novaParc, valorParcela)
                End If

                'Buscar sempre a provisão do pedido
                Dim titulosEmProvisao = New ListTitulo(" Pedido = " & Me.Pedido.Codigo & " And Empresa = '" & Me.Pedido.CodigoEmpresa & "' and EndEmpresa = " & Me.Pedido.EnderecoEmpresa & " And Provisao = " & eProvisao.Provisao)

                Dim codigoProvisao = 0
                If titulosEmProvisao.Any() Then
                    codigoProvisao = titulosEmProvisao.OrderByDescending(Function(t) t.UsuarioInclusaoData).FirstOrDefault().Codigo
                End If

                Dim parcela = AtribuirDados(IIf(Me.Pedido.MomentoFinanceiro = 3, eProvisao.Provisao, eProvisao.Previsao), codigoProvisao,
                                        IIf(Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, Me.Pedido.Itens(0).Produto.CodigoCarteiraCompra, Me.Pedido.Itens(0).Produto.CodigoCarteiraVenda),
                                        Me.Pedido.Indexador.Codigo, Me.Pedido.Moeda.Codigo, Me.Pedido.DataPedido, dateVencimento,
                                        Me.Pedido.UnidadeNegocio, Me.Pedido.Empresa, Me.Pedido.Cliente, Me.Pedido.SubOperacao,
                                        novaParc.ValorLiquidoOficial, novaParc.ValorLiquidoMoeda, strHistorico, Me.Pedido, Me.Pedido.ContaBancariaSelecionada, 0)
                parcela.IUD = "I"
                Me.Add(parcela)

                saldoComplemento -= valorParcela
            Next

            'Dim dateVencimento As DateTime = Pedido.DataPedido.AddDays(Me.Pedido.CondicaoPagamento.Periodo(0))

            Exit Sub
        End If

        'altera a provisão
        If totalDeParcelas = 1 AndAlso Me.Pedido.CondicaoPagamento.Parcelas = 1 Then
            For Each parc In nParcelas.Where(Function(p) p.Provisao = eProvisao.Provisao And p.Situacao = eSituacao.Normal).OrderByDescending(Function(s) s.Codigo)
                If (saldoComplemento < 0) Then
                    Dim vlrParc = parc.ValorLiquidoOficial + saldoComplemento
                    parc.IUD = "U"
                    If (vlrParc <= 0) Then
                        vlrParc = parc.ValorLiquidoOficial
                        AddValorMoeda(parc, 0)
                        parc.Situacao = eSituacao.Cancelado
                    Else
                        AddValorMoeda(parc, (parc.ValorLiquidoOficial + saldoComplemento))
                    End If

                    saldoComplemento = vlrParc + saldoComplemento
                Else
                    'Dim vlrParc = parc.ValorLiquidoOficial + saldoComplemento
                    Dim vlrParc = saldoComplemento
                    parc.IUD = "U"
                    AddValorMoeda(parc, vlrParc)
                    saldoComplemento = 0
                End If

                If saldoComplemento = 0 Then Exit For
            Next
        Else
            If totalDeParcelas = Me.Pedido.CondicaoPagamento.Parcelas Then

                valorParcela = Math.Round(saldoComplemento / Me.Pedido.CondicaoPagamento.Parcelas, 2, MidpointRounding.AwayFromZero)

                intParcela = 0

                For Each parc In nParcelas.Where(Function(p) p.Provisao = eProvisao.Provisao And p.Situacao = eSituacao.Normal).OrderBy(Function(s) s.Codigo)

                    parc.IUD = "U"

                    If bDataFixa Then
                        dateVencimento = dataFixa
                    Else
                        dateVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(Me.Pedido.CondicaoPagamento.Periodo(intParcela)))
                    End If

                    parc.Vencimento = dateVencimento
                    parc.DataProrrogacao = dateVencimento
                    parc.DataMoeda = dateVencimento
                    parc.Historico = String.Format(HISTORICO_PEDIDO, (intParcela + 1).ToString(), (Me.Pedido.CondicaoPagamento.Periodo.Count).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))

                    AddValorMoeda(parc, valorParcela)
                    If (intParcela + 1) = Me.Pedido.CondicaoPagamento.Periodo.Count Then
                        AddValorMoeda(parc, saldoComplemento)
                    Else
                        AddValorMoeda(parc, valorParcela)
                    End If

                    intParcela += 1

                    saldoComplemento -= valorParcela
                Next
            Else

                valorParcela = Math.Round(saldoComplemento / Me.Pedido.CondicaoPagamento.Parcelas, 2, MidpointRounding.AwayFromZero)

                If totalDeParcelas > Me.Pedido.CondicaoPagamento.Parcelas Then
                    intParcela = 0
                    For Each parc In nParcelas.Where(Function(p) p.Provisao = eProvisao.Provisao And p.Situacao = eSituacao.Normal).OrderBy(Function(s) s.Codigo)

                        parc.IUD = "U"

                        If intParcela < Me.Pedido.CondicaoPagamento.Periodo.Count Then

                            If bDataFixa Then
                                dateVencimento = dataFixa
                            Else
                                dateVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(Me.Pedido.CondicaoPagamento.Periodo(intParcela)))
                            End If

                            parc.Vencimento = dateVencimento
                            parc.DataProrrogacao = dateVencimento
                            parc.DataMoeda = dateVencimento
                            parc.Historico = String.Format(HISTORICO_PEDIDO, (intParcela + 1).ToString(), (Me.Pedido.CondicaoPagamento.Periodo.Count).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))

                            AddValorMoeda(parc, valorParcela)
                        End If

                        If (intParcela + 1) > Me.Pedido.CondicaoPagamento.Periodo.Count OrElse valorParcela < 0 Then
                            parc.Situacao = eSituacao.Excluido
                        ElseIf (intParcela + 1) = Me.Pedido.CondicaoPagamento.Periodo.Count Then
                            AddValorMoeda(parc, saldoComplemento)
                        Else
                            AddValorMoeda(parc, valorParcela)
                        End If

                        intParcela += 1

                        saldoComplemento -= valorParcela
                    Next
                Else
                    intParcela = 1
                    For Each p In Me.Pedido.CondicaoPagamento.Periodo

                        If intParcela <= totalDeParcelas Then

                            nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString).IUD = "U"

                            If bDataFixa Then
                                dateVencimento = dataFixa
                            Else
                                dateVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(Me.Pedido.CondicaoPagamento.Periodo(intParcela)))
                            End If

                            nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString).Vencimento = dateVencimento
                            nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString).DataProrrogacao = dateVencimento
                            nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString).DataMoeda = dateVencimento
                            nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString).Historico = String.Format(HISTORICO_PEDIDO, (intParcela + 1).ToString(), (Me.Pedido.CondicaoPagamento.Periodo.Count).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))

                            AddValorMoeda(nParcelas(objPedVencimentos.Keys(intParcela - 1).ToString), valorParcela)
                        Else

                            If bDataFixa Then
                                dateVencimento = dataFixa
                            Else
                                dateVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.DataPedido.AddDays(Me.Pedido.CondicaoPagamento.Periodo(intParcela - 1)))
                            End If

                            Dim strHistorico As String = String.Format(HISTORICO_PEDIDO, (intParcela + 1).ToString(), (Me.Pedido.CondicaoPagamento.Periodo.Count).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))

                            Dim novaParc As New MovimentacaoFinanceira

                            If intParcela = Me.Pedido.CondicaoPagamento.Periodo.Count Then
                                AddValorMoeda(novaParc, saldoComplemento)
                            Else
                                AddValorMoeda(novaParc, valorParcela)
                            End If

                            Dim parcela = AtribuirDados(IIf(Me.Pedido.MomentoFinanceiro = 3, eProvisao.Provisao, eProvisao.Previsao), 0,
                                     IIf(Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, Me.Pedido.Itens(0).Produto.CodigoCarteiraCompra, Me.Pedido.Itens(0).Produto.CodigoCarteiraVenda),
                                     Me.Pedido.Indexador.Codigo, Me.Pedido.Moeda.Codigo, Me.Pedido.DataPedido, dateVencimento,
                                     Me.Pedido.UnidadeNegocio, Me.Pedido.Empresa, Me.Pedido.Cliente, Me.Pedido.SubOperacao,
                                     novaParc.ValorLiquidoOficial, novaParc.ValorLiquidoMoeda, strHistorico, Me.Pedido, Me.Pedido.ContaBancariaSelecionada, 0)
                            parcela.IUD = "I"
                            nParcelas.Add(parcela)
                        End If

                        intParcela += 1

                        saldoComplemento -= valorParcela

                    Next
                End If
            End If
        End If
    End Sub

    Public Sub AddValorMoeda(ByRef parc As MovimentacaoFinanceira, ByVal pValor As Decimal)
        Dim dblValorOficial As Decimal = 0.0
        Dim dblValorMoeda As Decimal = 0
        If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
            dblValorOficial = pValor
        Else
            dblValorMoeda = pValor
            If Me.Pedido.IndiceFixado > 0 Then
                dblValorOficial = Math.Round(Me.Pedido.IndiceFixado * dblValorMoeda, 2, MidpointRounding.AwayFromZero)
            Else
                dblValorOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(dblValorMoeda, Me.Pedido.Indexador.Codigo, Me.Pedido.DataPedido), 2, MidpointRounding.AwayFromZero)
            End If
        End If

        parc.IUD = "U"
        parc.ValorDocumentoOficial = dblValorOficial
        parc.ValorLiquidoOficial = dblValorOficial
        parc.ValorDocumentoMoeda = dblValorMoeda
        parc.ValorLiquidoMoeda = dblValorMoeda
    End Sub

    Public Function AjustaParcelas(ByVal Apartir As Integer, ByVal ValorOriginal As Decimal, ByVal ValorNovo As Decimal) As String
        If (Apartir = Me.Count - 1) Then
            If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Me(Apartir).ValorDocumentoOficial = ValorOriginal
                Me(Apartir).ValorLiquidoOficial = ValorOriginal
            Else
                Me(Apartir).ValorDocumentoMoeda = ValorOriginal
                Me(Apartir).ValorLiquidoMoeda = ValorOriginal
            End If

            If ValorNovo = ValorOriginal Then
                Return ""
            Else
                Return "Parcela Unica ou Ultima Parcela nao pode ser alterada"
            End If
        End If

        Dim saldo As Decimal = ValorOriginal - ValorNovo
        Dim numParcelas As Integer = (Me.Count - 1 - Apartir)
        Dim parcelas As Decimal = Math.Round(saldo / numParcelas, 2)
        Dim diferenca As Decimal = saldo - (parcelas * numParcelas)

        saldo = 0

        If (Apartir < Me.Count - 1) And ValorNovo > ValorOriginal Then
            For i As Integer = Apartir + 1 To Me.Count - 1
                If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    saldo += Me(i).ValorDocumentoOficial
                Else
                    saldo += Me(i).ValorDocumentoMoeda
                End If
            Next

            If saldo <= ValorNovo - ValorOriginal Then
                If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    Me(Apartir).ValorDocumentoOficial = ValorOriginal
                    Me(Apartir).ValorLiquidoOficial = ValorOriginal
                Else
                    Me(Apartir).ValorDocumentoMoeda = ValorOriginal
                    Me(Apartir).ValorLiquidoMoeda = ValorOriginal
                End If
                Return "Valor Informado ultrapassa o Valor da transacao"
            End If
        End If

        For i As Integer = Apartir + 1 To Me.Count - 1
            If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                If i = Me.Count - 1 Then
                    Me(i).ValorDocumentoOficial += parcelas + diferenca
                Else
                    Me(i).ValorDocumentoOficial += parcelas
                End If
                Me(i).ValorLiquidoOficial = Me(i).ValorDocumentoOficial
            Else
                If i = Me.Count - 1 Then
                    Me(i).ValorDocumentoMoeda += parcelas + diferenca
                Else
                    Me(i).ValorDocumentoMoeda += parcelas
                End If
                Me(i).ValorLiquidoMoeda = Me(i).ValorDocumentoMoeda
            End If
        Next

        If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
            Me(Apartir).ValorDocumentoOficial = ValorNovo
            Me(Apartir).ValorLiquidoOficial = ValorNovo
        Else
            Me(Apartir).ValorDocumentoMoeda = ValorNovo
            Me(Apartir).ValorLiquidoMoeda = ValorNovo
        End If

        Return ""
    End Function

    'Public Sub ModificarParcela(ByVal Indice As Integer, ByVal NovoVencimento As DateTime, ByVal NovoValor As Double)
    '    Dim objParcela As Negocio.MovimentacaoFinanceira = Me(Indice)
    '    Dim dblSomaOficial As Double = 0
    '    Dim dblSomaMoeda As Double = 0
    '    Dim dblValorMoeda As Double = 0
    '    Dim dblResto As Double = 0
    '    Dim intParcelasAbertas As Integer = 0
    '    Dim intParcelasDividir As Integer = 0

    '    Me.OficialAMovimentar = 0
    '    Me.OficialMovimentado = 0

    '    'Passo 1 - Calcular o total das parcelas usadas para ver ser precisa ser recalculado
    '    For intPos As Integer = 0 To Me.Count - 1
    '        If intPos <> Indice Then
    '            dblSomaOficial += Me(intPos).ValorDocumentoOficial
    '            dblSomaMoeda += Me(intPos).ValorDocumentoMoeda

    '            If Me(intPos).Provisao = eProvisao.Baixa OrElse (Me.Pedido.MomentoFinanceiro = 3 AndAlso objParcela.Provisao = eProvisao.Previsao) Then
    '                Me.OficialMovimentado += Me(intPos).ValorDocumentoOficial
    '                Me.MoedaMovimentado += Me(intPos).ValorDocumentoMoeda
    '            Else
    '                intParcelasAbertas += 1
    '                If intPos > Indice Then intParcelasDividir += 1
    '            End If
    '        End If
    '    Next intPos

    '    'Passo 2 - Se não houver outras parcelas abertas, a parcela não será alterada.
    '    'If intParcelasAbertas = 0 Then Exit Sub

    '    'Passo 3 - Verificar se o valor digitado ultrapassa o valor total ou se não sobra resto.
    '    If Pedido.Moeda.Classificacao = Negocio.eTiposMoeda.MoedaEstrangeira Then
    '        If Me.Pedido.IndiceFixado > 0 Then
    '            dblValorMoeda = Math.Round(NovoValor / Me.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
    '        Else
    '            dblValorMoeda = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(NovoValor, Me.Pedido.Indexador.Codigo, Me.Pedido.DataPedido), 2, MidpointRounding.AwayFromZero)
    '        End If
    '        dblResto = Pedido.Itens.LiquidoMoeda - (dblSomaMoeda + dblValorMoeda)
    '    Else
    '        dblResto = Pedido.Itens.LiquidoOficial - (dblSomaOficial + NovoValor)
    '        dblValorMoeda = 0
    '        'If Me.Pedido.IndiceFixado > 0 Then
    '        '    dblValorMoeda = Math.Round(NovoValor / Me.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
    '        'Else
    '        '    dblValorMoeda = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(NovoValor, Me.Pedido.Indexador.Codigo, Me.Pedido.DataPedido), 2, MidpointRounding.AwayFromZero)
    '        'End If
    '    End If

    '    If Me.Count = 1 Or Me.Count = Indice + 1 Or intParcelasDividir = 0 Then
    '        dblValorMoeda += dblResto
    '        dblResto = 0
    '    End If

    '    'Passo 3 - Atribuir os valores novos à lista
    '    With objParcela
    '        If Pedido.IUD = "I" Then .Vencimento = Funcoes.ValidaDataUtil(.CodigoEmpresa, .EnderecoEmpresa, NovoVencimento)
    '        .DataProrrogacao = Funcoes.ValidaDataUtil(.CodigoEmpresa, .EnderecoEmpresa, NovoVencimento)
    '        .DataMoeda = .DataProrrogacao
    '        .DataBaixa = .DataProrrogacao
    '        .ValorDocumentoOficial = NovoValor
    '        .ValorDocumentoMoeda = Math.Round(dblValorMoeda, 2, MidpointRounding.AwayFromZero)
    '        .ValorLiquidoOficial = NovoValor
    '        .ValorLiquidoMoeda = .ValorDocumentoMoeda
    '    End With

    '    'Passo 4 - Se houver resto, é redistribuído para outras parcelas.
    '    If dblResto <> 0 Then
    '        Dim dblRestoDividido As Double = 0

    '        Dim intIndiceInicio As Integer = -1
    '        Dim intIndiceFim As Integer = 0

    '        intIndiceInicio = Indice + 1
    '        intIndiceFim = Me.Count - 1
    '        dblRestoDividido = dblResto / intParcelasDividir
    '        Dim dblResidual As Double = dblResto - (dblRestoDividido * intParcelasDividir)

    '        For intPos As Integer = intIndiceInicio To intIndiceFim
    '            With Me(intPos)
    '                If .Provisao = Negocio.eProvisao.Previsao Then
    '                    If Pedido.Moeda.Classificacao = Negocio.eTiposMoeda.Oficial Then
    '                        .ValorDocumentoOficial += dblRestoDividido + IIf(intPos = intIndiceInicio, dblResidual, 0)
    '                        .ValorDocumentoMoeda = 0
    '                        'If Me.Pedido.IndiceFixado > 0 Then
    '                        '    .ValorDocumentoMoeda = Math.Round(.ValorDocumentoOficial / Me.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
    '                        'Else
    '                        '    .ValorDocumentoMoeda = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(.ValorDocumentoOficial, Pedido.Indexador.Codigo, Pedido.DataPedido), 2, MidpointRounding.AwayFromZero)
    '                        'End If
    '                    Else
    '                        .ValorDocumentoMoeda += dblRestoDividido + IIf(intPos = intIndiceInicio, dblResidual, 0)
    '                        If Me.Pedido.IndiceFixado > 0 Then
    '                            .ValorDocumentoOficial = Math.Round(Me.Pedido.IndiceFixado * .ValorDocumentoMoeda, 2, MidpointRounding.AwayFromZero)
    '                        Else
    '                            .ValorDocumentoOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(.ValorDocumentoMoeda, Pedido.Indexador.Codigo, Pedido.DataPedido), 2, MidpointRounding.AwayFromZero)
    '                        End If
    '                    End If

    '                    .ValorLiquidoOficial = .ValorDocumentoOficial
    '                    .ValorLiquidoMoeda = .ValorDocumentoMoeda
    '                End If
    '            End With
    '        Next
    '    Else
    '        If (Me.Count = 1 Or Me.Count = Indice + 1) And Me(Indice).Provisao = eProvisao.Previsao Then
    '            If Pedido.Moeda.Classificacao = Negocio.eTiposMoeda.Oficial Then
    '                Me(Indice).ValorDocumentoMoeda = 0
    '                'If Me.Pedido.IndiceFixado > 0 Then
    '                '    Me(Indice).ValorDocumentoMoeda = Math.Round(Me(Indice).ValorLiquidoOficial / Me.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
    '                'Else
    '                '    Me(Indice).ValorDocumentoMoeda = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(Me(Indice).ValorLiquidoOficial, Pedido.Indexador.Codigo, Pedido.DataPedido), 2)
    '                'End If
    '            Else
    '                If Me.Pedido.IndiceFixado > 0 Then
    '                    Me(Indice).ValorDocumentoOficial = Math.Round(Me.Pedido.IndiceFixado * Me(Indice).ValorLiquidoMoeda, 2, MidpointRounding.AwayFromZero)
    '                Else
    '                    Me(Indice).ValorDocumentoOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(Me(Indice).ValorLiquidoMoeda, Pedido.Indexador.Codigo, Pedido.DataPedido), 2)
    '                End If
    '            End If

    '            Me(Indice).ValorLiquidoOficial = Me(Indice).ValorDocumentoOficial
    '            Me(Indice).ValorLiquidoMoeda = Me(Indice).ValorDocumentoMoeda
    '        End If
    '    End If

    '    'Passo 5 - Calcular os valores finais
    '    Me.OficialAMovimentar = 0
    '    Me.MoedaAMovimentar = 0

    '    For intPos As Integer = 0 To Me.Count - 1
    '        Me.OficialAMovimentar += Me(intPos).ValorDocumentoOficial
    '        Me.MoedaAMovimentar += Me(intPos).ValorDocumentoMoeda
    '    Next intPos
    'End Sub


    'Public Sub ModificarParcelaDolar(ByVal Indice As Integer, ByVal NovoVencimento As DateTime, ByVal NovoValor As Double)
    '    Dim objParcela As Negocio.MovimentacaoFinanceira = Me(Indice)
    '    Dim dblSomaOficial As Double = 0
    '    Dim dblSomaMoeda As Double = 0
    '    Dim dblValorOficial As Double = 0
    '    Dim dblValorMoeda As Double = 0
    '    Dim dblResto As Double = 0
    '    Dim intParcelasAbertas As Integer = 0
    '    Dim intParcelasDividir As Integer = 0

    '    Me.OficialAMovimentar = 0
    '    Me.OficialMovimentado = 0

    '    'Passo 1 - Calcular o total das parcelas usadas para ver ser precisa ser recalculado
    '    For intPos As Integer = 0 To Me.Count - 1
    '        If intPos <> Indice Then
    '            dblSomaOficial += Me(intPos).ValorDocumentoOficial
    '            dblSomaMoeda += Me(intPos).ValorDocumentoMoeda

    '            If Me(intPos).Provisao = eProvisao.Baixa OrElse (Me.Pedido.MomentoFinanceiro = 3 AndAlso objParcela.Provisao = eProvisao.Previsao) Then
    '                Me.OficialMovimentado += Me(intPos).ValorDocumentoOficial
    '                Me.MoedaMovimentado += Me(intPos).ValorDocumentoMoeda
    '            Else
    '                intParcelasAbertas += 1
    '                If intPos > Indice Then intParcelasDividir += 1
    '            End If
    '        End If
    '    Next intPos

    '    'Passo 2 - Se não houver outras parcelas abertas, a parcela não será alterada.
    '    'If intParcelasAbertas = 0 Then Exit Sub

    '    'Passo 3 - Verificar se o valor digitado ultrapassa o valor total ou se não sobra resto.
    '    If Pedido.Moeda.Classificacao = Negocio.eTiposMoeda.MoedaEstrangeira Then
    '        If Me.Pedido.IndiceFixado > 0 Then
    '            dblValorOficial = Math.Round(Me.Pedido.IndiceFixado * NovoValor, 2, MidpointRounding.AwayFromZero)
    '        Else
    '            dblValorOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(NovoValor, Me.Pedido.Indexador.Codigo, Me.Pedido.DataPedido), 2, MidpointRounding.AwayFromZero)
    '        End If
    '        dblResto = Pedido.Itens.LiquidoMoeda - (dblSomaMoeda + NovoValor)
    '    Else
    '        'If Me.Pedido.IndiceFixado > 0 Then
    '        '    dblValorMoeda = Math.Round(NovoValor / Me.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
    '        'Else
    '        '    dblValorMoeda = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(NovoValor, Me.Pedido.Indexador.Codigo, Me.Pedido.DataPedido), 2, MidpointRounding.AwayFromZero)
    '        'End If
    '        dblValorMoeda = 0
    '        dblResto = Pedido.Itens.LiquidoOficial - (dblSomaOficial + dblValorOficial)
    '    End If

    '    If Me.Count = 1 Or Me.Count = Indice + 1 Or intParcelasDividir = 0 Then
    '        If Pedido.Moeda.Classificacao = Negocio.eTiposMoeda.MoedaEstrangeira Then
    '            dblValorMoeda += dblResto
    '        Else
    '            dblValorOficial += dblResto
    '        End If
    '        dblResto = 0
    '    End If

    '    'Passo 3 - Atribuir os valores novos à lista
    '    With objParcela
    '        If Pedido.IUD = "I" Then .Vencimento = Funcoes.ValidaDataUtil(.CodigoEmpresa, .EnderecoEmpresa, NovoVencimento)
    '        .DataProrrogacao = Funcoes.ValidaDataUtil(.CodigoEmpresa, .EnderecoEmpresa, NovoVencimento)
    '        .DataMoeda = .DataProrrogacao
    '        .DataBaixa = .DataProrrogacao

    '        If Pedido.Moeda.Classificacao = Negocio.eTiposMoeda.MoedaEstrangeira Then
    '            .ValorDocumentoMoeda = NovoValor
    '            .ValorDocumentoOficial = Math.Round(dblValorOficial, 2, MidpointRounding.AwayFromZero)
    '            .ValorLiquidoMoeda = NovoValor
    '            .ValorLiquidoOficial = .ValorDocumentoOficial
    '        Else
    '            .ValorDocumentoOficial = NovoValor
    '            .ValorLiquidoOficial = NovoValor
    '            .ValorDocumentoMoeda = 0
    '            .ValorLiquidoMoeda = 0
    '            '.ValorDocumentoMoeda = Math.Round(dblValorMoeda, 2, MidpointRounding.AwayFromZero)
    '            '.ValorLiquidoMoeda = .ValorDocumentoMoeda
    '        End If
    '    End With

    '    'Passo 4 - Se houver resto, é redistribuído para outras parcelas.
    '    If dblResto <> 0 Then
    '        Dim dblRestoDividido As Double = 0

    '        Dim intIndiceInicio As Integer = -1
    '        Dim intIndiceFim As Integer = 0

    '        intIndiceInicio = Indice + 1
    '        intIndiceFim = Me.Count - 1
    '        dblRestoDividido = dblResto / intParcelasDividir
    '        Dim dblResidual As Double = dblResto - (dblRestoDividido * intParcelasDividir)

    '        For intPos As Integer = intIndiceInicio To intIndiceFim
    '            With Me(intPos)
    '                If .Provisao = Negocio.eProvisao.Previsao Then
    '                    If Pedido.Moeda.Classificacao = Negocio.eTiposMoeda.Oficial Then
    '                        .ValorDocumentoOficial += dblRestoDividido + IIf(intPos = intIndiceInicio, dblResidual, 0)
    '                        .ValorDocumentoMoeda = 0
    '                        'If Me.Pedido.IndiceFixado > 0 Then
    '                        '    .ValorDocumentoMoeda = Math.Round(.ValorDocumentoOficial / Me.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
    '                        'Else
    '                        '    .ValorDocumentoMoeda = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(.ValorDocumentoOficial, Pedido.Indexador.Codigo, Pedido.DataPedido), 2, MidpointRounding.AwayFromZero)
    '                        'End If
    '                    Else
    '                        .ValorDocumentoMoeda += dblRestoDividido + IIf(intPos = intIndiceInicio, dblResidual, 0)
    '                        If Me.Pedido.IndiceFixado > 0 Then
    '                            .ValorDocumentoOficial = Math.Round(Me.Pedido.IndiceFixado * .ValorDocumentoMoeda, 2, MidpointRounding.AwayFromZero)
    '                        Else
    '                            .ValorDocumentoOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(.ValorDocumentoMoeda, Pedido.Indexador.Codigo, Pedido.DataPedido), 2, MidpointRounding.AwayFromZero)
    '                        End If
    '                    End If

    '                    .ValorLiquidoOficial = .ValorDocumentoOficial
    '                    .ValorLiquidoMoeda = .ValorDocumentoMoeda
    '                End If
    '            End With
    '        Next
    '    Else
    '        If (Me.Count = 1 Or Me.Count = Indice + 1) And Me(Indice).Provisao = eProvisao.Previsao Then
    '            If Pedido.Moeda.Classificacao = Negocio.eTiposMoeda.Oficial Then
    '                Me(Indice).ValorDocumentoMoeda = 0
    '                'If Me.Pedido.IndiceFixado > 0 Then
    '                '    Me(Indice).ValorDocumentoMoeda = Math.Round(Me(Indice).ValorLiquidoOficial / Me.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
    '                'Else
    '                '    Me(Indice).ValorDocumentoMoeda = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(Me(Indice).ValorLiquidoOficial, Pedido.Indexador.Codigo, Pedido.DataPedido), 2, MidpointRounding.AwayFromZero)
    '                'End If
    '            Else
    '                If Me.Pedido.IndiceFixado > 0 Then
    '                    Me(Indice).ValorDocumentoOficial = Math.Round(Me.Pedido.IndiceFixado * Me(Indice).ValorLiquidoMoeda, 2, MidpointRounding.AwayFromZero)
    '                Else
    '                    Me(Indice).ValorDocumentoOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(Me(Indice).ValorLiquidoMoeda, Pedido.Indexador.Codigo, Pedido.DataPedido), 2, MidpointRounding.AwayFromZero)
    '                End If
    '            End If

    '            Me(Indice).ValorLiquidoOficial = Me(Indice).ValorDocumentoOficial
    '            Me(Indice).ValorLiquidoMoeda = Me(Indice).ValorDocumentoMoeda
    '        End If
    '    End If

    '    'Passo 5 - Calcular os valores finais
    '    Me.OficialAMovimentar = 0
    '    Me.MoedaAMovimentar = 0

    '    For intPos As Integer = 0 To Me.Count - 1
    '        Me.OficialAMovimentar += Me(intPos).ValorDocumentoOficial
    '        Me.MoedaAMovimentar += Me(intPos).ValorDocumentoMoeda
    '    Next intPos
    'End Sub

    Public Function Excluir() As String
        Dim strSQL As String = ""

        strSQL = "DELETE ContasAPagar" & vbCrLf &
                 " WHERE EmpresaPedido    ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf &
                 "   AND EndEmpresaPedido = " & Me.Pedido.EnderecoEmpresa & vbCrLf &
                 "   AND Pedido           = " & Me.Pedido.Codigo & vbCrLf &
                 "   AND Provisao         > 1; " & vbCrLf

        strSQL &= "DELETE ContasAReceber" & vbCrLf &
                  " WHERE EmpresaPedido    ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf &
                  "   and EndEmpresaPedido = " & Me.Pedido.EnderecoEmpresa & vbCrLf &
                  "   and Pedido           = " & Me.Pedido.Codigo & vbCrLf &
                  "   AND Provisao         > 1; "

        Return strSQL
    End Function

    Public Overloads Function GetSQL(ByVal Modo As eModoAlteracao, Optional ByVal Servidor As String = "", Optional ByVal pMomentoFinanceiro As Integer = 0) As ArrayList
        Dim arrSQL As New ArrayList()

        Select Case Modo
            Case eModoAlteracao.Inclusao
                For Each objItem As Negocio.MovimentacaoFinanceira In Me
                    If Me.Pedido.IUD = "I" OrElse (Me.Pedido.IUD = "U" AndAlso (objItem.IUD IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objItem.IUD))) Then
                        'Furlan - Só pode tentar incluir se Provisão não for 1 - 02/06/2010
                        If Modo = eModoAlteracao.Inclusao And objItem.Provisao <> 1 Then
                            arrSQL.Add(objItem.Excluir())
                            arrSQL.Add(objItem.Incluir(Servidor, objItem.Codigo, pMomentoFinanceiro))
                            ' Inicio do historico financeiro.
                            Dim historico = New HistoricoFinanceiro()
                            arrSQL.Add(historico.AdicionarHistoricoPedido(objItem))
                        End If
                    ElseIf Me.Pedido.IUD = "U" Then
                        If objItem.Situacao = 3 Then
                            arrSQL.Add(objItem.Excluido())
                        Else
                            arrSQL.Add(objItem.Alterar())
                        End If
                    End If
                Next
            Case eModoAlteracao.Exclusao
                If Me.Pedido.Codigo > 0 Then arrSQL.Add(Me.Excluir())
        End Select

        Return arrSQL
    End Function

    Public Function ToDataTablePedidos() As DataTable
        Dim dtParcelamento As New DataTable("MovimentacoesFinanceiras")
        dtParcelamento.Columns.Add("TipoFinanceiro", Type.GetType("System.String"))
        dtParcelamento.Columns.Add("Codigo", Type.GetType("System.Int32"))

        dtParcelamento.Columns.Add("Provisao", Type.GetType("System.Int32"))
        dtParcelamento.Columns.Add("Vencimento", Type.GetType("System.DateTime"))
        dtParcelamento.Columns.Add("DataBaixa", Type.GetType("System.DateTime"))

        dtParcelamento.Columns.Add("DescMoeda", Type.GetType("System.String"))

        dtParcelamento.Columns.Add("DocumentoOficial", Type.GetType("System.Double"))
        dtParcelamento.Columns.Add("AcrescimoOficial", Type.GetType("System.Double"))
        dtParcelamento.Columns.Add("DeducaoOficial", Type.GetType("System.Double"))
        dtParcelamento.Columns.Add("LiquidoOficial", Type.GetType("System.Double"))

        dtParcelamento.Columns.Add("DocumentoMoeda", Type.GetType("System.Double"))
        dtParcelamento.Columns.Add("AcrescimoMoeda", Type.GetType("System.Double"))
        dtParcelamento.Columns.Add("DeducaoMoeda", Type.GetType("System.Double"))
        dtParcelamento.Columns.Add("LiquidoMoeda", Type.GetType("System.Double"))

        For Each objParcelamento As Negocio.MovimentacaoFinanceira In Me.OrderBy(Function(s) s.Provisao)
            Dim drParcela As DataRow = dtParcelamento.NewRow()
            drParcela("TipoFinanceiro") = objParcelamento.TipoFinanceiro

            If objParcelamento.Codigo > 0 Then
                drParcela("Codigo") = objParcelamento.Codigo
            Else
                drParcela("Codigo") = 0
            End If
            drParcela("Provisao") = objParcelamento.Provisao
            drParcela("Vencimento") = objParcelamento.DataProrrogacao
            drParcela("DataBaixa") = IIf(objParcelamento.Provisao = eProvisao.Baixa, objParcelamento.DataBaixa, DBNull.Value)
            drParcela("DescMoeda") = objParcelamento.Moeda.Descricao

            drParcela("DocumentoOficial") = objParcelamento.ValorDocumentoOficial
            drParcela("AcrescimoOficial") = objParcelamento.AcrescimosOficial + objParcelamento.JurosOficial
            drParcela("DeducaoOficial") = objParcelamento.DescontosOficial + objParcelamento.DeducoesOficial
            drParcela("LiquidoOficial") = objParcelamento.ValorLiquidoOficial

            drParcela("DocumentoMoeda") = objParcelamento.ValorDocumentoMoeda
            drParcela("AcrescimoMoeda") = objParcelamento.AcrescimosMoeda + objParcelamento.JurosMoeda
            drParcela("DeducaoMoeda") = objParcelamento.DescontosMoeda + objParcelamento.DeducoesMoeda
            drParcela("LiquidoMoeda") = objParcelamento.ValorLiquidoMoeda

            dtParcelamento.Rows.Add(drParcela)
        Next

        Return dtParcelamento
    End Function

    Public Sub CriarParcelamentoOnMobile(ByVal paramTotal As Decimal)
        Dim nParcelas As PedidoXParcelas = Me
        Dim dateVencimento = Me.Pedido.DataPedido

        Dim strHistorico As String = String.Format(HISTORICO_PEDIDO, 1.ToString(), (Me.Pedido.CondicaoPagamento.Periodo.Count).ToString(), IIf(Me.Pedido.Codigo > 0, Me.Pedido.Codigo.ToString(), "{PEDIDO}"), IIf(Me.EntradaSaida = 0, "PAGTO", "RECTO"))

        Dim novaParc As New MovimentacaoFinanceira

        AddValorMoeda(novaParc, paramTotal)

        Dim parcela = AtribuirDados(IIf(Me.Pedido.MomentoFinanceiro = 3, eProvisao.Provisao, eProvisao.Previsao), 0,
                 IIf(Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, Me.Pedido.Itens(0).Produto.CodigoCarteiraCompra, Me.Pedido.Itens(0).Produto.CodigoCarteiraVenda),
                 Me.Pedido.Indexador.Codigo, Me.Pedido.Moeda.Codigo, Me.Pedido.DataPedido, dateVencimento,
                 Me.Pedido.UnidadeNegocio, Me.Pedido.Empresa, Me.Pedido.Cliente, Me.Pedido.SubOperacao,
                 novaParc.ValorLiquidoOficial, novaParc.ValorLiquidoMoeda, strHistorico, Me.Pedido, Me.Pedido.ContaBancariaSelecionada, 0)
        parcela.IUD = "I"
        nParcelas.Add(parcela)
    End Sub

#End Region

End Class