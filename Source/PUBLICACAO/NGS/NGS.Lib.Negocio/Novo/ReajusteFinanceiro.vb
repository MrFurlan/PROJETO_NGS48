Imports System.Globalization
Imports NGS.Lib.Uteis


'*********************************************************************************************************************
'**************************************  Reajuste Financeiro Pedidos *************************************************
'*********************************************************************************************************************
<Serializable()>
Public Class ReajusteFinanceiro

#Region "Contrutor"

    Public Sub New()

    End Sub
    Public Sub New(ByRef pNotaFiscal As NotaFiscal, Optional FinanceiroNovo As Boolean = True)
        If FinanceiroNovo Then
            _Pedido = pNotaFiscal.Pedido
            _ResumoFinanceiro = New ResumoFinanceiro(pNotaFiscal.Pedido)
            _NF_Original = pNotaFiscal.NotaFiscalOriginal
            _NF_Modificada = pNotaFiscal
            _ValorAtual = 0 'IIf(pNotaFiscal.IUD = "D" Or pNotaFiscal.IUD = "C" Or pNotaFiscal.CodigoSituacao = 2, 0, pNotaFiscal.TotalNota)
            _ValorOriginal = pNotaFiscal.NotaFiscalOriginal.TotalNota
            _Devolucao = pNotaFiscal.SubOperacao.Devolucao
            _TitulosNotaModificada = Nothing
            _TitulosOld = Nothing
            _TitulosNew = Nothing
        Else
            _VencimentosPedido = New ListTitulo
            _VencimentosNota = New ListTitulo

            _Pedido = pNotaFiscal.Pedido
            _NF_Modificada = pNotaFiscal
            _ValorAtual = pNotaFiscal.TotalNota
            _ValorBrutoAtual = pNotaFiscal.TotalNotaBruto

            _NF_Original = pNotaFiscal.NotaFiscalOriginal

            If pNotaFiscal.IUD = "U" Then
                _ValorOriginal = pNotaFiscal.NotaFiscalOriginal.TotalNota

            Else
                _ValorOriginal = NF_Modificada.VencimentosNota.Sum(Function(x) x.ValorDoDocumento)
            End If
            'Para Notas de Devolução
            If pNotaFiscal.SubOperacao.Devolucao Then
                _Devolucao = pNotaFiscal.SubOperacao.Devolucao
                _NF_Devolucoes = New ListNotaFiscalDevolucaoXNotaFiscal
                For Each item As [Lib].Negocio.NotaFiscalXItem In pNotaFiscal.Itens
                    For Each notaDevolucao As [Lib].Negocio.NotaFiscalDevolucaoXNotaFiscal In item.NotasDevolucao
                        If notaDevolucao.QuantidadeDevolucao > 0 Or notaDevolucao.ValorDevolucao > 0 Then
                            _VencimentosNota.AddRange(notaDevolucao.Nota.VencimentosNota)
                            _NF_Devolucoes.Add(notaDevolucao)
                        End If
                    Next
                Next
            End If

            If _NF_Original IsNot Nothing AndAlso Not pNotaFiscal.SubOperacao.Devolucao Then
                For Each tit In NF_Original.VencimentosPedido
                    _VencimentosPedido.Add(tit.Copy())
                Next

                For Each tit In NF_Original.VencimentosNota
                    _VencimentosNota.Add(tit.Copy())
                Next
            Else
                For Each tit In NF_Modificada.VencimentosPedido
                    _VencimentosPedido.Add(tit.Copy())
                Next

                For Each tit In NF_Modificada.VencimentosNota
                    _VencimentosNota.Add(tit.Copy())
                Next
            End If

        End If

    End Sub
#End Region

#Region "Fields"
    Private _Pedido As Pedido

    Private _NF_Original As NotaFiscal
    Private _NF_Modificada As NotaFiscal

    Private _ValorAtual As Decimal
    Private _ValorBrutoAtual As Decimal

    Private _ValorOriginal As Decimal

    Private _Devolucao As Boolean

    Private _TitulosOld As Novo.ListTituloNovo
    Private _TitulosNew As Novo.ListTituloNovo
    Private _TitulosNotaModificada As Novo.ListTituloNovo

    'Financeiro Antigo
    Private _VencimentosPedido As ListTitulo
    Private _VencimentosNota As ListTitulo

    Private _ResumoFinanceiro As ResumoFinanceiro

    Private _NF_Devolucoes As ListNotaFiscalDevolucaoXNotaFiscal
#End Region

#Region "Property"
    Public ReadOnly Property Pedido As Pedido
        Get
            Return _Pedido
        End Get
    End Property

    Public ReadOnly Property NF_Original As NotaFiscal
        Get
            Return _NF_Original
        End Get
    End Property

    Public ReadOnly Property NF_Devolucoes As ListNotaFiscalDevolucaoXNotaFiscal
        Get
            Return _NF_Devolucoes
        End Get
    End Property

    Public ReadOnly Property NF_Modificada As NotaFiscal
        Get
            Return _NF_Modificada
        End Get
    End Property

    Public ReadOnly Property ValorAtual As Decimal
        Get
            Return _ValorAtual
        End Get
    End Property

    Public ReadOnly Property ValorOriginal As Decimal
        Get
            Return _ValorOriginal
        End Get
    End Property

    Public ReadOnly Property Devolucao As Boolean
        Get
            Return _Devolucao
        End Get
    End Property

    Public Property TitulosOld As Novo.ListTituloNovo
        Get
            If _TitulosOld Is Nothing Then
                Dim where As String
                where = "     Empresa    ='" & Pedido.CodigoEmpresa & "'" & vbCrLf &
                        " And EndEmpresa = " & Pedido.EnderecoEmpresa & vbCrLf &
                        " And Pedido     = " & Pedido.Codigo
                _TitulosOld = New Novo.ListTituloNovo(where)
            End If
            Return _TitulosOld
        End Get
        Set(value As Novo.ListTituloNovo)
            _TitulosOld = value
        End Set
    End Property

    Public Property TitulosNew As Novo.ListTituloNovo
        Get
            If _TitulosNew Is Nothing Then
                _TitulosNew = New Novo.ListTituloNovo
                _TitulosNew.NF = NF_Modificada
            End If
            Return _TitulosNew
        End Get
        Set(value As Novo.ListTituloNovo)
            _TitulosNew = value
        End Set
    End Property

    Public Property TitulosNotaModificada As Novo.ListTituloNovo
        Get
            If _TitulosNotaModificada Is Nothing Then
                _TitulosNotaModificada = New Novo.ListTituloNovo
            End If
            Return _TitulosNotaModificada
        End Get
        Set(value As Novo.ListTituloNovo)
            _TitulosNotaModificada = value
        End Set
    End Property

    'Financeiro
    Public Property VencimentosPedido As ListTitulo
        Get
            If _VencimentosPedido Is Nothing Then
                _VencimentosPedido = New ListTitulo
            End If
            Return _VencimentosPedido
        End Get
        Set(value As ListTitulo)
            _VencimentosPedido = value
        End Set
    End Property

    Public Property VencimentosNota As ListTitulo
        Get
            If _VencimentosNota Is Nothing Then
                _VencimentosNota = New ListTitulo
            End If
            Return _VencimentosNota
        End Get
        Set(value As ListTitulo)
            _VencimentosNota = value
        End Set
    End Property

    Public Property ResumoFinanceiro As ResumoFinanceiro
        Get
            Return _ResumoFinanceiro
        End Get
        Set(value As ResumoFinanceiro)
            _ResumoFinanceiro = value
        End Set
    End Property

    Public Property ValorBrutoAtual As Decimal
        Get
            Return _ValorBrutoAtual
        End Get
        Set(value As Decimal)
            _ValorBrutoAtual = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function Valida(pvalor As Decimal) As ArrayList
        'pValor se refere ao valor reajustado na nota positivo se a nota aumentou e negativo se a nota diminuiu
        '0 Verdadeiro Ou Falso
        '1 Msg - Sucesso ou Erro 
        Dim Retorno As New ArrayList

        If pvalor = 0 Then
            Retorno.Add(False)
            Retorno.Add("Não Houve alteração no financeiro, a nota pode ser alterada.")
            Return Retorno
        End If

        If pvalor > 0 Then
            '*****************************************************************************
            '****************** AUMENTOU O VALOR *****************************************
            '*****************************************************************************
            If Devolucao Then
                '*****************************************************************************
                '******************  É UMA DEVOLUCAO  ****************************************
                '*****************************************************************************
                ' + S
                If ResumoFinanceiro.ValorTitulosEmProvisao >= pvalor Then
                    Retorno.Add(True)
                    Retorno.Add("O Pedido tem Saldo de Provisao para compensar, a nota pode ser alterada.")
                    Return Retorno
                ElseIf pvalor - ResumoFinanceiro.ValorTitulosEmProvisao - (ResumoFinanceiro.ValorAdiantamentoOriginal - ResumoFinanceiro.SaldoAdiantamento) <= 0 Then
                    Retorno.Add(True)
                    If ResumoFinanceiro.ValorTitulosEmProvisao > 0 Then
                        Retorno.Add("Pedido com Provisao para compensar o saldo sera transformado em adiantamento, a nota pode ser alterada.")
                    Else
                        Retorno.Add("O Valor Alterado sera transformado em adiantamento, a nota pode ser alterada.")
                    End If
                    Return Retorno
                Else
                    Retorno.Add(False)
                    Retorno.Add("O Pedido nao tem saldo suficiente para suportar esta alteração")
                    Return Retorno
                End If
            Else
                '*****************************************************************************
                '****************** NÃO É UMA DEVOLUCAO  *************************************
                '*****************************************************************************
                ' + N
                If ResumoFinanceiro.SaldoAdiantamento >= pvalor Then
                    Retorno.Add(True)
                    Retorno.Add("O Pedido tem Saldo de Adiantamento para baixar, a nota pode ser alterada.")
                    Return Retorno
                ElseIf pvalor - ResumoFinanceiro.SaldoAdiantamento - ResumoFinanceiro.ValorTitulosEmPrevisao <= 0 Then
                    Retorno.Add(True)
                    If ResumoFinanceiro.SaldoAdiantamento > 0 Then
                        Retorno.Add("Pedido tera Adiantamentos baixados e o saldo sera transformado em Provisão, a nota vai ser alterada.")
                    Else
                        Retorno.Add("O Valor Alterado sera transformado em Provisao, a nota pode ser alterada.")
                    End If
                    Return Retorno
                Else
                    Retorno.Add(False)
                    Retorno.Add("O Pedido nao tem saldo suficiente para suportar esta alteração")
                    Return Retorno
                End If
            End If
        Else
            'Transfor pvalor num valor absoluto tira o negativo dele
            pvalor = Math.Abs(pvalor)
            '*****************************************************************************
            '****************** DIMINUIU O VALOR *****************************************
            '*****************************************************************************
            Dim VlrNFProvisao As Decimal = NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Provisao).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
            Dim VlrNFBaixado As Decimal = NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And Not s.isAdiantamento And Not s.isBaixaAdiantamento).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
            Dim VlrNFBaixaAdiantamento As Decimal = NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isBaixaAdiantamento).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)

            Dim VlrNFCompensado As Decimal = NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Compensado).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
            Dim VlrNFSaldoAdiantamento As Decimal = NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isAdiantamento).Sum(Function(s) s.Adiantamento.SaldoValor)

            If Devolucao Then
                '*****************************************************************************
                '******************  É UMA DEVOLUCAO  ****************************************
                '*****************************************************************************
                ' - S
                If VlrNFCompensado >= pvalor Then
                    Retorno.Add(True)
                    Retorno.Add("A compensacao sera diminuida e o valor compensado voltara para provisao, a nota pode ser alterada.")
                    Return Retorno
                ElseIf VlrNFCompensado + VlrNFSaldoAdiantamento >= pvalor Then
                    Retorno.Add(True)
                    If VlrNFCompensado > 0 Then
                        Retorno.Add("A Compensacao sera desfeita e os titulos compensadoa voltaram para provisao o valor restante sera Diminuido do Adiantamento, a nota pode ser alterada.")
                    Else
                        Retorno.Add("O Valor Alterado sera Diminuido do Adiantamento, a nota pode ser alterada.")
                    End If
                    Return Retorno
                ElseIf ResumoFinanceiro.SaldoAdiantamento >= pvalor - VlrNFCompensado Then
                    Retorno.Add(True)
                    Retorno.Add("O Adiantamento gerado pela nota ja tem baixas, os titulos serao transferidos para baixar outros adiantamentos ou serao criados titulos em provisão, a nota pode ser alterada.")
                    Return Retorno
                Else
                    Retorno.Add(False)
                    Retorno.Add("Alteracao da Nota Para Valor Menor numa Devolucao gerou um erro inexperado contacte o suporte")
                    Return Retorno
                End If
            Else
                '*****************************************************************************
                '****************** NÃO É UMA DEVOLUCAO  *************************************
                '*****************************************************************************
                ' - N
                If VlrNFProvisao >= pvalor Then
                    Retorno.Add(True)
                    Retorno.Add("Os titulos em provisao serao diminuidos e o valor voltara para previsao, a nota pode ser alterada.")
                    Return Retorno
                ElseIf VlrNFBaixaAdiantamento + VlrNFProvisao >= pvalor Then
                    Retorno.Add(True)
                    If VlrNFProvisao > 0 Then
                        Retorno.Add("Os titulos em provisao serao diminuidos e o valor voltara para previsao o saldo sera diminuido da baixa de adiantamento efetuada, a nota pode ser alterada.")
                    Else
                        Retorno.Add("O valor Alterado sera diminuido do valor da baixa de adiantamento, a nota pode ser alterada.")
                    End If
                    Return Retorno
                ElseIf VlrNFCompensado + VlrNFBaixaAdiantamento + VlrNFProvisao >= pvalor Then
                    Retorno.Add(True)
                    Retorno.Add("Os Titulos compensados nesta nossa serao diminuidos e compensados por provisoes de outras notas caso nao exista outras provisoes sera gerado um adiantamento na nota que originou a compensacao, a nota pode ser alterada.")
                    Return Retorno
                ElseIf VlrNFBaixado > 0 Then
                    Retorno.Add(True)
                    Retorno.Add("O Titulo desta nota já foi baixado parte ou  o valor total dela sera transformada num Adiantamento, a nota pode ser alterada.")
                    Return Retorno
                Else
                    Retorno.Add(False)
                    Retorno.Add("Alteracao da Nota Para Valor Menor gerou um erro inexperado contacte o suporte")
                    Return Retorno
                End If
            End If
        End If
        Retorno.Add(False)
        Retorno.Add("Erro inexperado contacte o suporte")
        Return Retorno
    End Function

    ' Funcao Que reajusta o Financeiro do Pedido de acordo com a alteracao da nota
    Public Function Reajusta() As ArrayList
        '0 Verdadeiro Ou Falso
        '1 Msg - Sucesso ou Erro 
        Dim Retorno As New ArrayList

        Retorno = Valida(ValorAtual - ValorOriginal)

        If Retorno(0) = False Then
            Return Retorno
        End If

        If ValorAtual > ValorOriginal Then
            If Devolucao Then
                '+S
                AumentouValorNotaDevolucao(ValorAtual - ValorOriginal)
            Else
                '+N
                AumentouValorNota(ValorAtual - ValorOriginal)
            End If
        Else
            If Devolucao Then
                '-S
                DiminuiuValorNotaDevolucao(Math.Abs(ValorAtual - ValorOriginal))
            Else
                '-N
                DiminuiuValorNota(Math.Abs(ValorAtual - ValorOriginal))
            End If
        End If

        Return Retorno
    End Function

    Public Function ReajustaNotaDeEntrada() As ArrayList
        '0 Verdadeiro Ou Falso
        '1 Msg - Sucesso ou Erro 
        Dim Retorno As New ArrayList

        If ValorAtual > ValorOriginal Then
            AumentouValorNotaDeEntrada(ValorAtual - ValorOriginal)
        Else
            DiminuiuValorNotaDeEntrada(Math.Abs(ValorAtual - ValorOriginal))
        End If

        Return Retorno
    End Function

    Public Function ReajustaNotaDeSaida() As ArrayList
        '0 Verdadeiro Ou Falso
        '1 Msg - Sucesso ou Erro 
        Dim Retorno As New ArrayList

        If ValorAtual > ValorOriginal Then
            AumentouValorNotaDeSaida(ValorAtual - ValorOriginal)
        Else
            DiminuiuValorNotaDeSaida(Math.Abs(ValorAtual - ValorOriginal))
        End If

        Return Retorno
    End Function

    Public Function ReajustaNota() As ArrayList
        '0 Verdadeiro Ou Falso
        '1 Msg - Sucesso ou Erro 
        Dim Retorno As New ArrayList

        If ValorAtual > ValorOriginal Then
            AumentouValorNota(ValorAtual - ValorOriginal)
        Else
            DiminuiuValorNota(Math.Abs(ValorAtual - ValorOriginal))
        End If

        Return Retorno
    End Function

    Public Function ReajustaNotaDeDevolucao() As ArrayList
        '0 Verdadeiro Ou Falso
        '1 Msg - Sucesso ou Erro 
        Dim Retorno As New ArrayList

        EstornoNotaDeDevolucao(ValorAtual, ValorBrutoAtual)

        Return Retorno
    End Function

    Public Function ReajustaTroca() As ArrayList
        Dim Retorno As New ArrayList
        Retorno = Valida(ValorAtual - ValorOriginal)
        Dim vlrNota As Decimal = Math.Abs(ValorAtual - ValorOriginal)
        If Devolucao Then
            For Each tit In NF_Original.Titulos
                tit.IUD = "D"
                'Desfaz a compensação dos títulos que foram compensados pela NF Dev.
                Dim CodigoTiuloComp As Integer = tit.Codigo
                For Each titComp As Novo.TituloNovo In NF_Original.TitulosPedido.Where(Function(s) s.CodigoTituloCompensacao = CodigoTiuloComp)
                    titComp.IUD = "U"
                    titComp.CodigoProvisao = eProvisao.Provisao
                    titComp.CodigoTituloCompensacao = 0
                    TitulosNew.Add(titComp)
                Next
                TitulosNew.Add(tit)
            Next
        Else
            For Each tit In NF_Original.Titulos
                tit.IUD = "D"
                If tit.CodigoTituloCompensacao = 0 Then
                    'Pedido de Troca
                    Dim CodigoTituloComp As Integer = tit.Codigo
                    For Each titTroca In NF_Original.Pedido.PedidoTroca.Titulos.Where(Function(s) s.CodigoTituloCompensacao = CodigoTituloComp)
                        titTroca.IUD = "U"
                        titTroca.CodigoProvisao = eProvisao.Provisao
                        titTroca.CodigoTituloCompensacao = 0
                        TitulosNew.Add(titTroca)
                    Next
                Else
                    'Pedido de Troca
                    Dim CodigoTituloComp As Integer = tit.CodigoTituloCompensacao
                    For Each titTroca In NF_Original.Pedido.PedidoTroca.Titulos.Where(Function(s) s.Codigo = CodigoTituloComp)
                        titTroca.IUD = "U"
                        titTroca.CodigoProvisao = eProvisao.Provisao
                        titTroca.CodigoTituloCompensacao = 0
                        TitulosNew.Add(titTroca)
                    Next
                End If

                TitulosNew.Add(tit)
                Dim origem As Integer = tit.CodigoTituloOrigem
                If origem > 0 Then
                    Dim titprev = NF_Original.TitulosPedido.Where(Function(s) s.Codigo = origem).FirstOrDefault
                    If titprev IsNot Nothing Then
                        titprev.IUD = "U"
                        titprev.Valores.EncargoValorDocumento.Valor += vlrNota
                    Else
                        Dim where As String = "Titulo_Id = " & origem
                        Dim list = New Novo.ListTituloNovo(where)
                        titprev = list(0)
                        titprev.IUD = "U"
                        titprev.CodigoSituacao = 1
                        titprev.Valores.EncargoValorDocumento.Valor = vlrNota
                    End If
                    TitulosNew.Add(titprev)
                End If
            Next
        End If

        Return Retorno
    End Function

    Public Function ReajustaAFixar() As ArrayList
        Dim Retorno As New ArrayList

        For Each Tit In NF_Original.Titulos
            If Tit.CodigoProvisao = eProvisao.Provisao Then
                Tit.IUD = "U"
                Tit.CodigoProvisao = eProvisao.Previsao
                TitulosNew.Add(Tit)
            ElseIf Tit.Baixas_AdiantamentoEfetuadas.Count > 0 Then
                Tit.IUD = "C"
                TitulosNew.Add(Tit)
            Else
                Dim titAdi As Novo.TituloNovo = CriaTitulo("ADI", Tit.Valores.EncargoValorDocumento.Valor, Nothing, NF_Original, 0, Tit.CodigoFixacao)
                TitulosNew.Add(titAdi)
            End If
        Next

        Retorno.Add(True)
        Retorno.Add("A fixação voltou o valor financeiro para previsão")

        Return Retorno
    End Function

    ' + N ok 
    Private Sub AumentouValorNota(ByRef pValor As Decimal)
        For Each tit In NF_Original.Titulos
            TitulosNotaModificada.Add(tit)
        Next
        '***************************************************************************
        '******************* Se tiver saldo Baixa Adiantamento *********************
        '***************************************************************************
        If ResumoFinanceiro.SaldoAdiantamento > 0 Then
            'Faz um laço para verificar se existe adiantamentos abertos que nao foram baixados na nota, se existir vai ser criado um titulo de baixa
            For Each adab As Novo.AdiantamentoNovo In Pedido.AdiantamentosAbertos.Where(Function(x) x.SaldoValor > 0)
                'Se chegou aqui os adiantamentos que foram feito baixa na nota nao terao + saldo entao eles tem q ser ignorados
                Dim titBaixa As Novo.TituloNovo
                If adab.SaldoValor >= pValor Then
                    titBaixa = CriaTitulo("BX", pValor, adab)
                    TitulosNew.Add(titBaixa)
                Else
                    titBaixa = CriaTitulo("BX", adab.SaldoValor, adab)
                    TitulosNew.Add(titBaixa)
                    pValor -= adab.SaldoValor
                End If
                TitulosNotaModificada.Add(titBaixa)
            Next
        End If
        '****************************************************************
        '******************* Previsao Para Provisao *********************
        '****************************************************************
        For Each titprev In TitulosOld.Where(Function(x) x.CodigoProvisao = 2 And x.CodigoSituacao = 1).OrderBy(Function(s) s.Reprogramacao).ThenBy(Function(s) s.Codigo)
            Dim titProv As Novo.TituloNovo
            If titprev.Valores.EncargoValorDocumento.Valor >= pValor Then
                titprev.Valores.EncargoValorDocumento.Valor -= pValor

                titprev.IUD = "U"
                TitulosNew.Add(titprev)


                titProv = CriaTitulo("PRO", pValor, Nothing, Nothing, titprev.Codigo)
                TitulosNew.Add(titProv)
                TitulosNotaModificada.Add(titProv)
                Exit For
            Else
                titprev.IUD = "C"
                pValor -= titprev.Valores.EncargoValorDocumento.Valor
                TitulosNew.Add(titprev)

                titProv = CriaTitulo("PRO", titprev.Valores.EncargoValorDocumento.Valor, Nothing, Nothing, titprev.Codigo)
                TitulosNew.Add(titProv)
                TitulosNotaModificada.Add(titProv)
            End If
            If pValor = 0 Then Exit For
        Next
    End Sub

    ' + S ok
    Private Sub AumentouValorNotaDevolucao(ByRef pValor As Decimal)
        For Each tit In NF_Original.Titulos
            TitulosNotaModificada.Add(tit)
        Next

        '**********************************************************************************
        '******************* Se tiver provisao no pedido é compensado *********************
        '**********************************************************************************
        Dim vlrComp As Decimal
        If ResumoFinanceiro.ValorTitulosEmProvisao > 0 Then
            Dim titcomp As Novo.TituloNovo
            If ResumoFinanceiro.ValorTitulosEmProvisao >= pValor Then
                titcomp = CriaTitulo("COMP", pValor)
                TitulosNew.Add(titcomp)
                TitulosNotaModificada.Add(titcomp)

                vlrComp = pValor
                pValor = 0
            Else
                titcomp = CriaTitulo("COMP", ResumoFinanceiro.ValorTitulosEmProvisao)
                TitulosNew.Add(titcomp)
                TitulosNotaModificada.Add(titcomp)

                pValor -= ResumoFinanceiro.ValorTitulosEmProvisao
                vlrComp += ResumoFinanceiro.ValorTitulosEmProvisao
            End If
        End If

        'Tem q transformar prov em compensacao vlrcomp
        For Each titprov In TitulosOld.Where(Function(S) S.CodigoSituacao = 1 And S.CodigoProvisao = 3).OrderBy(Function(s) s.Reprogramacao).ThenBy(Function(s) s.Codigo)
            If titprov.Valores.EncargoValorDocumento.Valor >= vlrComp Then
                titprov.Valores.EncargoValorDocumento.Valor = vlrComp
                titprov.IUD = "U"
                titprov.CodigoProvisao = eProvisao.Compensado
                TitulosNew.Add(titprov)
                Exit For
            Else
                vlrComp -= titprov.Valores.EncargoValorDocumento.Valor
                titprov.IUD = "U"
                titprov.CodigoProvisao = eProvisao.Compensado
                TitulosNew.Add(titprov)
            End If
        Next

        'criar adiantamento
        If pValor > 0 Then
            Dim titadi As Novo.TituloNovo = CriaTitulo("ADI", pValor)
            TitulosNew.Add(titadi)
            TitulosNotaModificada.Add(titadi)
        End If

    End Sub

    ' - N ok
    Private Sub DiminuiuValorNota(ByRef pValor As Decimal)
        For Each tit In NF_Original.Titulos
            TitulosNotaModificada.Add(tit)
        Next

        '*******************************************************************************
        '********** Se tiver provisao na nota  volta pra previsao **********************
        '*******************************************************************************
        For Each titProvisao In NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Provisao)
            'pega o mesmo titulo na lista antiga para altera-lo e manter a lista da nota original sem mudanças
            Dim num As Integer = titProvisao.Codigo
            Dim titProvOld As Novo.TituloNovo = TitulosOld.Where(Function(s) s.Codigo = num).FirstOrDefault

            If titProvisao.Valores.EncargoValorDocumento.Valor > pValor Then
                'Provisão
                titProvOld.IUD = "U"
                titProvOld.Valores.EncargoValorDocumento.Valor -= pValor
                TitulosNew.Add(titProvOld)
                'Localiza o titulo q sofreu alteracao na nova lista de titulos e substitui pela alterada
                Dim i As Integer = TitulosNotaModificada.FindIndex(Function(s) s.Codigo = num)
                TitulosNotaModificada(i) = titProvOld

                'previsão.
                Dim titPrevisao As Novo.TituloNovo
                titPrevisao = TitulosNew.Find(Function(s) s.Codigo = titProvOld.CodigoTituloOrigem)
                If titPrevisao Is Nothing Then titPrevisao = New Novo.TituloNovo(titProvOld.CodigoTituloOrigem)
                titPrevisao.IUD = "U"
                titPrevisao.Valores.EncargoValorDocumento.Valor = IIf(titPrevisao.CodigoSituacao = eSituacao.Normal, titPrevisao.Valores.EncargoValorDocumento.Valor + pValor, pValor)
                titPrevisao.CodigoSituacao = eSituacao.Normal
                TitulosNew.Add(titPrevisao)

                pValor = 0
            Else
                'Provisão
                titProvOld.IUD = "C"
                TitulosNew.Add(titProvOld)

                'Remove da lista de titulos originais
                TitulosOld.Remove(titProvOld)
                'Remove da nova lista de titulos da nota
                TitulosNotaModificada.RemoveAll(Function(s) s.Codigo = titProvOld.Codigo)

                'Previsao
                Dim titPrevisao As Novo.TituloNovo
                titPrevisao = TitulosNew.Find(Function(s) s.Codigo = titProvOld.CodigoTituloOrigem)
                If titPrevisao Is Nothing Then titPrevisao = New Novo.TituloNovo(titProvOld.CodigoTituloOrigem)
                titPrevisao.IUD = "U"
                titPrevisao.Valores.EncargoValorDocumento.Valor = IIf(titPrevisao.CodigoSituacao = eSituacao.Normal, titPrevisao.Valores.EncargoValorDocumento.Valor + titProvisao.Valores.EncargoValorDocumento.Valor, titProvisao.Valores.EncargoValorDocumento.Valor)
                titPrevisao.CodigoSituacao = eSituacao.Normal
                TitulosNew.Add(titPrevisao)
                pValor -= titProvisao.Valores.EncargoValorDocumento.Valor
            End If
            If pValor = 0 Then Exit Sub
        Next


        '***************************************************************************************************************************************************************
        '***************************************************************************************************************************************************************

        For Each titBaixa In NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isBaixaAdiantamento)
            'pega o mesmo titulo na lista antiga para altera-lo e manter a lista da nota original sem mudanças
            Dim num As Integer = titBaixa.Codigo
            Dim titBaixaOld As Novo.TituloNovo = TitulosOld.Where(Function(s) s.Codigo = num).FirstOrDefault

            If titBaixaOld.Valores.EncargoValorDocumento.Valor > pValor Then
                titBaixaOld.IUD = "U"
                titBaixaOld.Valores.EncargoValorDocumento.Valor -= pValor
                For Each baixa In titBaixaOld.Baixas_AdiantamentoEfetuadas
                    If pValor >= baixa.Valor Then
                        baixa.IUD = "D"
                        pValor -= baixa.Valor
                    Else
                        baixa.IUD = "U"
                        baixa.Valor -= pValor
                        pValor = 0
                    End If
                Next
                'Localiza o titulo q sofreu alteracao na nova lista de titulos e substitui pela alterada
                Dim i As Integer = TitulosNotaModificada.FindIndex(Function(s) s.Codigo = titBaixaOld.Codigo)
                TitulosNotaModificada(i) = titBaixaOld
            Else
                titBaixaOld.IUD = "D"
                titBaixaOld.NotaTitulo.IUD = "D"
                'Remove da nova lista de titulos da nota
                TitulosNotaModificada.RemoveAll(Function(s) s.Codigo = titBaixaOld.Codigo)
                pValor -= titBaixaOld.Valores.EncargoValorDocumento.Valor
            End If
            TitulosNew.Add(titBaixaOld)
        Next

        If pValor = 0 Then Exit Sub
        '***************************************************************************************************************************************************************
        '***************************************************************************************************************************************************************
        Dim VlrNFCompensado As Decimal = NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Compensado).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
        Dim VlrPedProvisao As Decimal = TitulosOld.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Provisao).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)

        If VlrNFCompensado > 0 Then
            Dim valorCompProvisao As Decimal = 0

            For Each titCompensacao As Novo.TituloNovo In NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Compensado)
                'pega o mesmo titulo na lista antiga para altera-lo e manter a lista da nota original sem mudanças
                Dim num As Integer = titCompensacao.Codigo
                Dim titComOld As Novo.TituloNovo = TitulosOld.Where(Function(s) s.Codigo = num).FirstOrDefault

                'Retira o valor das compensações da nota.
                If titComOld.Valores.EncargoValorDocumento.Valor > pValor Then
                    titComOld.IUD = "U"
                    titComOld.Valores.EncargoValorDocumento.Valor -= pValor
                    valorCompProvisao = pValor

                    'Localiza o titulo q sofreu alteracao na nova lista de titulos e substitui pela alterada
                    Dim i As Integer = TitulosNotaModificada.FindIndex(Function(s) s.Codigo = titComOld.Codigo)
                    TitulosNotaModificada(i) = titComOld
                Else
                    titComOld.IUD = "D"
                    titComOld.NotaTitulo.IUD = "D"
                    valorCompProvisao += titComOld.Valores.EncargoValorDocumento.Valor
                    TitulosNotaModificada.RemoveAll(Function(s) s.Codigo = titComOld.Codigo)
                End If
                TitulosNew.Add(titComOld)


                'Verifica se há títulos de provisão nas outras notas e compensa o valor retirado das compensações da NF Modificada
                For Each titProvisao In TitulosOld.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao AndAlso s.CodigoSituacao = eSituacao.Normal)
                    If titProvisao.CodigoProvisao <> eProvisao.Compensado Then
                        If titProvisao.Valores.EncargoValorDocumento.Valor > valorCompProvisao Then
                            'Provisão
                            titProvisao.IUD = "U"
                            titProvisao.Valores.EncargoValorDocumento.Valor -= valorCompProvisao
                            TitulosNew.Add(titProvisao)
                            'Compensação
                            TitulosNew.Add(CriaTitulo("COMP", valorCompProvisao, Nothing, titProvisao.NotaTitulo.NotaFiscal, titCompensacao.CodigoTituloCompensacao))

                            pValor -= valorCompProvisao
                            valorCompProvisao = 0
                        Else
                            titProvisao.IUD = "U"
                            titProvisao.CodigoProvisao = eProvisao.Compensado
                            titProvisao.CodigoTituloCompensacao = titCompensacao.CodigoTituloCompensacao
                            TitulosNew.Add(titProvisao)

                            pValor -= titProvisao.Valores.EncargoValorDocumento.Valor
                            valorCompProvisao -= titProvisao.Valores.EncargoValorDocumento.Valor
                        End If
                    End If
                    If valorCompProvisao = 0 Then Exit For
                Next

                If valorCompProvisao > 0 Then
                    If titComOld.TituloCompensacao.Valores.EncargoValorDocumento.Valor = valorCompProvisao Then
                        titComOld.TituloCompensacao.IUD = "D"
                        titComOld.TituloCompensacao.NotaTitulo.IUD = "D"
                    Else
                        titComOld.TituloCompensacao.IUD = "U"
                        titComOld.TituloCompensacao.Valores.EncargoValorDocumento.Valor -= valorCompProvisao
                    End If


                    TitulosNew.Add(titComOld.TituloCompensacao)
                    TitulosNew.Add(CriaTitulo("ADI", valorCompProvisao, Nothing, titComOld.TituloCompensacao.NotaTitulo.NotaFiscal))
                    pValor -= valorCompProvisao
                End If
            Next
        End If

        If pValor = 0 Then Exit Sub


        'Titulos Baixados devem ser desmembrados e gerar adiantamento
        For Each titBaixado As Novo.TituloNovo In NF_Original.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Baixa AndAlso s.CodigoSituacao = eSituacao.Normal AndAlso Not s.isAdiantamento AndAlso Not s.isBaixaAdiantamento)
            'pega o mesmo titulo na lista antiga para altera-lo e manter a lista da nota original sem mudanças
            Dim num As Integer = titBaixado.Codigo
            Dim titBaixadoOld As Novo.TituloNovo = TitulosOld.Where(Function(s) s.Codigo = num).FirstOrDefault
            'Localiza o titulo q sofreu alteracao na nova lista de titulos e substitui pela alterada
            Dim i As Integer = TitulosNotaModificada.FindIndex(Function(s) s.Codigo = titBaixadoOld.Codigo)

            If titBaixadoOld.Valores.EncargoValorDocumento.Valor > pValor Then
                titBaixadoOld.IUD = "U"
                titBaixadoOld.Valores.EncargoValorDocumento.Valor -= pValor
                TitulosNew.Add(titBaixadoOld)
                TitulosNotaModificada(i) = titBaixadoOld

                Dim titAdiantamento As New Novo.TituloNovo(titBaixadoOld.Codigo)
                titAdiantamento.IUD = "I"
                titAdiantamento.Valores.EncargoValorDocumento.Valor = pValor
                titAdiantamento.CodigoContaContabilCliFor = Pedido.SubOperacao.CodigoContaAdiantamento
                titAdiantamento.Historico = " Adiantamento criado pelo desmembramento do título: " & titBaixado.Codigo & " devido a alteração da NF: " & NF_Original.Codigo &
                                            " Série " & NF_Original.Serie & " no valor de " & NF_Original.TotalNota.ToString("N2") & " para " & NF_Modificada.TotalNota.ToString("N2")
                titAdiantamento.NotaTitulo = Nothing

                pValor = 0
                TitulosNew.Add(titAdiantamento)
                TitulosNotaModificada.Add(titAdiantamento)
            Else
                titBaixadoOld.IUD = "D"
                'titBaixadoOld.CodigoContaContabilCliFor = Pedido.SubOperacao.CodigoContaAdiantamento
                'titBaixadoOld.Historico = " Adiantamento criado pelo desmembramento do título: " & titBaixadoOld.Codigo & " devido a alteração da NF: " & NF_Original.Codigo & _
                '                         " Série " & NF_Original.Serie & " no valor de " & NF_Original.TotalNota.ToString("N2") & " para " & NF_Modificada.TotalNota.ToString("N2")
                titBaixadoOld.NotaTitulo.IUD = "D"
                TitulosNew.Add(titBaixadoOld)
                TitulosNotaModificada.RemoveAll(Function(s) s.Codigo = titBaixadoOld.Codigo)

                'Adiantamento
                Dim titAdiantamento As New Novo.TituloNovo(titBaixadoOld.Codigo)
                titAdiantamento.IUD = "I"
                titAdiantamento.Valores.EncargoValorDocumento.Valor = titBaixadoOld.Valores.EncargoValorDocumento.Valor
                titAdiantamento.CodigoContaContabilCliFor = Pedido.SubOperacao.CodigoContaAdiantamento
                titAdiantamento.Historico = " Adiantamento criado pelo desmembramento do título: " & titBaixado.Codigo & " devido a alteração da NF: " & NF_Original.Codigo &
                                            " Série " & NF_Original.Serie & " no valor de " & NF_Original.TotalNota.ToString("N2") & " para " & NF_Modificada.TotalNota.ToString("N2")
                titAdiantamento.NotaTitulo = Nothing
                TitulosNew.Add(titAdiantamento)

                pValor -= titBaixadoOld.Valores.EncargoValorDocumento.Valor
            End If
        Next
    End Sub

    ' - S
    Private Sub DiminuiuValorNotaDevolucao(ByRef pValor As Decimal)
        For Each tit In NF_Original.Titulos
            TitulosNotaModificada.Add(tit)
        Next

        'Apura o valor compensado em titulos da nota
        Dim VlrNFCompensado As Decimal = NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Compensado).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)

        'Se tem titulos compensados diminui a compensacao
        If VlrNFCompensado > 0 Then
            For Each titComp In NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Compensado)
                'Pega o mesmo titulo na lista antiga para altera-lo e manter a lista da nota original sem mudanças
                Dim num As Integer = titComp.Codigo
                Dim titCompOld As Novo.TituloNovo = TitulosOld.Where(Function(s) s.Codigo = num).FirstOrDefault

                Dim vlrACompensar As Decimal = 0
                If titCompOld.Valores.EncargoValorDocumento.Valor <= pValor Then
                    titCompOld.IUD = "C"
                    vlrACompensar = titCompOld.Valores.EncargoValorDocumento.Valor
                    pValor -= titCompOld.Valores.EncargoValorDocumento.Valor

                    TitulosNotaModificada.RemoveAll(Function(s) s.Codigo = titCompOld.Codigo)
                Else
                    titCompOld.IUD = "U"
                    titCompOld.Valores.EncargoValorDocumento.Valor -= pValor
                    vlrACompensar = pValor
                    pValor = 0

                    'Localiza o titulo q sofreu alteracao na nova lista de titulos e substitui pela alterada
                    Dim i As Integer = TitulosNotaModificada.FindIndex(Function(s) s.Codigo = titCompOld.Codigo)
                    TitulosNotaModificada(i) = titCompOld
                End If
                TitulosNew.Add(titCompOld)



                'Vai nos titulos em q a devolucao transformou provisao em compensacao e diminui a compensacao e cria uma provisao
                For Each titcompNf In TitulosOld.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Compensado And s.CodigoTituloCompensacao = num)
                    If titcompNf.Valores.EncargoValorDocumento.Valor = vlrACompensar Then
                        titcompNf.IUD = "U"
                        titcompNf.CodigoProvisao = eProvisao.Provisao
                        titcompNf.CodigoTituloCompensacao = 0
                        titcompNf.TituloCompensacao = Nothing
                        TitulosNew.Add(titcompNf)
                        Exit For

                    ElseIf titcompNf.Valores.EncargoValorDocumento.Valor > vlrACompensar Then
                        Dim titprov As New Novo.TituloNovo(titcompNf.Codigo)
                        titprov.IUD = "I"
                        titprov.CodigoTituloCompensacao = 0
                        titprov.TituloCompensacao = Nothing
                        titprov.Valores.EncargoValorDocumento.Valor = vlrACompensar
                        titprov.CodigoProvisao = eProvisao.Provisao
                        TitulosNew.Add(titprov)

                        titcompNf.IUD = "U"
                        titcompNf.Valores.EncargoValorDocumento.Valor -= vlrACompensar
                        TitulosNew.Add(titcompNf)
                        Exit For
                    Else
                        titcompNf.IUD = "U"
                        titcompNf.CodigoTituloCompensacao = 0
                        titcompNf.TituloCompensacao = Nothing
                        titcompNf.CodigoProvisao = eProvisao.Provisao
                        TitulosNew.Add(titcompNf)
                        vlrACompensar -= titcompNf.Valores.EncargoValorDocumento.Valor
                    End If
                Next

                If pValor = 0 Then Exit Sub
            Next
        End If


        Dim VlrNFSaldoAdiantamento As Decimal = NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isAdiantamento).Sum(Function(s) s.Adiantamento.SaldoValor)
        If VlrNFSaldoAdiantamento > pValor Then
            For Each titAdi In NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isAdiantamento)
                If titAdi.Adiantamento.SaldoValor = pValor Then
                    titAdi.IUD = IIf(titAdi.Valores.EncargoValorDocumento.Valor = pValor, "C", "U")
                    titAdi.Valores.EncargoValorDocumento.Valor -= IIf(titAdi.IUD = "C", 0, pValor)
                    TitulosNew.Add(titAdi)

                    If titAdi.IUD = "U" Then TitulosNotaModificada.Add(titAdi)

                    Exit Sub

                ElseIf titAdi.Adiantamento.SaldoValor > pValor Then
                    titAdi.IUD = "U"
                    titAdi.Valores.EncargoValorDocumento.Valor -= pValor
                    TitulosNew.Add(titAdi)

                    TitulosNotaModificada.Add(titAdi)

                    Exit Sub

                Else
                    titAdi.IUD = "U"
                    titAdi.Valores.EncargoValorDocumento.Valor -= titAdi.Adiantamento.SaldoValor
                    pValor -= titAdi.Adiantamento.SaldoValor
                    TitulosNew.Add(titAdi)

                    TitulosNotaModificada.Add(titAdi)
                End If
            Next
        Else
            Dim CodigoAdi As Integer = 0
            For Each titAdi In NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isAdiantamento)
                titAdi.IUD = "C"
                TitulosNew.Add(titAdi)
                CodigoAdi = titAdi.Codigo
                TitulosNotaModificada.RemoveAll(Function(s) s.Codigo = CodigoAdi)
            Next
        End If

        'lista com os adiantamentos do pedido - os adiantamentos da NF Alterada
        'Dim listAdi As Novo.ListTitulo = TitulosOld.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isAdiantamento AndAlso s.Adiantamento.SaldoValor > 0)
        TitulosOld.RemoveAll(Function(s) s.CodigoSituacao <> eSituacao.Normal OrElse s.CodigoProvisao <> eProvisao.Baixa OrElse Not s.isAdiantamento)
        Dim Codigo As Integer = 0
        For Each titNota In NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isAdiantamento)
            'listAdi.Remove(titNota)
            Codigo = titNota.Codigo
            TitulosOld.RemoveAll(Function(s) s.Codigo = Codigo)
        Next

        '********************************************************* 2
        For Each Baixa In NF_Original.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isAdiantamento).SelectMany(Function(s) s.Adiantamento.Baixas)
            Dim TituloDaBaixa As Novo.TituloNovo = Baixa.TituloBaixa
            Dim num As Integer = Baixa.CodigoTituloAdiantamento
            Dim seq As Integer = Baixa.Sequencia
            For Each AdiaPedido In TitulosOld.Where(Function(s) s.Adiantamento.SaldoValor > 0)
                Dim bx As Novo.AdiantamentoBaixaNovo = TituloDaBaixa.Baixas_AdiantamentoEfetuadas.Where(Function(s) s.CodigoTituloAdiantamento = num And s.Sequencia = seq).FirstOrDefault
                Dim NovaBx As New Novo.AdiantamentoBaixaNovo(AdiaPedido.Adiantamento)
                Dim saldo As Decimal = IIf(AdiaPedido.Adiantamento.SaldoValor > pValor, pValor, AdiaPedido.Adiantamento.SaldoValor)

                If bx.Valor > saldo Then
                    bx.IUD = "U"
                    bx.Valor -= saldo

                    NovaBx.IUD = "I"
                    NovaBx.CodigoTituloBaixa = TituloDaBaixa.Codigo
                    NovaBx.IndiceBaixa = TituloDaBaixa.IndiceTitulo
                    NovaBx.Movimento = DateTime.Now
                    NovaBx.Valor = saldo
                    NovaBx.TipoLancamento = "B"
                    TituloDaBaixa.Baixas_AdiantamentoEfetuadas.Add(NovaBx)

                    pValor = 0
                ElseIf bx.Valor <= saldo Then
                    bx.IUD = "D"
                    pValor -= bx.Valor

                    NovaBx.IUD = "I"
                    NovaBx.CodigoTituloBaixa = TituloDaBaixa.Codigo
                    NovaBx.IndiceBaixa = TituloDaBaixa.IndiceTitulo
                    NovaBx.Movimento = DateTime.Now
                    NovaBx.Valor = bx.Valor
                    NovaBx.TipoLancamento = "B"
                    TituloDaBaixa.Baixas_AdiantamentoEfetuadas.Add(NovaBx)
                End If

                AdiaPedido.Baixas_AdiantamentoEfetuadas.Add(NovaBx)
                TitulosNew.Add(TituloDaBaixa)
            Next

            If TitulosOld.Sum(Function(s) s.Adiantamento.SaldoValor) = 0 And pValor > 0 Then
                'transforma em provisao
                'ver se ja nao esta na titulos new
                Dim bx As Novo.AdiantamentoBaixaNovo = TituloDaBaixa.Baixas_AdiantamentoEfetuadas.Where(Function(s) s.IUD = "U")
                bx.IUD = "D"
                TituloDaBaixa.Valores.EncargoValorDocumento.Valor -= bx.Valor
                TitulosNew.Add(CriaTitulo("PRO", pValor, Nothing, Nothing, Baixa.TituloAdiantamento.CodigoTituloOrigem))
            End If
        Next
    End Sub

    ' + N de Entrada
    Public Sub AumentouValorNotaDeEntrada(ByRef pValor As Decimal)
        '*******************************************************************************
        '********** Se tiver previsão na nota aumenta o valor do titulo ****************
        '*******************************************************************************
        Dim qtdeTitPrevisao = _VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Previsao).Count()

        For Each titPrevisao In _VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Previsao)

            Dim valorPorParcela As Decimal = IIf(qtdeTitPrevisao > 1, pValor / qtdeTitPrevisao, pValor)

            'Previsão
            titPrevisao.IUD = IIf(_NF_Modificada.IUD = "I", "I", "U")
            titPrevisao.ValorDoDocumento += valorPorParcela

            'provisão.
            Dim titProvisao As Titulo

            If _VencimentosPedido.Any(Function(t) t.CodigoProvisao = eProvisao.Provisao) Then
                titProvisao = _VencimentosPedido.First(Function(t) t.CodigoProvisao = eProvisao.Provisao)
                If titProvisao.CodigoSituacao = eSituacao.Excluido OrElse titProvisao.ValorDoDocumento < 0 Then
                    titProvisao.ValorDoDocumento = 0
                End If
                titProvisao.IUD = "U"
            Else
                titProvisao = New Titulo(titPrevisao.Codigo)
                Dim ListHistoricoFinan = New ListHistoricoFinanceiro(NF_Original)
                titProvisao.Codigo = ListHistoricoFinan.First(Function(t) t.CodigoProvisao = eProvisao.Provisao).Titulo
                titProvisao.ValorDoDocumento = 0
                titProvisao.IUD = "I"
            End If

            titProvisao.ValorDoDocumento -= valorPorParcela
            titProvisao.CodigoSituacao = eSituacao.Normal

            'Historico Financeiro Previsão.
            titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, _NF_Modificada, titProvisao.Codigo, valorPorParcela, "Adicionado o valor de R$: " & valorPorParcela.ToString("C2", CultureInfo.CurrentCulture) & ", ref. ao aumento do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

            'Historico Financeiro Provisão.
            titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, 0, valorPorParcela, "Retirado o valor de R$: " & valorPorParcela.ToString("C2", CultureInfo.CurrentCulture) & ", ref. ao aumento do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

            pValor -= valorPorParcela
            qtdeTitPrevisao = qtdeTitPrevisao - 1

            If pValor = 0 Then
                _NF_Modificada.VencimentosNota = _VencimentosNota
                _NF_Modificada.VencimentosPedido = _VencimentosPedido
                _NF_Modificada.VencimentosNota.NF = _NF_Modificada
                Exit Sub
            End If
        Next
    End Sub

    ' - N de Entrada 
    Public Sub DiminuiuValorNotaDeEntrada(ByRef pValor As Decimal)
        '*******************************************************************************
        '********** Se tiver previsão na nota volta pra provisão  **********************
        '*******************************************************************************
        For Each titPrevisao In _VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Previsao)
            Dim num As Integer = titPrevisao.Codigo

            If titPrevisao.ValorDoDocumento > pValor Then
                'Previsão
                titPrevisao.IUD = IIf(_NF_Modificada.IUD = "I", "I", "U")
                titPrevisao.ValorDoDocumento -= pValor
                'provisão.
                Dim titProvisao As Titulo

                If _VencimentosPedido.Any(Function(t) t.CodigoProvisao = eProvisao.Provisao) Then
                    titProvisao = _VencimentosPedido.First(Function(t) t.CodigoProvisao = eProvisao.Provisao)
                    If titProvisao.CodigoSituacao = eSituacao.Excluido OrElse titProvisao.ValorDoDocumento < 0 Then
                        titProvisao.ValorDoDocumento = 0
                    End If
                    titProvisao.IUD = "U"
                Else
                    titProvisao = New Titulo(titPrevisao.Codigo)
                    Dim ListHistoricoFinan = New ListHistoricoFinanceiro(NF_Original)
                    titProvisao.Codigo = ListHistoricoFinan.First(Function(t) t.CodigoProvisao = eProvisao.Provisao).Titulo
                    titProvisao.ValorDoDocumento = 0
                    titProvisao.CodigoProvisao = eProvisao.Provisao
                    titProvisao.CodigoSituacao = eSituacao.Normal
                    titProvisao.IUD = "I"
                    titProvisao.Historico = "Valor: " & pValor & " estornado do título " & titPrevisao.Codigo
                    _VencimentosPedido.Add(titProvisao)

                End If

                titProvisao.ValorDoDocumento += pValor
                titPrevisao.CodigoSituacao = eSituacao.Normal

                'Historico Financeiro Previsão.
                titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, _NF_Modificada, 0, pValor, "Retirado o valor de R$: " & pValor.ToString("C2", CultureInfo.CurrentCulture) & ", ref. a diminuição do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                'Historico Financeiro PROVISÃO.
                titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, titPrevisao.Codigo, pValor, "Adicionado o valor de R$: " & pValor.ToString("C2", CultureInfo.CurrentCulture) & ", ref. a diminuição do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                pValor = 0
            Else
                'previsão
                titPrevisao.CodigoProvisao = eProvisao.Provisao
                titPrevisao.IUD = "U"

                'provisão.
                Dim titProvisao As Titulo

                If _VencimentosPedido.Any(Function(t) t.CodigoProvisao = eProvisao.Provisao) Then
                    titProvisao = _VencimentosPedido.First(Function(t) t.CodigoProvisao = eProvisao.Provisao)
                    titProvisao.IUD = "U"
                Else
                    titProvisao = New Titulo(titPrevisao.Codigo)
                    titProvisao.ValorDoDocumento = 0
                    titProvisao.CodigoProvisao = eProvisao.Provisao
                    titProvisao.CodigoSituacao = eSituacao.Normal
                    titProvisao.IUD = "I"
                    titProvisao.Historico = "Valor: " & pValor & " estornado do título " & titPrevisao.Codigo
                    _VencimentosPedido.Add(titProvisao)
                End If

                titProvisao.ValorDoDocumento += titPrevisao.ValorDoDocumento
                titPrevisao.CodigoSituacao = eSituacao.Normal

                'Historico Financeiro.
                titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, titPrevisao.Codigo, titPrevisao.ValorDoDocumento))

                pValor -= titPrevisao.ValorDoDocumento

                'Historico Financeiro Previsão.
                titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, _NF_Modificada, 0, titPrevisao.ValorDoDocumento, "Retirado o valor de R$: " & titPrevisao.ValorDoDocumento.ToString("C2", CultureInfo.CurrentCulture) & ", ref. a diminuição do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                'Historico Financeiro PROVISÃO.
                titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, titPrevisao.Codigo, titPrevisao.ValorDoDocumento, "Adicionado o valor de R$: " & titPrevisao.ValorDoDocumento.ToString("C2", CultureInfo.CurrentCulture) & ", ref. a diminuição do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))
            End If

            If pValor = 0 Then
                _NF_Modificada.VencimentosNota = _VencimentosNota
                _NF_Modificada.VencimentosPedido = _VencimentosPedido
                _NF_Modificada.VencimentosNota.NF = _NF_Modificada
                Exit Sub
            End If
        Next
    End Sub

    Public Sub AumentouValorNotaDeSaida(ByRef pValor As Decimal)
        '*******************************************************************************
        '********** Se tiver previsão na nota aumenta o valor do titulo ****************
        '*******************************************************************************
        Dim qtdeTitPrevisao = _VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Previsao).Count()

        For Each titPrevisao In _VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Previsao)

            Dim valorPorParcela As Decimal = IIf(qtdeTitPrevisao > 1, pValor / qtdeTitPrevisao, pValor)

            'Previsão
            titPrevisao.IUD = IIf(_NF_Modificada.IUD = "I", "I", "U")
            titPrevisao.ValorDoDocumento += valorPorParcela

            'provisão.
            Dim titProvisao As Titulo

            If _VencimentosPedido.Any(Function(t) t.CodigoProvisao = eProvisao.Provisao) Then
                titProvisao = _VencimentosPedido.First(Function(t) t.CodigoProvisao = eProvisao.Provisao)
                If titProvisao.CodigoSituacao = eSituacao.Excluido OrElse titProvisao.ValorDoDocumento < 0 Then
                    titProvisao.ValorDoDocumento = 0
                End If
                titProvisao.IUD = "U"
            Else
                titProvisao = New Titulo(titPrevisao.Codigo)
                Dim ListHistoricoFinan = New ListHistoricoFinanceiro(NF_Original)
                titProvisao.Codigo = ListHistoricoFinan.First(Function(t) t.CodigoProvisao = eProvisao.Provisao).Titulo
                titProvisao.ValorDoDocumento = 0
                titProvisao.IUD = "I"
            End If

            titProvisao.ValorDoDocumento -= valorPorParcela
            titProvisao.CodigoSituacao = eSituacao.Normal

            'Historico Financeiro Previsão.
            titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, _NF_Modificada, titProvisao.Codigo, valorPorParcela, "Adicionado o valor de R$: " & valorPorParcela.ToString("C2", CultureInfo.CurrentCulture) & ", ref. ao aumento do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

            'Historico Financeiro Provisão.
            titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, 0, valorPorParcela, "Retirado o valor de R$: " & valorPorParcela.ToString("C2", CultureInfo.CurrentCulture) & ", ref. ao aumento do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

            pValor -= valorPorParcela
            qtdeTitPrevisao = qtdeTitPrevisao - 1

            If pValor = 0 Then
                _NF_Modificada.VencimentosNota = _VencimentosNota
                _NF_Modificada.VencimentosPedido = _VencimentosPedido
                _NF_Modificada.VencimentosNota.NF = _NF_Modificada
                Exit Sub
            End If
        Next
    End Sub

    Public Sub DiminuiuValorNotaDeSaida(ByRef pValor As Decimal)
        '*******************************************************************************
        '********** Se tiver previsão na nota volta pra provisão  **********************
        '*******************************************************************************
        For Each titPrevisao In _VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Previsao)
            Dim num As Integer = titPrevisao.Codigo

            If titPrevisao.ValorDoDocumento > pValor Then
                'Previsão
                titPrevisao.IUD = IIf(_NF_Modificada.IUD = "I", "I", "U")
                titPrevisao.ValorDoDocumento -= pValor
                'provisão.
                Dim titProvisao As Titulo

                If _VencimentosPedido.Any(Function(t) t.CodigoProvisao = eProvisao.Provisao) Then
                    titProvisao = _VencimentosPedido.First(Function(t) t.CodigoProvisao = eProvisao.Provisao)
                    If titProvisao.CodigoSituacao = eSituacao.Excluido OrElse titProvisao.ValorDoDocumento < 0 Then
                        titProvisao.ValorDoDocumento = 0
                    End If
                    titProvisao.IUD = "U"
                Else
                    titProvisao = New Titulo(titPrevisao.Codigo)
                    Dim ListHistoricoFinan = New ListHistoricoFinanceiro(NF_Original)
                    titProvisao.Codigo = ListHistoricoFinan.First(Function(t) t.CodigoProvisao = eProvisao.Provisao).Titulo
                    titProvisao.ValorDoDocumento = 0
                    titProvisao.CodigoProvisao = eProvisao.Provisao
                    titProvisao.CodigoSituacao = eSituacao.Normal
                    titProvisao.IUD = "I"
                    titProvisao.Historico = "Valor: " & pValor & " estornado do título " & titPrevisao.Codigo
                    _VencimentosPedido.Add(titProvisao)

                End If

                titProvisao.ValorDoDocumento += pValor
                titPrevisao.CodigoSituacao = eSituacao.Normal

                'Historico Financeiro Previsão.
                titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, _NF_Modificada, 0, pValor, "Retirado o valor de R$: " & pValor.ToString("C2", CultureInfo.CurrentCulture) & ", ref. a diminuição do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                'Historico Financeiro PROVISÃO.
                titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, titPrevisao.Codigo, pValor, "Adicionado o valor de R$: " & pValor.ToString("C2", CultureInfo.CurrentCulture) & ", ref. a diminuição do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                pValor = 0
            Else
                'previsão
                titPrevisao.CodigoProvisao = eProvisao.Provisao
                titPrevisao.IUD = "U"

                'provisão.
                Dim titProvisao As Titulo

                If _VencimentosPedido.Any(Function(t) t.CodigoProvisao = eProvisao.Provisao) Then
                    titProvisao = _VencimentosPedido.First(Function(t) t.CodigoProvisao = eProvisao.Provisao)
                    titProvisao.IUD = "U"
                Else
                    titProvisao = New Titulo(titPrevisao.Codigo)
                    titProvisao.ValorDoDocumento = 0
                    titProvisao.CodigoProvisao = eProvisao.Provisao
                    titProvisao.CodigoSituacao = eSituacao.Normal
                    titProvisao.IUD = "I"
                    titProvisao.Historico = "Valor: " & pValor & " estornado do título " & titPrevisao.Codigo
                    _VencimentosPedido.Add(titProvisao)
                End If

                titProvisao.ValorDoDocumento += titPrevisao.ValorDoDocumento
                titPrevisao.CodigoSituacao = eSituacao.Normal

                'Historico Financeiro.
                titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, titPrevisao.Codigo, titPrevisao.ValorDoDocumento))

                pValor -= titPrevisao.ValorDoDocumento

                'Historico Financeiro Previsão.
                titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, _NF_Modificada, 0, titPrevisao.ValorDoDocumento, "Retirado o valor de R$: " & titPrevisao.ValorDoDocumento.ToString("C2", CultureInfo.CurrentCulture) & ", ref. a diminuição do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                'Historico Financeiro PROVISÃO.
                titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, titPrevisao.Codigo, titPrevisao.ValorDoDocumento, "Adicionado o valor de R$: " & titPrevisao.ValorDoDocumento.ToString("C2", CultureInfo.CurrentCulture) & ", ref. a diminuição do valor da NF: " & _NF_Modificada.Codigo & "-" & _NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))
            End If

            If pValor = 0 Then
                _NF_Modificada.VencimentosNota = _VencimentosNota
                _NF_Modificada.VencimentosPedido = _VencimentosPedido
                _NF_Modificada.VencimentosNota.NF = _NF_Modificada
                Exit Sub
            End If
        Next
    End Sub

    'Nota de Devolução 
    Public Sub EstornoNotaDeDevolucao(ByRef pValor As Decimal, ByRef pValorBruto As Decimal)

        Dim indiceDiferencaLiquido As Decimal
        Dim indiceTitulo As Decimal

        indiceDiferencaLiquido = pValor / pValorBruto

        '*******************************************************************************
        '********** Notas associadas a NF de Devolução *********************************
        '*******************************************************************************
        Dim titProvisao = _VencimentosPedido.Find(Function(x) x.CodigoProvisao = eProvisao.Provisao)
        Dim listTitulosExtornoDeDevolucao As New Negocio.ListTitulo

        If _NF_Modificada.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then

            If _NF_Modificada.Pedido.CodigoIndexador = 99 OrElse _NF_Modificada.Pedido.IndexadorFixo Then

                indiceTitulo = _NF_Modificada.Pedido.IndiceFixado

            Else

                Dim indiceDolar As Decimal = New Cotacao(_NF_Modificada.Pedido.CodigoIndexador, _NF_Modificada.DataNota).Indice 'PTAX
                indiceTitulo = indiceDolar

            End If

        End If

        If Not titProvisao Is Nothing Then
            titProvisao.IndiceTitulo = indiceTitulo
        End If

        For Each nf As NotaFiscalDevolucaoXNotaFiscal In _NF_Devolucoes

            nf.ValorLiquidoDevolucao = Math.Round(nf.ValorDevolucao * indiceDiferencaLiquido, 2, MidpointRounding.AwayFromZero)
            Dim pValorDesconto As Decimal = Math.Round(nf.ValorDevolucao * indiceDiferencaLiquido, 2, MidpointRounding.AwayFromZero)

            '*******************************************************************************
            '********** Se tiver previsão na notas volta pra provisão  *********************
            '*******************************************************************************
            For Each titPrevisao In nf.Nota.VencimentosNota.Where(Function(x) x.CodigoProvisao = eProvisao.Previsao)

                titPrevisao.IUD = "U"
                titPrevisao.IndiceTitulo = indiceTitulo
                titPrevisao.Historico = titPrevisao.Historico

                If listTitulosExtornoDeDevolucao.Any(Function(x) x.Codigo = titPrevisao.Codigo) Then
                    Dim titPrevisaoExtornado As Titulo = listTitulosExtornoDeDevolucao.First(Function(x) x.Codigo = titPrevisao.Codigo)

                    titPrevisao.CodigoProvisao = titPrevisaoExtornado.CodigoProvisao
                    titPrevisao.ValorDoDocumento = titPrevisaoExtornado.ValorDoDocumento
                    titPrevisao.MoedaValorDoDocumento = titPrevisaoExtornado.MoedaValorDoDocumento
                    titPrevisao.Deducoes = titPrevisaoExtornado.Deducoes
                    titPrevisao.Descontos = titPrevisaoExtornado.Descontos
                    titPrevisao.Juros = titPrevisaoExtornado.Juros
                    titPrevisao.Acrescimos = titPrevisaoExtornado.Acrescimos
                    titPrevisao.MoedaValorDoDocumento = titPrevisaoExtornado.MoedaValorDoDocumento
                    titPrevisao.MoedaDeducoes = titPrevisaoExtornado.MoedaDeducoes
                    titPrevisao.MoedaDescontos = titPrevisaoExtornado.MoedaDescontos
                    titPrevisao.MoedaJuros = titPrevisaoExtornado.MoedaJuros
                    titPrevisao.MoedaAcrescimos = titPrevisaoExtornado.MoedaAcrescimos
                    titPrevisao.CodigoSituacao = titPrevisaoExtornado.CodigoSituacao

                    If titPrevisao.CodigoProvisao <> eProvisao.Previsao Then
                        'Se o titulo estiver em uma condição diferente de Previsao, saimos do for
                        Continue For
                    End If
                End If

                'Se a devolução for maior que o 
                If pValorDesconto >= titPrevisao.ValorDoDocumento Then

                    Dim valor = titPrevisao.ValorDoDocumento
                    Dim valorDolar = titPrevisao.MoedaValorDoDocumento
                    titPrevisao.CodigoProvisao = eProvisao.Previsao
                    titPrevisao.ValorDoDocumento = 0
                    titPrevisao.Deducoes = 0
                    titPrevisao.Descontos = 0
                    titPrevisao.Juros = 0
                    titPrevisao.Acrescimos = 0
                    titPrevisao.MoedaValorDoDocumento = 0
                    titPrevisao.MoedaDeducoes = 0
                    titPrevisao.MoedaDescontos = 0
                    titPrevisao.MoedaJuros = 0
                    titPrevisao.MoedaAcrescimos = 0
                    titPrevisao.CodigoSituacao = eSituacao.Excluido
                    'Historico Financeiro Previsão.
                    titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, _NF_Modificada, 0, valor, "Retirado o valor de R$ : " & valor.ToString("C2", CultureInfo.CurrentCulture) & ", ref a NF de devolução: " & NF_Modificada.Codigo & "-" & NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                    'Jogamos a diferença de saldo da nota de devolução em um titulo de Provisao
                    If titProvisao Is Nothing Then
                        titProvisao = New Titulo(titPrevisao.Codigo)
                        titProvisao.ListHistoricoFinanceiro.Clear()
                        Dim n As Numerador = New Numerador(1)
                        Dim NumeroTitulo As Integer = n.Sequencia + 1
                        titProvisao.Codigo = NumeroTitulo
                        titProvisao.CodigoProvisao = eProvisao.Provisao
                        titProvisao.ValorDoDocumento = 0
                        titProvisao.IUD = "I"
                    End If

                    'Adiciona na Provisão.
                    titProvisao.CodigoProvisao = eProvisao.Provisao
                    titProvisao.Deducoes = 0
                    titProvisao.Descontos = 0
                    titProvisao.Juros = 0
                    titProvisao.Acrescimos = 0
                    titProvisao.ValorDoDocumento += valor
                    titProvisao.MoedaDeducoes = 0
                    titProvisao.MoedaDescontos = 0
                    titProvisao.MoedaJuros = 0
                    titProvisao.MoedaAcrescimos = 0
                    titProvisao.MoedaValorDoDocumento += valorDolar
                    titProvisao.IUD = IIf(titProvisao.IUD = "I", "I", "U")
                    'Historico Financeiro Provisão.
                    titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, titPrevisao.Codigo, valor, "Adicionado o valor de R$ : " & valor.ToString("C2", CultureInfo.CurrentCulture) & ", ref a NF de devolução: " & NF_Modificada.Codigo & "-" & NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                    pValor -= valor

                Else

                    'Desconta o restante da previsão
                    titPrevisao.Deducoes = 0
                    titPrevisao.Descontos = 0
                    titPrevisao.Juros = 0
                    titPrevisao.Acrescimos = 0
                    titPrevisao.ValorDoDocumento -= pValorDesconto
                    'Historico Financeiro Previsão.
                    titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, _NF_Modificada, 0, pValorDesconto, "Retirado o valor de R$ : " & pValorDesconto.ToString("C2", CultureInfo.CurrentCulture) & ", ref a NF de devolução: " & NF_Modificada.Codigo & "-" & NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                    'Jogamos a diferença de saldo da nota de devolução em um titulo de Provisao
                    If titProvisao Is Nothing Then
                        titProvisao = New Titulo(titPrevisao.Codigo)
                        titProvisao.ListHistoricoFinanceiro.Clear()
                        Dim n As Numerador = New Numerador(1)
                        Dim NumeroTitulo As Integer = n.Sequencia + 1
                        titProvisao.Codigo = NumeroTitulo
                        titProvisao.CodigoProvisao = eProvisao.Provisao
                        titProvisao.ValorDoDocumento = 0
                        titProvisao.IUD = "I"
                    End If

                    'Adiciona na Provisão.
                    titProvisao.CodigoProvisao = eProvisao.Provisao
                    titProvisao.Deducoes = 0
                    titProvisao.Descontos = 0
                    titProvisao.Juros = 0
                    titProvisao.Acrescimos = 0
                    titProvisao.ValorDoDocumento += pValorDesconto
                    titProvisao.MoedaDeducoes = 0
                    titProvisao.MoedaDescontos = 0
                    titProvisao.MoedaJuros = 0
                    titProvisao.MoedaAcrescimos = 0
                    titProvisao.MoedaValorDoDocumento += 0
                    titProvisao.IUD = IIf(titProvisao.IUD = "I", "I", "U")
                    'Historico Financeiro Provisão.
                    titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, titPrevisao.Codigo, pValorDesconto, "Adicionado o valor de R$ : " & pValorDesconto.ToString("C2", CultureInfo.CurrentCulture) & ", ref a NF de devolução: " & NF_Modificada.Codigo & "-" & NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                    pValor -= pValorDesconto
                    pValorDesconto -= pValorDesconto

                End If

                If Not listTitulosExtornoDeDevolucao.Any(Function(x) x.Codigo = titPrevisao.Codigo) Then
                    listTitulosExtornoDeDevolucao.Add(titPrevisao)
                End If

                If pValorDesconto <= 0 Then
                    Exit For
                End If

            Next

            If pValor <= 0 Then
                Exit For
            End If

        Next


        If pValor > 0 Then

            'Jogamos a diferença de saldo da nota de devolução em um titulo de Provisao
            If titProvisao Is Nothing Then
                titProvisao = New Titulo(_VencimentosPedido.FirstOrDefault().Codigo)
                titProvisao.ListHistoricoFinanceiro.Clear()
                Dim n As Numerador = New Numerador(1)
                Dim NumeroTitulo As Integer = n.Sequencia + 1
                titProvisao.Codigo = NumeroTitulo
                titProvisao.CodigoProvisao = eProvisao.Provisao
                titProvisao.ValorDoDocumento = 0
                titProvisao.IUD = "I"
            End If

            'Adiciona na Provisão.
            titProvisao.CodigoProvisao = eProvisao.Provisao
            titProvisao.Deducoes = 0
            titProvisao.Descontos = 0
            titProvisao.Juros = 0
            titProvisao.Acrescimos = 0
            titProvisao.ValorDoDocumento += pValor
            titProvisao.MoedaDeducoes = 0
            titProvisao.MoedaDescontos = 0
            titProvisao.MoedaJuros = 0
            titProvisao.MoedaAcrescimos = 0
            titProvisao.IUD = IIf(titProvisao.IUD = "I", "I", "U")

            'Historico Financeiro Provisão.
            titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, 0, pValor, "Adicionado o valor de R$ : " & pValor.ToString("C2", CultureInfo.CurrentCulture) & ", ref a NF de devolução: " & NF_Modificada.Codigo & "-" & NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

        End If


        _NF_Modificada.VencimentosNota = _VencimentosNota

        If Not titProvisao Is Nothing Then
            _NF_Modificada.VencimentosNota.Add(titProvisao)
        End If

    End Sub

    Public Sub EstornoNotaDeDevolucao_bkp(ByRef pValor As Decimal)
        '*******************************************************************************
        '********** Notas associadas a NF de Devolução *********************************
        '*******************************************************************************
        Dim titProvisao = _VencimentosPedido.Find(Function(x) x.CodigoProvisao = eProvisao.Provisao)

        For Each nf As NotaFiscalDevolucaoXNotaFiscal In _NF_Devolucoes
            pValor = nf.ValorDevolucao
            '*******************************************************************************
            '********** Se tiver previsão na notas volta pra provisão  *********************
            '*******************************************************************************
            For Each titPrevisao In nf.Nota.VencimentosNota

                If titProvisao Is Nothing Then
                    titProvisao = New Titulo(titPrevisao.Codigo)
                    titProvisao.ListHistoricoFinanceiro.Clear()
                    Dim n As Numerador = New Numerador(1)
                    Dim NumeroTitulo As Integer = n.Sequencia + 1
                    titProvisao.Codigo = NumeroTitulo
                    titProvisao.CodigoProvisao = eProvisao.Provisao
                    titProvisao.ValorDoDocumento = 0
                    titProvisao.IUD = "I"
                End If

                titPrevisao.IUD = "U"
                titPrevisao.Historico = titPrevisao.Historico

                If (pValor >= titPrevisao.ValorDoDocumento) Then
                    Dim valor = titPrevisao.ValorDoDocumento
                    Dim valorDolar = titPrevisao.MoedaValorDoDocumento
                    titPrevisao.CodigoProvisao = eProvisao.Provisao
                    titPrevisao.ValorDoDocumento = 0
                    titPrevisao.Deducoes = 0
                    titPrevisao.Descontos = 0
                    titPrevisao.Juros = 0
                    titPrevisao.Acrescimos = 0
                    titPrevisao.MoedaValorDoDocumento = 0
                    titPrevisao.MoedaDeducoes = 0
                    titPrevisao.MoedaDescontos = 0
                    titPrevisao.MoedaJuros = 0
                    titPrevisao.MoedaAcrescimos = 0
                    titPrevisao.CodigoSituacao = eSituacao.Excluido
                    'Historico Financeiro Previsão.
                    titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, _NF_Modificada, 0, valor, "Retirado o valor de R$ : " & valor.ToString("C2", CultureInfo.CurrentCulture) & ", ref a NF de devolução: " & NF_Modificada.Codigo & "-" & NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))
                    'Adiciona na Provisão.
                    titProvisao.CodigoProvisao = eProvisao.Provisao
                    titProvisao.Deducoes = 0
                    titProvisao.Descontos = 0
                    titProvisao.Juros = 0
                    titProvisao.Acrescimos = 0
                    titProvisao.ValorDoDocumento += valor
                    titProvisao.MoedaDeducoes = 0
                    titProvisao.MoedaDescontos = 0
                    titProvisao.MoedaJuros = 0
                    titProvisao.MoedaAcrescimos = 0
                    titProvisao.MoedaValorDoDocumento += valorDolar
                    titProvisao.IUD = IIf(titProvisao.IUD = "I", "I", "U")

                    pValor -= valor

                    'Historico Financeiro Provisão.
                    titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, titPrevisao.Codigo, valor, "Adicionado o valor de R$ : " & valor.ToString("C2", CultureInfo.CurrentCulture) & ", ref a NF de devolução: " & NF_Modificada.Codigo & "-" & NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))
                Else
                    'Caso não exista provisão no pedido
                    titProvisao.IUD = IIf(titProvisao.IUD = "I", "I", "U")
                    titProvisao.Deducoes = 0
                    titProvisao.Descontos = 0
                    titProvisao.Juros = 0
                    titProvisao.Acrescimos = 0
                    titProvisao.ValorDoDocumento += pValor
                    'Desconta o restante da previsão
                    titPrevisao.Deducoes = 0
                    titPrevisao.Descontos = 0
                    titPrevisao.Juros = 0
                    titPrevisao.Acrescimos = 0
                    titPrevisao.ValorDoDocumento -= pValor
                    'Historico Financeiro Previsão.
                    titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, _NF_Modificada, 0, pValor, "Retirado o valor de R$ : " & pValor.ToString("C2", CultureInfo.CurrentCulture) & ", ref a NF de devolução: " & NF_Modificada.Codigo & "-" & NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))
                    'Historico Financeiro Provisão.
                    titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, _NF_Modificada, titPrevisao.Codigo, pValor, "Adicionado o valor de R$ : " & pValor.ToString("C2", CultureInfo.CurrentCulture) & ", ref a NF de devolução: " & NF_Modificada.Codigo & "-" & NF_Modificada.Serie & "-" & IIf(_NF_Modificada.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                    pValor = 0
                End If

                If pValor = 0 Then
                    Exit For
                End If
            Next
        Next

        _NF_Modificada.VencimentosNota = _VencimentosNota
        _NF_Modificada.VencimentosNota.Add(titProvisao)

    End Sub

    'Exclusão da NF de Devolução
    Public Sub CancelamentoNotaDeDevolucao(ByVal NF As NotaFiscal)
        '*******************************************************************************
        '********** Notas associadas a NF de Devolução *********************************
        '*******************************************************************************
        NF.VencimentosNota = New ListTitulo()
        Dim ListHistoricoFinan = New ListHistoricoFinanceiro(NF)
        _VencimentosNota = New ListTitulo

        If ListHistoricoFinan.Any() Then

            If ListHistoricoFinan.Any(Function(x) x.TituloOrigem > 0) Then

                For Each hf In ListHistoricoFinan.Where(Function(x) x.TituloOrigem > 0).Distinct()
                    Dim titPrevisao = New Titulo(hf.TituloOrigem)

                    If titPrevisao.CodigoProvisao <> eProvisao.Baixa Then
                        If titPrevisao.ReceberPagar Is Nothing Then
                            titPrevisao = NF.VencimentosPedido.FirstOrDefault().Copy()
                            titPrevisao.Codigo = hf.TituloOrigem
                            titPrevisao.IUD = "I"
                            titPrevisao.ValorDoDocumento = hf.Valor
                        Else
                            titPrevisao.IUD = "U"
                            titPrevisao.Historico = titPrevisao.Historico
                            titPrevisao.ValorDoDocumento += hf.Valor
                        End If

                        titPrevisao.CodigoProvisao = eProvisao.Previsao
                        titPrevisao.CodigoSituacao = eSituacao.Normal
                    End If

                    Dim codigoProvisao = hf.Titulo
                    'Provisão
                    If NF.VencimentosPedido.Any(Function(x) x.CodigoProvisao = eProvisao.Provisao) Then
                        Dim titProvisao = NF.VencimentosPedido.Find(Function(x) x.CodigoProvisao = eProvisao.Provisao)
                        titProvisao.IUD = "U"
                        titProvisao.ValorDoDocumento -= hf.Valor
                        If titProvisao.ValorDoDocumento <= 0 Then
                            titProvisao.CodigoSituacao = eSituacao.Excluido
                            titProvisao.ValorDoDocumento = 0
                        End If
                        codigoProvisao = titProvisao.Codigo
                        'Historico Financeiro Provisão.
                        titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, NF, 0, hf.Valor, "Retirado o valor de R$ : " & hf.Valor.ToString("C2", CultureInfo.CurrentCulture) & ", ref ao cancelamento da NF de devolução: " & NF.Codigo & "-" & NF.Serie & "-" & IIf(NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))
                        If Not _VencimentosNota.Where(Function(v) v.Codigo = titProvisao.Codigo).Any() Then
                            _VencimentosNota.Add(titProvisao)
                        End If
                    Else
                        Dim where As String = " Pedido = " & NF.Pedido.Codigo & vbCrLf &
                                      "    and Empresa = '" & NF.Pedido.CodigoEmpresa & "'" & vbCrLf &
                                      "    and EndEmpresa = " & NF.Pedido.EnderecoEmpresa &
                                      "    and Provisao = " & eProvisao.Provisao

                        Dim _vencimentosPedido = New Negocio.ListTitulo(where)
                        titPrevisao = _vencimentosPedido.FirstOrDefault(Function(x) x.CodigoProvisao = eProvisao.Provisao)
                        If titPrevisao Is Nothing Then

                            'Throw New Exception("Não é possivel excluir a NF, pois o pedido não possui título de provisão")
                            For Each item As [Lib].Negocio.NotaFiscalXItem In NF.Itens
                                For Each notaDevolucao As [Lib].Negocio.NotaFiscalDevolucaoXNotaFiscal In item.NotasDevolucao

                                    titPrevisao = notaDevolucao.Nota.VencimentosNota.FirstOrDefault().Copy()
                                    titPrevisao.Codigo = hf.TituloOrigem
                                    titPrevisao.IUD = "I"
                                    titPrevisao.ValorDoDocumento = NF.TotalNota
                                    'Se for devolução informamos a nota
                                    _VencimentosNota.NF = notaDevolucao.Nota
                                    Exit For

                                Next
                            Next

                        Else

                            codigoProvisao = titPrevisao.Codigo
                            titPrevisao.IUD = "U"
                            titPrevisao.CodigoProvisao = eProvisao.Provisao
                            titPrevisao.CodigoSituacao = eSituacao.Normal
                            titPrevisao.ValorDoDocumento += hf.Valor

                        End If

                    End If

                    'Historico Financeiro Previsão.
                    titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, NF, codigoProvisao, hf.Valor, "Adicionado o valor de R$ : " & hf.Valor.ToString("C2", CultureInfo.CurrentCulture) & ", ref ao cancelamento da NF de devolução: " & NF.Codigo & "-" & NF.Serie & "-" & IIf(NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))
                    _VencimentosNota.Add(titPrevisao)

                Next

            Else

                For Each hf In ListHistoricoFinan.Where(Function(x) x.TituloOrigem = 0).Distinct()

                    Dim titPrevisao = New Titulo(hf.Titulo)

                    'Se o titulo do historico ainda estiver aberto iremos usar ele
                    If titPrevisao.CodigoProvisao = eProvisao.Previsao Then

                        titPrevisao.IUD = "U"
                        titPrevisao.CodigoSituacao = eSituacao.Normal
                        titPrevisao.ValorDoDocumento += hf.Valor

                        'Historico Financeiro Previsão.
                        titPrevisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titPrevisao, NF, 0, hf.Valor, "Adicionado o valor de R$ : " & hf.Valor.ToString("C2", CultureInfo.CurrentCulture) & ", ref ao cancelamento da NF de devolução: " & NF.Codigo & "-" & NF.Serie & "-" & IIf(NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))
                        _VencimentosNota.Add(titPrevisao)

                    Else

                        'Senão tiver titulo de previsão, vamos usar um de provisão, e caso não tenha, teremos que criar
                        Dim where As String = " Pedido = " & NF.Pedido.Codigo & vbCrLf &
                        "    and Empresa = '" & NF.Pedido.CodigoEmpresa & "'" & vbCrLf &
                        "    and EndEmpresa = " & NF.Pedido.EnderecoEmpresa &
                        "    and Provisao = " & eProvisao.Provisao

                        Dim _vencimentosPedido = New Negocio.ListTitulo(where)
                        Dim titProvisao As Titulo
                        titProvisao = _vencimentosPedido.FirstOrDefault(Function(x) x.CodigoProvisao = eProvisao.Provisao)

                        If titProvisao Is Nothing Then

                            'Throw New Exception("Não é possivel excluir a NF, pois o pedido não possui título de provisão")
                            For Each item As [Lib].Negocio.NotaFiscalXItem In NF.Itens
                                For Each notaDevolucao As [Lib].Negocio.NotaFiscalDevolucaoXNotaFiscal In item.NotasDevolucao

                                    titProvisao = notaDevolucao.Nota.VencimentosNota.FirstOrDefault().Copy()
                                    titProvisao.IUD = "I"
                                    titProvisao.ValorDoDocumento = NF.TotalNota
                                    'Se for devolução informamos a nota
                                    _VencimentosNota.NF = notaDevolucao.Nota
                                    Exit For

                                Next
                            Next

                        Else

                            titProvisao.IUD = "U"
                            titProvisao.CodigoProvisao = eProvisao.Provisao
                            titProvisao.CodigoSituacao = eSituacao.Normal
                            titProvisao.ValorDoDocumento += hf.Valor

                        End If

                        'Historico Financeiro Provisao.
                        titProvisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(titProvisao, NF, 0, titProvisao.ValorDoDocumento, "Adicionado o valor de R$ : " & titProvisao.ValorDoDocumento.ToString("C2", CultureInfo.CurrentCulture) & ", ref ao cancelamento da NF de devolução: " & NF.Codigo & "-" & NF.Serie & "-" & IIf(NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))
                        _VencimentosNota.Add(titProvisao)

                    End If

                Next

            End If

        Else
            For Each item As [Lib].Negocio.NotaFiscalXItem In NF.Itens
                For Each notaDevolucao As [Lib].Negocio.NotaFiscalDevolucaoXNotaFiscal In item.NotasDevolucao

                    Dim previsao = notaDevolucao.Nota.VencimentosNota.Where(Function(x) x.CodigoProvisao = eProvisao.Previsao).FirstOrDefault()

                    If notaDevolucao.QuantidadeDevolucao > 0 Or notaDevolucao.ValorDevolucao > 0 Then
                        If previsao IsNot Nothing Then
                            previsao.IUD = "U"
                            previsao.ValorDoDocumento += notaDevolucao.ValorDevolucao
                        Else
                            previsao = New Titulo(notaDevolucao.Nota.VencimentosPedido.Where(Function(x) x.CodigoProvisao = eProvisao.Previsao).First().Codigo)
                            previsao.CodigoProvisao = eProvisao.Previsao
                            previsao.ValorDoDocumento = notaDevolucao.ValorDevolucao
                            previsao.IUD = "I"
                        End If

                        'Historico Financeiro Previsão.
                        previsao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(previsao, NF, 0, notaDevolucao.ValorDevolucao, "Adicionado o valor de R$ : " & notaDevolucao.ValorDevolucao.ToString("C2", CultureInfo.CurrentCulture) & ", ref ao cancelamento da NF de devolução: " & NF.Codigo & "-" & NF.Serie & "-" & IIf(NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

                        _VencimentosNota.Add(previsao)
                    End If
                Next
            Next

            'Desconta a provisão
            Dim provisao = NF.VencimentosPedido.Where(Function(x) x.CodigoProvisao = eProvisao.Provisao).First()
            provisao.IUD = "U"
            provisao.ValorDoDocumento -= NF.TotalNota

            'Historico Financeiro Provisão.
            provisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(provisao, NF, 0, NF.TotalNota, "Retirado o valor de R$ : " & NF.TotalNota.ToString("C2", CultureInfo.CurrentCulture) & ", ref ao cancelamento da NF de devolução: " & NF.Codigo & "-" & NF.Serie & "-" & IIf(NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))

            _VencimentosNota.Add(provisao)
        End If

        NF.VencimentosNota = _VencimentosNota

    End Sub

    ' Cria titulos de Baixa, Compensacao, Adiantamento e Provisao
    Private Function CriaTitulo(pTipo As String, ByRef pValor As Decimal, Optional Adiantamento As Novo.AdiantamentoNovo = Nothing, Optional NF As NotaFiscal = Nothing, Optional codigoPrevisaoOrigem As Integer = 0, Optional pCodigoFixacao As Integer = 0) As Novo.TituloNovo
        '***********************************
        '************ TIPO *****************
        '***********************************
        ' BX - Baixa Adiantamento
        '   pValor a ser baixado
        ' PRO - Provisao:
        '   pValor a ser convertito emProvisao
        '   pCodigoContaContabilRecPag conta contabil do titulo q estava em Previsao
        ' COM - Compensados
        '***********************************
        '***********************************
        Dim tit As New Novo.TituloNovo


        If pTipo = "BX" Then
            tit = New Novo.TituloNovo(Adiantamento.CodigoTitulo)
            tit.IUD = "I"
            tit.Codigo = 0

            tit.Adiantamento = Nothing
            tit.DataBaixa = Date.Now
            tit.Reprogramacao = Date.Now
            tit.DataMoeda = Date.Now
            tit.CodigoSituacao = eSituacao.Normal

            tit.CodigoMoeda = Pedido.CodigoMoeda

            tit.CodigoTipoPgto = 1
            tit.CodigoCarteiraDoTitulo = 1
            tit.CodigoContaContabilRecPag = Pedido.SubOperacao.CodigoGrupoContas

            tit.NotaTitulo = New Novo.NotaFiscalXTitulo(NF_Original, tit)

            tit.ReceberPagar = IIf(tit.ReceberPagar = "R", "P", "R")
            tit.AdiantamentosAbertos.Clear()
            tit.AdiantamentosAbertos.Add(Adiantamento)

            If NF_Modificada.NFG Then
                tit.CodigoPedido = NF_Modificada.CodigoPedido
            End If

            For Each vlr In tit.Valores
                If Not vlr.Equals(tit.Valores.EncargoValorDocumento) And Not vlr.Equals(tit.Valores.EncargoValorLiquido) Then
                    tit.Valores.Remove(vlr)
                End If
            Next

            tit.Valores.EncargoValorDocumento.DC = IIf(tit.Valores.EncargoValorDocumento.DC = "C", "D", "C")
            tit.Valores.EncargoValorLiquido.DC = IIf(tit.Valores.EncargoValorDocumento.DC = "C", "D", "C")


            If Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then

                tit.Valores.EncargoValorDocumento.ValorOficial = pValor
                tit.Valores.EncargoValorDocumento.ValorMoeda = Math.Round(pValor / NF_Modificada.IndiceNota, 2)

                tit.Valores.EncargoValorLiquido.ValorOficial = pValor
                tit.Valores.EncargoValorLiquido.ValorMoeda = Math.Round(pValor / NF_Modificada.IndiceNota, 2)
            Else
                tit.Valores.EncargoValorDocumento.ValorOficial = Math.Round(pValor * NF_Modificada.IndiceNota, 2)
                tit.Valores.EncargoValorDocumento.ValorMoeda = pValor

                tit.Valores.EncargoValorLiquido.ValorOficial = Math.Round(pValor * NF_Modificada.IndiceNota, 2)
                tit.Valores.EncargoValorLiquido.ValorMoeda = pValor
            End If

            tit.Historico = "Baixa do Adiantamento " & Adiantamento.CodigoTitulo & " do Pedido " & Pedido.Codigo
            tit.Observacoes = ""
        End If

        If pTipo = "PRO" Then
            tit.IUD = "I"
            tit.Codigo = 0
            tit.CodigoProvisao = eProvisao.Provisao
            tit.CodigoPedido = NF_Modificada.CodigoPedido
            tit.CodigoSituacao = eSituacao.Normal
            tit.ReceberPagar = IIf(Pedido.SubOperacao.EntradaSaida, "P", "R")

            'CLI-FOR
            tit.CodigoUnidadeDeNegocio = Pedido.CodigoUnidadeNegocio
            tit.CodigoEmpresa = Pedido.CodigoEmpresa
            tit.EnderecoEmpresa = Pedido.EnderecoEmpresa
            tit.CodigoCliFor = Pedido.CodigoCliente
            tit.EnderecoCliFor = Pedido.EnderecoCliente

            'REC-PAG
            tit.CodigoUnidadeDeNegocioRecPag = Pedido.CodigoUnidadeNegocio
            tit.CodigoEmpresaRecPag = Pedido.CodigoEmpresa
            tit.EndEmpresaRecPag = Pedido.EnderecoEmpresa

            tit.CodigoIndexador = Pedido.CodigoIndexador
            tit.CodigoTipoPgto = 1
            tit.CodigoCarteiraDoTitulo = 1
            tit.CodigoProvisao = 3

            tit.Movimento = Date.Now
            tit.Vencimento = Date.Now
            tit.DataBaixa = Date.Now
            tit.Reprogramacao = Date.Now
            tit.DataMoeda = Date.Now

            tit.CodigoMoeda = Pedido.CodigoMoeda

            tit.Historico = "Titulo Provisionado pela Emissão da Nota/Serie " & NF_Modificada.Codigo & "-" & NF_Modificada.Serie & "  Referente ao pedido " & Pedido.Codigo
            tit.Observacoes = ""

            tit.CodigoContaContabilCliFor = Pedido.SubOperacao.CodigoGrupoContas
            tit.CodigoContaContabilRecPag = Pedido.Empresa.Empresa.CodigoContaGrupoBanco
            'tit.CodigoContaContabilRecPag = pCodigoContaContabilRecPag

            'Previsão de origem do titulo
            tit.CodigoTituloOrigem = codigoPrevisaoOrigem

            tit.NotaTitulo = New Novo.NotaFiscalXTitulo(IIf(NF IsNot Nothing, NF, NF_Modificada), tit)

            tit.Valores.EncargoValorDocumento.Valor = pValor
        End If

        If pTipo = "COMP" Then
            tit.IUD = "I"
            tit.Codigo = 0
            tit.CodigoProvisao = eProvisao.Compensado
            tit.CodigoPedido = Pedido.Codigo

            tit.ReceberPagar = IIf(Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "R", "P")
            tit.CodigoSituacao = eSituacao.Normal
            tit.Movimento = Date.Now
            tit.Vencimento = Date.Now
            tit.DataBaixa = Date.Now
            tit.Reprogramacao = Date.Now
            tit.DataMoeda = Date.Now

            tit.CodigoMoeda = Pedido.CodigoMoeda

            'CLI-FOR
            tit.CodigoUnidadeDeNegocio = Pedido.CodigoUnidadeNegocio
            tit.CodigoEmpresa = Pedido.CodigoEmpresa
            tit.EnderecoEmpresa = Pedido.EnderecoEmpresa
            tit.CodigoCliFor = Pedido.CodigoCliente
            tit.EnderecoCliFor = Pedido.EnderecoCliente

            'REC-PAG
            tit.CodigoUnidadeDeNegocioRecPag = Pedido.CodigoUnidadeNegocio
            tit.CodigoEmpresaRecPag = Pedido.CodigoEmpresa
            tit.EndEmpresaRecPag = Pedido.EnderecoEmpresa

            tit.CodigoIndexador = Pedido.CodigoIndexador
            tit.CodigoTipoPgto = 1
            tit.CodigoCarteiraDoTitulo = 1
            tit.CodigoProvisao = 4

            tit.Historico = "Titulo Compensado pela Emissão da Nota/Serie " & NF_Modificada.Codigo & "-" & NF_Modificada.Serie & "  Referente ao pedido " & NF_Modificada.CodigoPedido
            tit.Observacoes = ""

            tit.CodigoContaContabilCliFor = Pedido.SubOperacao.CodigoGrupoContas
            tit.CodigoContaContabilRecPag = Pedido.Empresa.Empresa.CodigoContaGrupoBanco

            tit.NotaTitulo = New Novo.NotaFiscalXTitulo(NF_Modificada, tit)
            tit.CodigoTituloCompensacao = codigoPrevisaoOrigem

            tit.Valores.EncargoValorDocumento.Valor = pValor
        End If

        If pTipo = "ADI" Then
            tit.IUD = "I"
            tit.Codigo = 0
            tit.CodigoProvisao = eProvisao.Baixa
            tit.CodigoPedido = Pedido.Codigo
            tit.CodigoSituacao = eSituacao.Normal
            tit.Movimento = Date.Now
            tit.Vencimento = Date.Now
            tit.DataBaixa = Date.Now
            tit.Reprogramacao = Date.Now
            tit.DataMoeda = Date.Now

            tit.ReceberPagar = IIf(Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "P", "R")
            tit.CodigoMoeda = Pedido.CodigoMoeda

            'CLI-FOR
            tit.CodigoUnidadeDeNegocio = Pedido.CodigoUnidadeNegocio
            tit.CodigoEmpresa = Pedido.CodigoEmpresa
            tit.EnderecoEmpresa = Pedido.EnderecoEmpresa
            tit.CodigoCliFor = Pedido.CodigoCliente
            tit.EnderecoCliFor = Pedido.EnderecoCliente

            'REC-PAG
            tit.CodigoUnidadeDeNegocioRecPag = Pedido.CodigoUnidadeNegocio
            tit.CodigoEmpresaRecPag = Pedido.CodigoEmpresa
            tit.EndEmpresaRecPag = Pedido.EnderecoEmpresa

            tit.CodigoIndexador = Pedido.CodigoIndexador
            tit.CodigoTipoPgto = 1
            tit.CodigoCarteiraDoTitulo = 1

            'Fixação
            If pCodigoFixacao > 0 Then
                tit.CodigoFixacao = pCodigoFixacao
            End If

            tit.Historico = "ADIANTAMENTO REF. NF: " & NF_Modificada.Codigo & " / " & NF_Modificada.Cliente.Nome & "-" & IIf(NF_Modificada.SubOperacao.EntradaSaida = eEntradaSaida.Saida, "Saida", "Entrada")

            tit.Observacoes = ""

            tit.CodigoContaContabilCliFor = Pedido.SubOperacao.CodigoContaAdiantamento
            tit.CodigoContaContabilRecPag = Pedido.SubOperacao.CodigoGrupoContas

            tit.NotaTitulo = New Novo.NotaFiscalXTitulo(NF_Modificada, tit)

            tit.Valores.EncargoValorDocumento.Valor = pValor

            tit.Adiantamento.Taxa = 0
            tit.Adiantamento.Vencimento = Date.Now
        End If

        Return tit
    End Function
#End Region

End Class


