Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports System.Web
Imports NGS.Lib.Uteis

'**************************************************************************************************************************
'************************************** LISTA Reconstroi Financeiro Pedidos ***********************************************
'**************************************************************************************************************************
Public Class ListReconstroiFinanceiroPedido
    Inherits List(Of ReconstroiFinanceiroPedido)

#Region "Contrutor"
    Public Sub New(pNF As NotaFiscal)
        _Pedido = pNF.Pedido
        _NF = pNF

        Dim sql As String = ""
        sql = "Select Documento, NumeroDocumento, Serie, ES, Lancamento, CondicaoPagamento, Marcacao, ValorOficial, ROUND(ValorMoeda,2) as ValorMoeda" & vbCrLf & _
              "  From" & vbCrLf & _
              "		(" & vbCrLf & _
              "		SELECT 'P' as Documento," & vbCrLf & _
              "			   PxV.Pedido_Id as NumeroDocumento," & vbCrLf & _
              "			   '' as Serie," & vbCrLf & _
              "			   '' as ES," & vbCrLf & _
              "			   PxV.DataAlteracao as Lancamento," & vbCrLf & _
              "			   P.CondicaoPagamento," & vbCrLf & _
              "			   case" & vbCrLf & _
              "			     when ValorLiquido  = ValorAjustado then 'N'" & vbCrLf & _
              "			     when ValorAjustado > 0 then 'C'" & vbCrLf & _
              "			     when ValorAjustado < 0 then 'E'" & vbCrLf & _
              "			   end as Marcacao," & vbCrLf & _
              "			   case" & vbCrLf & _
              "			     when M.Classificacao = 'O'" & vbCrLf & _
              "			       then ValorAjustado" & vbCrLf & _
              "			       else 0" & vbCrLf & _
              "			    end as ValorOficial," & vbCrLf & _
              "			   case" & vbCrLf & _
              "			     when M.Classificacao = 'M'" & vbCrLf & _
              "			       then ValorAjustado" & vbCrLf & _
              "			       else 0" & vbCrLf & _
              "			    end as ValorMoeda" & vbCrLf & _
              "		  FROM Pedidos P" & vbCrLf & _
              "		 Inner Join PedidoxValorLiquido AS PxV" & vbCrLf & _
              "		    on P.Empresa_Id    = PxV.Empresa_Id" & vbCrLf & _
              "		   and P.EndEmpresa_Id = PxV.EndEmpresa_Id" & vbCrLf & _
              "		   and P.Pedido_Id     = PxV.Pedido_id" & vbCrLf & _
              "         Inner Join Moedas M" & vbCrLf & _
              "            on P.moeda = M.Moeda_id" & vbCrLf & _
              "		 Where P.Empresa_Id    ='" & pNF.Pedido.CodigoEmpresa & "'" & vbCrLf & _
              "		   and P.EndEmpresa_Id = " & pNF.Pedido.EnderecoEmpresa & vbCrLf & _
              "		   and P.Pedido_id     ='" & pNF.Pedido.Codigo & "'" & vbCrLf & _
              "		union all" & vbCrLf & _
              "        SELECT 'A', T.Titulo_Id," & vbCrLf & _
              "		       '' as Serie," & vbCrLf & _
              "			   '' as ES," & vbCrLf & _
              "			   T.HoraBaixa," & vbCrLf & _
              "			   0," & vbCrLf & _
              "			   T.RecPag," & vbCrLf & _
              "			   ValorOficial," & vbCrLf & _
              "			   ValorMoeda" & vbCrLf & _
              "		  from Titulos T" & vbCrLf & _
              "		 inner join TitulosxContaContabil TC" & vbCrLf & _
              "			on T.Titulo_Id           = TC.Titulo_Id" & vbCrLf & _
              "		   and T.ContaContabilCliFor = TC.Conta_Id" & vbCrLf & _
              "		   and TC.DC_Id = Case" & vbCrLf & _
              "							When T.RecPag in ('P','C')" & vbCrLf & _
              "							  then 'D'" & vbCrLf & _
              "							  else 'C'" & vbCrLf & _
              "						  end" & vbCrLf & _
              "		 inner join Adiantamentos A" & vbCrLf & _
              "			on T.Titulo_Id = A.Titulo_Id" & vbCrLf & _
              "         inner join BancosXContas BxC" & vbCrLf & _
              "            on T.ContaContabilRecPag = BxC.ContaContabil" & vbCrLf & _
              "		 Where T.Empresa    ='" & pNF.Pedido.CodigoEmpresa & "'" & vbCrLf & _
              "		   and T.EndEmpresa = " & pNF.Pedido.EnderecoEmpresa & vbCrLf & _
              "		   and T.Pedido     ='" & pNF.Pedido.Codigo & "'" & vbCrLf & _
              "		union all" & vbCrLf & _
              "		select 'BA',Bx.TituloBaixa, '','',TB.HoraBaixa, 0, TB.RecPag, BX.ValorOficial, Bx.ValorMoeda" & vbCrLf & _
              "		  from Adiantamentosxbaixas Bx" & vbCrLf & _
              "		 inner join Titulos T" & vbCrLf & _
              "		    on T.Titulo_Id = Bx.Titulo_Id" & vbCrLf & _
              "		 inner join Pedidos P" & vbCrLf & _
              "		    on T.Empresa    = P.Empresa_Id" & vbCrLf & _
              "		   and T.EndEmpresa = P.EndEmpresa_Id" & vbCrLf & _
              "		   and T.Pedido     = P.Pedido_Id" & vbCrLf & _
              "		 inner Join Titulos TB" & vbCrLf & _
              "		    on TB.Titulo_Id = Bx.TituloBaixa" & vbCrLf & _
              "		 inner join BancosXContas BC" & vbCrLf & _
              "		    on TB.EmpresaRecPag       = BC.Empresa_Id" & vbCrLf & _
              "		   and TB.EndEmpresaRecPag    = BC.EndEmpresa_Id" & vbCrLf & _
              "		   and TB.ContaContabilRecPag = BC.ContaContabil" & vbCrLf & _
              "         where T.Empresa ='" & pNF.Pedido.CodigoEmpresa & "'" & vbCrLf & _
              "		   and T.EndEmpresa = " & pNF.Pedido.EnderecoEmpresa & vbCrLf & _
              "		   and T.Pedido     ='" & pNF.Pedido.Codigo & "'" & vbCrLf & _
              "		union all" & vbCrLf & _
              "		SELECT 'NF'," & vbCrLf & _
              "		       NF.Nota_Id," & vbCrLf & _
              "		       NF.Serie_Id," & vbCrLf & _
              "		       NF.EntradaSaida_Id," & vbCrLf & _
              "		       NF.usuarioinclusaodata," & vbCrLf & _
              "		       0," & vbCrLf & _
              "		       SO.Devolucao," & vbCrLf & _
              "		       SUM(nfe.Valor)," & vbCrLf & _
              "            SUM(nfe.Valor) / (case " & vbCrLf & _
              "                                when P.IndexadorFixo = 1" & vbCrLf & _
              "                                  then P.IndiceFixado" & vbCrLf & _
              "                                  else C.indice" & vbCrLf & _
              "                              end) as ValorMoeda" & vbCrLf & _
              "		  FROM NotasFiscais AS NF" & vbCrLf & _
              "		 INNER JOIN NotasFiscaisXEncargos AS NFe" & vbCrLf & _
              "			ON NF.Empresa_Id      = NFe.Empresa_Id" & vbCrLf & _
              "		   AND NF.EndEmpresa_Id   = NFe.EndEmpresa_Id" & vbCrLf & _
              "		   AND NF.Cliente_Id      = NFe.Cliente_Id" & vbCrLf & _
              "		   AND NF.EndCliente_Id   = NFe.EndCliente_Id" & vbCrLf & _
              "		   AND NF.EntradaSaida_Id = NFe.EntradaSaida_Id" & vbCrLf & _
              "		   AND NF.Serie_Id        = NFe.Serie_Id" & vbCrLf & _
              "		   AND NF.Nota_Id         = NFe.Nota_Id" & vbCrLf & _
              "		   and NFe.Encargo_Id     = 'LIQUIDO'" & vbCrLf & _
              "		 Inner Join SubOperacoes SO" & vbCrLf & _
              "			on NF.Operacao    = SO.Operacao_Id" & vbCrLf & _
              "		   and NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
              "		 Inner Join Pedidos P" & vbCrLf & _
              "			on NF.Empresa_Id    = P.Empresa_Id" & vbCrLf & _
              "		   and NF.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
              "		   and NF.Pedido        = P.Pedido_Id" & vbCrLf & _
              "      Inner join Cotacoes C" & vbCrLf & _
              "         on C.Indexador_id = P.Indexador" & vbCrLf & _
              "        and C.Data_id      = NF.DataDaNota  " & vbCrLf & _
              "		 where P.Empresa_Id     ='" & pNF.Pedido.CodigoEmpresa & "'" & vbCrLf & _
              "		   and P.EndEmpresa_Id  = " & pNF.Pedido.EnderecoEmpresa & vbCrLf & _
              "		   and P.Pedido_id      ='" & pNF.Pedido.Codigo & "'" & vbCrLf & _
              "		 Group by NF.Nota_Id, NF.Serie_Id, NF.EntradaSaida_Id, NF.usuarioinclusaodata,so.Devolucao, P.IndiceFixado,P.IndexadorFixo, c.Indice" & vbCrLf & _
              "     ) sb" & vbCrLf & _
              " order by Lancamento," & vbCrLf & _
              "          case" & vbCrLf & _
              "            When Documento = 'P'  then 1" & vbCrLf & _
              "            When Documento = 'A'  then 2" & vbCrLf & _
              "            When Documento = 'NF' then 3" & vbCrLf & _
              "            else 4" & vbCrLf & _
              "          end" & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "RoteiroReconstrucao")

        For Each row In ds.Tables(0).Rows
            Dim RFP As New ReconstroiFinanceiroPedido(_Pedido)
            RFP.Documento = row("Documento")
            RFP.NumeroDocumento = row("NumeroDocumento")
            RFP.Serie = row("Serie")
            RFP.ES = row("ES")
            RFP.Lancamento = row("Lancamento")
            RFP.CondicaoPagamento = row("CondicaoPagamento")
            RFP.Marcacao = row("Marcacao")
            RFP.ValorOficial = row("ValorOficial")
            RFP.ValorMoeda = row("ValorMoeda")
            Me.Add(RFP)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Pedido As Pedido
    Private _NF As NotaFiscal
    Private _TitulosOld As Novo.ListTituloNovo
    Private _TitulosNew As Novo.ListTituloNovo

    Private _AdiantamentosAbertos As New Hashtable
#End Region

#Region "Property"
    Public ReadOnly Property Pedido As Pedido
        Get
            Return _Pedido
        End Get
    End Property

    Public ReadOnly Property NF As NotaFiscal
        Get
            Return _NF
        End Get
    End Property

    Public Property TitulosOld As Novo.ListTituloNovo
        Get
            If _TitulosOld Is Nothing Then
                Dim where As String
                where = "     Empresa    ='" & Pedido.CodigoEmpresa & "'" & vbCrLf & _
                        " And EndEmpresa = " & Pedido.EnderecoEmpresa & vbCrLf & _
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
                _TitulosNew.NF = NF
            End If
            Return _TitulosNew
        End Get
        Set(value As Novo.ListTituloNovo)
            _TitulosNew = value
        End Set
    End Property

    Public Property AdiantamentosAbertos As Hashtable
        Get
            Return _AdiantamentosAbertos
        End Get
        Set(value As Hashtable)
            _AdiantamentosAbertos = value
        End Set
    End Property

    Public ReadOnly Property SaldoAdiantamento As Decimal
        Get
            Dim Saldo As Decimal
            For Each row In AdiantamentosAbertos
                Saldo += row.Value
            Next
            Return Saldo
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Valida() As ArrayList
        '0 Verdadeiro Ou Falso
        '1 Msg - Sucesso ou Erro 
        Dim Retorno As New ArrayList

        Dim Previsao As Decimal
        Dim Provisao As Decimal
        Dim Adiantamento As Decimal
        Dim BxAdiantamento As Decimal
        Dim Vlr As Decimal

        For Each doc In Me
            If doc.NumeroDocumento = NF.Codigo _
            And doc.Serie = NF.Serie _
            And doc.ES = NF.EntradaSaida.ToString.Substring(0, 1) Then
                If NF.IUD = "D" Then Continue For
                If NF.IUD = "U" Then
                    Vlr = NF.TotalNota * IIf(Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, 1, NF.IndiceNota)
                End If
            Else
                Vlr = IIf(Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, doc.ValorOficial, doc.ValorMoeda)
            End If
            Vlr = Math.Round(Vlr, 2)

            Select Case doc.Documento
                Case "P"
                    Select Case doc.Marcacao
                        Case "N"
                            Previsao = Vlr
                        Case "C"
                            Previsao += Vlr
                        Case "E"
                            If Previsao < Vlr Then
                                Retorno.Add(False)
                                Retorno.Add("O Estorno realizado no dia " & doc.Lancamento.ToString("dd/MM/yyyy") & " Não vai ter Saldo Disponivel")
                                Return Retorno
                            Else
                                Previsao -= Vlr
                            End If
                    End Select

                Case "A"
                    If Vlr > Previsao Then
                        Retorno.Add(False)
                        Retorno.Add("O Adiantamento " & doc.NumeroDocumento & " realizado no dia " & doc.Lancamento.ToString("dd/MM/yyyy") & " Não vai ter Saldo Previsão Disponivel para ser Realizado")
                        Return Retorno
                    Else
                        Previsao -= Vlr
                        Adiantamento += Vlr
                    End If

                Case "NF"
                    If doc.Marcacao = "N" Then
                        If Adiantamento > 0 Then
                            If Vlr > Adiantamento Then
                                Previsao -= Vlr - Adiantamento
                                Provisao += Vlr - Adiantamento
                                BxAdiantamento += Adiantamento
                                Adiantamento = 0
                            Else
                                BxAdiantamento += Vlr
                                Adiantamento -= Vlr
                            End If
                        Else
                            Previsao -= Vlr
                            Provisao += Vlr
                        End If

                        If Previsao < 0 Then
                            Retorno.Add(False)
                            Retorno.Add("A Nota " & doc.NumeroDocumento & " realizado no dia " & doc.Lancamento.ToString("dd/MM/yyyy") & " Não vai ter Saldo Disponivel para ser Realizado")
                            Return Retorno
                        End If
                    Else
                        If Provisao > 0 Then
                            If Provisao > Vlr Then
                                Provisao -= Vlr
                                Previsao += Vlr
                                Vlr = 0
                            Else
                                Previsao += Provisao
                                Vlr -= Provisao
                                Provisao = 0
                            End If
                        End If

                        If BxAdiantamento > 0 And Vlr > 0 Then
                            If BxAdiantamento >= Vlr Then
                                Adiantamento += BxAdiantamento - Vlr
                                BxAdiantamento -= Vlr
                                Vlr = 0
                            Else
                                Adiantamento += BxAdiantamento
                                Vlr -= BxAdiantamento
                                BxAdiantamento = 0
                            End If
                        End If

                        If Vlr > 0 Then
                            Retorno.Add(False)
                            Retorno.Add("A Nota " & doc.NumeroDocumento & " realizado no dia " & doc.Lancamento.ToString("dd/MM/yyyy") & " Não vai ter Saldo Disponivel para ser Realizado")
                            Return Retorno
                        End If
                    End If
                Case "BA"
                    If Adiantamento < Vlr Then
                        Retorno.Add(False)
                        Retorno.Add("O Titulo " & doc.NumeroDocumento & " de baixa de adiantamento realizado no dia " & doc.Lancamento.ToString("dd/MM/yyyy") & " Não vai ter Saldo Disponivel para ser Realizado")
                        Return Retorno
                    Else
                        Previsao += Vlr
                        Adiantamento -= Vlr
                    End If
            End Select
        Next

        Retorno.Add(True)
        Retorno.Add("")
        Return Retorno
    End Function

    Public Function Reconstroi() As ArrayList
        '0 Verdadeiro Ou Falso
        '1 Msg - Sucesso ou Erro 
        Dim Retorno As New ArrayList

        Retorno = Valida()

        If Retorno(0) = False Then
            Return Retorno
        Else
            Retorno.Clear()
        End If

        For Each row In Me
            Select Case row.Documento
                Case "P" : ReconstroiLancamentosPedido(row)
                Case "A" : ReconstroiLancamentosAdiantamentos(row)
                Case "NF" : ReconstroiLancamentosNotaFiscal(row)
                Case "BA" : ReconstroiLancamentosBaixaDeAdiantamentos(row)
            End Select
        Next
        Return Retorno
    End Function

    '***********************************************************************************************************
    '********** Recontroi Lancamentos Pelo Pedido **************************************************************
    '***********************************************************************************************************
    Private Sub ReconstroiLancamentosPedido(row As ReconstroiFinanceiroPedido)
        'Marcacao representa Lancamentos N - Normais, C - Complementos, E - Estornos
        Select Case row.Marcacao
            Case "N" : ReconstroiLancamentosPedidoNormal(row)
            Case "C" : ReconstroiLancamentosPedidoComplemento(row)
            Case "E" : ReconstroiLancamentosPedidoEstorno(row)
        End Select
    End Sub

    Private Sub ReconstroiLancamentosPedidoNormal(row As ReconstroiFinanceiroPedido)
        'Inicio do Pedido, volta dos os titulos que estavao em previsao para seu estado original
        'mesmo que eles ja tenha sido excluido e seu valores alterados
        Dim Vlr As Decimal
        Dim VlrParcelas As Decimal
        Dim VlrDiferenca As Decimal
        Dim CP As New CondicaoPagamento(row.CondicaoPagamento)

        Vlr = row.Valor
        VlrParcelas = Math.Round(Vlr / CP.Parcelas, 2)
        VlrDiferenca = Vlr - (VlrParcelas * CP.Parcelas)

        For Each tit In TitulosOld.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao).OrderBy(Function(s) s.Codigo).Take(CP.Parcelas)
            tit.IUD = "U"
            tit.CodigoSituacao = 1
            'Remove todos os encargos a nao ser o valor do documento e o valor liquido
            Dim objTitulo As [Lib].Negocio.Novo.TituloNovo = tit
            tit.Valores.RemoveAll(Function(s) Not s.Equals(objTitulo.Valores.EncargoValorDocumento) And Not s.Equals(objTitulo.Valores.EncargoValorLiquido))
            tit.CodigoContaContabilRecPag = tit.Empresa.Empresa.CodigoContaGrupoBanco
            tit.Valores.EncargoValorDocumento.Valor = VlrParcelas + VlrDiferenca
            VlrDiferenca = 0
            TitulosOld.Remove(tit)
            TitulosNew.Add(tit)
        Next
    End Sub

    Private Sub ReconstroiLancamentosPedidoComplemento(row As ReconstroiFinanceiroPedido)
        'Complemento do Pedido
        Dim Vlr As Decimal
        Dim VlrParcelas As Decimal
        Dim VlrDiferenca As Decimal
        Dim NumTitulosEmPrevisaoAtivos As Decimal = TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao And s.CodigoSituacao = eSituacao.Normal).Count

        'Caso exista titulo em previsao ativos ele divide o valor do complemento entre eles
        If NumTitulosEmPrevisaoAtivos > 0 Then
            Vlr = row.Valor
            VlrParcelas = Math.Round(Vlr / NumTitulosEmPrevisaoAtivos, 2)
            VlrDiferenca = Vlr - (VlrParcelas * NumTitulosEmPrevisaoAtivos)

            For Each tit In TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao And s.CodigoSituacao = eSituacao.Normal)
                tit.IUD = "U"
                tit.Valores.EncargoValorDocumento.Valor += VlrParcelas + VlrDiferenca
                VlrDiferenca = 0
            Next
        Else
            'caso NAO exista titulo em previsao ativos ele seleciona o Titulo em previsao Cancelado de Maior vencimento e volta ele para Ativo com o valor do complemento
            Dim tit As Novo.TituloNovo = TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao).OrderByDescending(Function(s) s.Reprogramacao).FirstOrDefault()
            tit.IUD = "U"
            tit.CodigoSituacao = eSituacao.Normal
            tit.Valores.EncargoValorDocumento.Valor = row.Valor
        End If
    End Sub

    Private Sub ReconstroiLancamentosPedidoEstorno(row As ReconstroiFinanceiroPedido)
        Dim Vlr As Decimal
        Dim VlrParcelas As Decimal
        Dim VlrDiferenca As Decimal
        Dim NumTitulosEmPrevisaoAtivos As Decimal = TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao And s.CodigoSituacao = eSituacao.Normal).Count

        Vlr = row.Valor
        VlrParcelas = Math.Round(Vlr / NumTitulosEmPrevisaoAtivos, 2)
        VlrDiferenca = Vlr - (VlrParcelas * NumTitulosEmPrevisaoAtivos)

        While Vlr > 0
            For Each tit In TitulosOld.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao And s.CodigoSituacao = eSituacao.Normal)
                If tit.Valores.EncargoValorDocumento.Valor <= VlrParcelas Then
                    tit.IUD = "C"
                    tit.CodigoSituacao = eSituacao.Cancelado
                    Vlr -= tit.Valores.EncargoValorDocumento.Valor
                Else
                    tit.IUD = "U"
                    tit.Valores.EncargoValorDocumento.Valor -= VlrParcelas + VlrDiferenca
                    Vlr -= VlrParcelas + VlrDiferenca
                    VlrDiferenca = 0
                End If
            Next

            If Vlr > 0 Then
                NumTitulosEmPrevisaoAtivos = TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao And s.CodigoSituacao = eSituacao.Normal).Count
                VlrParcelas = Math.Round(Vlr / NumTitulosEmPrevisaoAtivos, 2)
                VlrDiferenca = Vlr - (VlrParcelas * NumTitulosEmPrevisaoAtivos)
            End If
        End While
    End Sub

    '***********************************************************************************************************
    '********** Reconstroi Lancamentos Adiantamentos ***********************************************************
    '***********************************************************************************************************
    Private Sub ReconstroiLancamentosAdiantamentos(row As ReconstroiFinanceiroPedido)
        Dim tit As Novo.TituloNovo = TitulosOld.Find(Function(s) s.Codigo = row.NumeroDocumento)

        Dim vlr As Decimal
        vlr = row.Valor
        AdiantamentosAbertos.Add(row.NumeroDocumento, row.Valor)

        For Each tit In TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao And s.CodigoSituacao = eSituacao.Normal)
            If vlr > tit.Valores.EncargoValorDocumento.Valor Then
                tit.IUD = "C"
                tit.CodigoSituacao = eSituacao.Cancelado
                vlr -= tit.Valores.EncargoValorDocumento.Valor
            Else
                tit.Valores.EncargoValorDocumento.Valor -= vlr
                Exit For
            End If
        Next
        TitulosOld.Remove(tit)
        TitulosNew.Add(tit)
    End Sub

    '***********************************************************************************************************
    '********** Reconstroi Lancamentos Nota Fiscal *************************************************************
    '***********************************************************************************************************
    Private Sub ReconstroiLancamentosNotaFiscal(row As ReconstroiFinanceiroPedido)
        'Marcacao Representa Devolucao S - Sim, N - Nao
        If row.Marcacao = "N" Then
            ReconstroiLancamentosNotaFiscalNormal(row)
        Else
            ReconstroiLancamentosNotaFiscalDevolucao(row)
        End If
    End Sub

    Private Function ReconstroiLancamentosNotaFiscalNormal(row As ReconstroiFinanceiroPedido) As ArrayList
        '0 Verdadeiro Ou Falso
        '1 Msg - Sucesso ou Erro 
        Dim Retorno As New ArrayList

        'Valor Total da Nota
        Dim vlr As Decimal = row.Valor

        '***********************************************************************************************************************************************************************************
        '************************************************ Totalizadores dos titulos da NOTA ************************************************************************************************
        '***********************************************************************************************************************************************************************************
        'Totalizadores da lista de novos Titulos
        Dim VlrPrevisao As Decimal = TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao And s.CodigoSituacao = eSituacao.Normal).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
        Dim VlrProvisao As Decimal = TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao And s.CodigoSituacao = eSituacao.Normal).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)

        'Totalizadores dos Titulos da Nota Antiga
        Dim VlrNFProvisao As Decimal = row.NF.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Provisao).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
        Dim VlrNFBaixado As Decimal = row.NF.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And Not s.isAdiantamento And Not s.isBaixaAdiantamento).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
        Dim VlrNFBaixaAdiantamento As Decimal = row.NF.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isBaixaAdiantamento).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)

        '***********************************************************************************************************************************************************************************
        '************************  Primeiro Mantem os Titulos Baixados postergando as baixas de adiantamento caso existam  *****************************************************************
        '***********************************************************************************************************************************************************************************
        If VlrNFBaixado > 0 Then
            If VlrNFBaixado > vlr Then
                Retorno.Add(False)
                Retorno.Add("Titulo Baixado em Nota sem Valor Fiscal suficiente")
                Return Retorno
            End If
            For Each t In row.NF.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And Not s.isAdiantamento And Not s.isBaixaAdiantamento)
                TitulosOld.Remove(t)
                TitulosNew.Add(t)
            Next
            vlr -= VlrNFBaixado
        End If

        '***********************************************************************************************************************************************************************************
        '**********************************************     Segundo Mata o Adiantamento    *************************************************************************************************
        '***********************************************************************************************************************************************************************************
        If SaldoAdiantamento > 0 Then
            If VlrNFBaixaAdiantamento > 0 Then
                Dim tit As Novo.TituloNovo = row.NF.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isBaixaAdiantamento).FirstOrDefault
                tit.IUD = "U"
                tit.Valores.EncargoValorDocumento.Valor = IIf(SaldoAdiantamento > vlr, vlr, SaldoAdiantamento)
                tit.AdiantamentosAbertos.Clear()

                For Each h In AdiantamentosAbertos
                    If h.value = 0 Then Continue For

                    Dim vlrbx As Decimal
                    If h.value > vlr Then
                        vlrbx = vlr
                        vlr = 0
                        h.value -= vlr
                    Else
                        vlrbx = h.value
                        vlr -= h.value
                        h.value = 0
                    End If

                    Dim adi As New Novo.AdiantamentoNovo(New Novo.TituloNovo(h.key))
                    adi.VlrBaixa = vlrbx
                    tit.AdiantamentosAbertos.Add(adi)
                Next
                TitulosOld.Remove(TitulosOld.Where(Function(s) s.Codigo = tit.Codigo).FirstOrDefault)
                TitulosNew.Add(tit)
            Else
                Dim tituloBxAd As Novo.TituloNovo = CriaTitulo(row, "BX", vlr)
                TitulosNew.Add(tituloBxAd)
            End If
        End If

        '***********************************************************************************************************************************************************************************
        '**********************************************   Terceiro Passa de Previsao Pra Provisao   ****************************************************************************************
        '***********************************************************************************************************************************************************************************
        If vlr > 0 Then
            For Each tit In TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao And s.CodigoSituacao = eSituacao.Normal)
                If vlr > tit.Valores.EncargoValorDocumento.Valor Then
                    tit.IUD = "C"
                    tit.CodigoSituacao = eSituacao.Cancelado
                    vlr -= tit.Valores.EncargoValorDocumento.Valor

                    Dim titProv As Novo.TituloNovo = row.NF.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Provisao).FirstOrDefault
                    If titProv IsNot Nothing Then
                        titProv.IUD = "U"
                        titProv.Valores.EncargoValorDocumento.Valor = tit.Valores.EncargoValorDocumento.Valor
                        vlr -= tit.Valores.EncargoValorDocumento.Valor
                        titProv.Historico = "Titulo Provisionado pela Emissão da Nota/Serie " & row.NF.Codigo & "-" & row.NF.Serie & "  Referente ao pedido " & row.NF.CodigoPedido
                        TitulosOld.Remove(titProv)
                        TitulosNew.Add(titProv)
                    Else
                        'o campo valor em criatitulo e referencial por isso uso a variavel auxiliar
                        Dim VlrAux As Decimal = tit.Valores.EncargoValorDocumento.Valor
                        titProv = CriaTitulo(row, "PRO", VlrAux, tit.CodigoContaContabilRecPag)
                        vlr -= tit.Valores.EncargoValorDocumento.Valor
                        TitulosNew.Add(titProv)
                    End If
                Else
                    tit.Valores.EncargoValorDocumento.Valor -= vlr
                    Exit For
                End If
            Next
        End If

        Retorno.Add(True)
        Retorno.Add("")
        Return Retorno
    End Function

    Private Function ReconstroiLancamentosNotaFiscalDevolucao(row As ReconstroiFinanceiroPedido) As ArrayList
        '0 Verdadeiro Ou Falso
        '1 Msg - Sucesso ou Erro 
        Dim Retorno As New ArrayList

        'Valor Total da Nota
        Dim vlr As Decimal = row.Valor

        '***********************************************************************************************************************************************************************************
        '************************************************ Totalizadores dos titulos da NOTA ************************************************************************************************
        '***********************************************************************************************************************************************************************************
        'Totalizadores da lista de novos Titulos
        Dim VlrPrevisao As Decimal = TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao And s.CodigoSituacao = eSituacao.Normal).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
        Dim VlrProvisao As Decimal = TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao And s.CodigoSituacao = eSituacao.Normal).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)

        'Totalizadores dos Titulos da Nota Antiga
        Dim VlrNFCompensado As Decimal = row.NF.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Compensado).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
        Dim VlrNFAdiantamento As Decimal = row.NF.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa And s.isAdiantamento).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)

        'Valor Pra Voltar para Previsao
        Dim VlrVoltarPrevisao As Decimal = IIf(vlr > VlrProvisao, VlrProvisao, vlr)

        '************************************************************************************************************************************************************************************
        '********************************************************* Primeiro Mata o Provisao *************************************************************************************************
        '************************************************************************************************************************************************************************************
        Dim vlrComp As Decimal
        Dim vlrVoltarPrev As Decimal

        If VlrProvisao > 0 Then
            '********************************************************************************************************************************************************************************
            '******************************************* Faz A Compensacao na Nota Emitida  *************************************************************************************************
            '********************************************************************************************************************************************************************************
            Dim TitComp As Novo.TituloNovo

            If vlr > VlrProvisao Then
                vlrVoltarPrev = VlrProvisao
                vlrComp = VlrProvisao
                vlr -= VlrProvisao
            Else
                vlrVoltarPrev = vlr
                vlrComp = vlr
                vlr = 0
            End If

            If VlrNFCompensado > 0 Then
                TitComp = row.NF.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Compensado).FirstOrDefault
                TitComp.IUD = "U"
                TitComp.Valores.EncargoValorDocumento.Valor = vlrComp
                TitulosOld.Remove(TitulosOld.Where(Function(s) s.Codigo = TitComp.Codigo).FirstOrDefault)
            Else
                TitComp = CriaTitulo(row, "COMP", vlrComp)
            End If
            TitulosNew.Add(TitComp)

            '********************************************************************************************************************************************************************************
            '***************************** Diminui o valor em provisao e cria um titulo compensando nas Notas que tem Titulos em Provisão Emitida  ******************************************
            '********************************************************************************************************************************************************************************
            For Each tit In TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao And s.CodigoSituacao = eSituacao.Normal).OrderBy(Function(s) s.Reprogramacao)
                If tit.Valores.EncargoValorDocumento.Valor <= vlrComp Then
                    tit.IUD = "U"
                    tit.Reprogramacao = row.Lancamento
                    tit.CodigoProvisao = eProvisao.Compensado

                    vlrComp -= tit.Valores.EncargoValorDocumento.Valor
                ElseIf tit.Valores.EncargoValorDocumento.Valor > vlrComp Then
                    tit.IUD = "U"
                    tit.Valores.EncargoValorDocumento.Valor -= vlrComp
                    TitulosNew.Add(CriaTitulo(row, "COMP", vlrComp))
                    Exit For
                End If
            Next
        End If

        '************************************************************************************************************************************************************************************
        '********************************************************* Segundo Volta o Valor em Provisao para Previsao **************************************************************************
        '************************************************************************************************************************************************************************************
        If VlrProvisao > 0 Then
            Dim VlrParcelas As Decimal
            Dim VlrDiferenca As Decimal
            Dim NumTitulosEmPrevisaoAtivos As Decimal = TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao And s.CodigoSituacao = eSituacao.Normal).Count

            'Caso exista titulo em previsao ativos ele divide o valor do complemento entre eles
            If NumTitulosEmPrevisaoAtivos > 0 Then
                VlrParcelas = Math.Round(vlrVoltarPrev / NumTitulosEmPrevisaoAtivos, 2)
                VlrDiferenca = vlrVoltarPrev - (VlrParcelas * NumTitulosEmPrevisaoAtivos)

                For Each tit In TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao And s.CodigoSituacao = eSituacao.Normal)
                    tit.IUD = "U"
                    tit.Valores.EncargoValorDocumento.Valor += VlrParcelas + VlrDiferenca
                    VlrDiferenca = 0
                Next
            Else
                'caso NAO exista titulo em previsao ativos ele seleciona o Titulo em previsao Cancelado de Maior vencimento e volta ele para Ativo com o valor do complemento
                Dim tit As Novo.TituloNovo = TitulosNew.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao).OrderByDescending(Function(s) s.Reprogramacao).FirstOrDefault()
                tit.IUD = "U"
                tit.CodigoSituacao = eSituacao.Normal
                tit.Valores.EncargoValorDocumento.Valor = vlrVoltarPrev
            End If
        End If
        '************************************************************************************************************************************************************************************
        '****************************************************************** Depois cria Adiantamento   **************************************************************************************
        '************************************************************************************************************************************************************************************        
        If vlr > 0 Then
            If VlrNFAdiantamento > 0 Then
                Dim titad As Novo.TituloNovo = row.NF.Titulos.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Compensado).FirstOrDefault
                titad.IUD = "U"
                titad.Valores.EncargoValorDocumento.Valor = vlr
                TitulosNew.Add(titad)
                AdiantamentosAbertos.Add(titad.Codigo, vlr)
            Else
                TitulosNew.Add(CriaTitulo(row, "ADI", vlr))
                AdiantamentosAbertos.Add("NF" + row.NumeroDocumento, vlr)
            End If
        End If

        Retorno.Add(True)
        Retorno.Add("")
        Return Retorno
    End Function

    '***********************************************************************************************************
    '********** Recontroi Lancamentos Baixa Adiantamento Em Moeda **********************************************
    '***********************************************************************************************************
    Private Sub ReconstroiLancamentosBaixaDeAdiantamentos(row As ReconstroiFinanceiroPedido)
        Dim tit As Novo.TituloNovo = TitulosOld.Find(Function(s) s.Codigo = row.NumeroDocumento)
        Dim vlr As Decimal = tit.Valores.EncargoValorDocumento.Valor

        tit.IUD = "U"
        tit.AdiantamentosAbertos.Clear()
        For Each h In AdiantamentosAbertos
            If h.value = 0 Then Continue For

            Dim vlrbx As Decimal
            If h.value > vlr Then
                vlrbx = vlr
                h.value -= vlr
                vlr = 0
            Else
                vlrbx = h.value
                vlr -= h.value
                h.value = 0
            End If

            Dim adi As New Novo.AdiantamentoNovo()
            adi.CodigoTitulo = h.keys
            adi.AtualizarJuroEVariacao()

            adi.VlrBaixa = vlrbx
            tit.AdiantamentosAbertos.Add(adi)
            If vlr = 0 Then Exit For
        Next

        TitulosOld.Remove(TitulosOld.Where(Function(s) s.Codigo = tit.Codigo).FirstOrDefault)
        TitulosNew.Add(tit)
    End Sub

    '************************************************************************************************************
    '************************************ Cria Titulo ***********************************************************
    '************************************************************************************************************
    Private Function CriaTitulo(row As ReconstroiFinanceiroPedido, pTipo As String, ByRef pValor As Decimal, Optional pCodigoContaContabilRecPag As String = "") As Novo.TituloNovo
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
            tit.IUD = "I"
            tit.CodigoPedido = row.NF.CodigoPedido
            tit.Adiantamento = Nothing

            tit.DataBaixa = row.NF.Movimento
            tit.Vencimento = row.NF.Movimento
            tit.Reprogramacao = row.NF.Movimento
            tit.DataMoeda = row.NF.DataNota
            tit.CodigoTipoPgto = 1
            tit.CodigoCarteiraDoTitulo = 1

            tit.CodigoContaContabilCliFor = Pedido.SubOperacao.CodigoContaAdiantamento
            tit.CodigoContaContabilRecPag = Pedido.SubOperacao.CodigoGrupoContas
            tit.NotaTitulo = New Novo.NotaFiscalXTitulo(row.NF, tit)

            tit.ReceberPagar = IIf(Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida, "R", "P")
            tit.AdiantamentosAbertos.Clear()

            Dim lblTitulosBaixados As String = ""
            For Each h In AdiantamentosAbertos
                If h.value = 0 Then Continue For

                Dim vlrbx As Decimal
                If h.value > pValor Then
                    vlrbx = pValor
                    pValor = 0
                    h.value -= pValor
                Else
                    vlrbx = h.value
                    pValor -= h.value
                    h.value = 0
                End If

                Dim adi As New Novo.AdiantamentoNovo(h.keys)
                adi.VlrBaixa = vlrbx
                tit.AdiantamentosAbertos.Add(adi)
                lblTitulosBaixados += ", " & h.keys
            Next

            For Each Lvlr In tit.Valores
                If Not Lvlr.Equals(tit.Valores.EncargoValorDocumento) And Not Lvlr.Equals(tit.Valores.EncargoValorLiquido) Then
                    tit.Valores.Remove(Lvlr)
                End If
            Next

            tit.Valores.EncargoValorDocumento.DC = IIf(tit.Valores.EncargoValorDocumento.DC = "C", "D", "C")
            tit.Valores.EncargoValorLiquido.DC = IIf(tit.Valores.EncargoValorDocumento.DC = "C", "D", "C")

            tit.Valores.EncargoValorDocumento.Valor = pValor

            tit.Historico = "Baixa do Adiantamento" & lblTitulosBaixados & " do Pedido " & tit.CodigoPedido
            tit.Observacoes = ""
        End If

        If pTipo = "PRO" Then
            tit.IUD = "I"
            tit.Codigo = 0
            tit.CodigoProvisao = eProvisao.Provisao
            tit.CodigoPedido = row.NF.CodigoPedido

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

            tit.CodigoTipoPgto = 1
            tit.CodigoCarteiraDoTitulo = 1

            tit.Historico = "Titulo Provisionado pela Emissão da Nota/Serie " & row.NF.Codigo & "-" & row.NF.Serie & "  Referente ao pedido " & row.NF.CodigoPedido
            tit.Observacoes = ""

            tit.CodigoContaContabilCliFor = Pedido.SubOperacao.CodigoGrupoContas
            tit.CodigoContaContabilRecPag = pCodigoContaContabilRecPag

            tit.NotaTitulo = New Novo.NotaFiscalXTitulo(row.NF, tit)

            tit.Valores.EncargoValorDocumento.Valor = pValor
        End If

        If pTipo = "COMP" Then
            tit.IUD = "I"
            tit.Codigo = 0
            tit.CodigoProvisao = eProvisao.Compensado
            tit.CodigoPedido = row.NF.CodigoPedido

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

            tit.CodigoTipoPgto = 1
            tit.CodigoCarteiraDoTitulo = 1

            tit.Historico = "Titulo Compensado pela Emissão da Nota/Serie " & row.NF.Codigo & "-" & row.NF.Serie & "  Referente ao pedido " & row.NF.CodigoPedido
            tit.Observacoes = ""

            tit.CodigoContaContabilCliFor = Pedido.SubOperacao.CodigoGrupoContas
            tit.CodigoContaContabilRecPag = Pedido.Empresa.Empresa.CodigoContaGrupoBanco

            tit.NotaTitulo = New Novo.NotaFiscalXTitulo(row.NF, tit)

            tit.Valores.EncargoValorDocumento.Valor = pValor
        End If

        If pTipo = "ADI" Then
            tit.IUD = "I"
            tit.Codigo = 0
            tit.CodigoProvisao = eProvisao.Compensado
            tit.CodigoPedido = row.NF.CodigoPedido

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

            tit.CodigoTipoPgto = 1
            tit.CodigoCarteiraDoTitulo = 1

            tit.Historico = "ADIANTAMENTO REF. NF: " & NF.Codigo & " / " & NF.Cliente.Nome & "-" & IIf(NF.SubOperacao.EntradaSaida = eEntradaSaida.Saida, "Saida", "Entrada")
            tit.Observacoes = ""

            tit.CodigoContaContabilCliFor = Pedido.SubOperacao.CodigoContaAdiantamento
            tit.CodigoContaContabilRecPag = Pedido.SubOperacao.CodigoGrupoContas

            tit.NotaTitulo = New Novo.NotaFiscalXTitulo(row.NF, tit)

            tit.Valores.EncargoValorDocumento.Valor = pValor
        End If

        Return tit
    End Function

    '************************************************************************************************************
    '************************************ SALVAR SQL  ***********************************************************
    '************************************************************************************************************
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        '**********************************************************
        '************** LISTA ANTIGA DOS TITULOS  *****************
        '**********************************************************
        For Each TitOld In TitulosOld
            TitOld.IUD = "C"
            TitOld.SalvarSql(Sqls)
        Next

        '**********************************************************
        '************** LISTA NOVA DOS TITULOS  *****************
        '**********************************************************
        Dim i As Integer = 1
        Dim NumeroTitulo As Integer
        Dim Num As New Numerador(1)
        NumeroTitulo = Num.Sequencia

        For Each TitNew In TitulosNew
            If TitNew.IUD = "I" Then
                TitNew.Codigo = NumeroTitulo + i
                i += 1
            End If
            TitNew.SalvarSql(Sqls, False)

        Next
        Sqls.Add(Num.IncrementarNumeradorSql(True, i))

    End Sub
#End Region
End Class

'**************************************************************************************************************************
'**************************************  Reconstroi Financeiro Pedidos ****************************************************
'**************************************************************************************************************************
Public Class ReconstroiFinanceiroPedido
#Region "Contrutor"
    Public Sub New(pPedido As Pedido)
        _Pedido = pPedido       
    End Sub
#End Region

#Region "Fields"
    Private _Pedido As Pedido
    Private _NF As NotaFiscal

    Private _Documento As String
    Private _NumeroDocumento As Integer
    Private _Serie As String = ""
    Private _ES As String = ""
    Private _Lancamento As DateTime
    Private _CondicaoPagamento As Integer
    Private _Marcacao As String
    Private _ValorOficial As Decimal
    Private _ValorMoeda As Decimal
#End Region

#Region "Property"
    Public ReadOnly Property Pedido As Pedido
        Get
            Return _Pedido
        End Get
    End Property

    Public ReadOnly Property NF As NotaFiscal
        Get
            If Documento = "NF" AndAlso _NF Is Nothing Then
                Dim NfConsulta As New NotaFiscal
                NfConsulta.CodigoEmpresa = Pedido.CodigoEmpresa
                NfConsulta.EnderecoEmpresa = Pedido.EnderecoEmpresa
                NfConsulta.CodigoPedido = Pedido.Codigo
                NfConsulta.Codigo = NumeroDocumento
                NfConsulta.Serie = Serie
                NfConsulta.EntradaSaida = IIf(ES = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)

                _NF = New NotaFiscal(NfConsulta)
            End If
            Return _NF
        End Get
    End Property

    Public Property Documento As String
        Get
            Return _Documento
        End Get
        Set(value As String)
            _Documento = value
        End Set
    End Property

    Public Property NumeroDocumento As Integer
        Get
            Return _NumeroDocumento
        End Get
        Set(value As Integer)
            _NumeroDocumento = value
        End Set
    End Property

    Public Property Serie As String
        Get
            Return _Serie
        End Get
        Set(value As String)
            _Serie = value
        End Set
    End Property

    Public Property ES As String
        Get
            Return _ES
        End Get
        Set(value As String)
            _ES = value
        End Set
    End Property

    Public Property Lancamento As DateTime
        Get
            Return _Lancamento
        End Get
        Set(value As DateTime)
            _Lancamento = value
        End Set
    End Property

    Public Property CondicaoPagamento As Integer
        Get
            Return _CondicaoPagamento
        End Get
        Set(value As Integer)
            _CondicaoPagamento = value
        End Set
    End Property

    Public Property Marcacao As String
        Get
            Return _Marcacao
        End Get
        Set(value As String)
            _Marcacao = value
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

    'Segue a moeda do Pedido
    Public Property Valor As Decimal
        Get
            If Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Return ValorOficial
            Else
                Return ValorMoeda
            End If
        End Get
        Set(value As Decimal)
            If Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                ValorOficial = value
            Else
                ValorMoeda = value
            End If
        End Set
    End Property
#End Region
End Class
