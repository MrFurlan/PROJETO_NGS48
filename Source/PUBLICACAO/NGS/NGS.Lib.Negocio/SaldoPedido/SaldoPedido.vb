Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

Public Class ListSaldoPedido
    Inherits List(Of SaldoPedido)
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Parametros As Hashtable)
        '*******************************
        '***** PARAMETROS TRATATOS.*****
        '*******************************
        'CodEmpresa As String  
        'EndEmpresa As Integer 
        'CodCliente As String
        'EndCliente As Integer
        'ConsolidaCliente as boolean
        'DataInicial DataFinal As String - Trata-se da data em que o pedido foi realizado pode informar uma ou a outra caso informe as duas se torna um intervalo LOGICO
        'Safra As String
        'Pedido As String - Codigo do pedido
        'Situacao - 0 Todos, 1 com Saldo, 2 sem saldo
        'Operacao As String 
        'SubOperacao As String 
        'DataReferencia As String 
        'FiscalAberto As String - Situacao Fiscal do pedido

        Dim DataReferencia As String
        If Parametros.Contains("DataReferencia") Then
            DataReferencia = Parametros("DataReferencia")
        Else
            DataReferencia = Date.Now.ToString("yyyy-MM-dd")
        End If

        Dim sql As String = ""
        sql = "SELECT P.DataPedido, " & vbCrLf & _
                "       P.Empresa_Id, " & vbCrLf & _
                "       P.EndEmpresa_Id, " & vbCrLf & _
                "       P.Cliente, " & vbCrLf & _
                "       P.EndCliente, " & vbCrLf & _
                "       P.Pedido_Id, " & vbCrLf & _
                "       case when isnull(P.PedidoBloqueado,0) = 1 " & vbCrLf & _
                "	        then 0 " & vbCrLf & _
                "	        else " & vbCrLf & _
                "		        case when isnull(P.FiscalAberto,1) = 1 " & vbCrLf & _
                "			        then 1 " & vbCrLf & _
                "			        else 0 " & vbCrLf & _
                "		        end " & vbCrLf & _
                "       end as FiscalAberto, " & vbCrLf & _
                "       Produtos.Nome, " & vbCrLf & _
                "       pxi.Produto_Id, " & vbCrLf & _
                "       Produtos.Unidade, " & vbCrLf & _
                "       P.Operacao, " & vbCrLf & _
                "       P.SubOperacao, " & vbCrLf & _
                "       SO.Descricao, " & vbCrLf & _
                "       SO.PrecoFixo, " & vbCrLf & _
                "       P.Moeda, " & vbCrLf & _
                "       (case " & vbCrLf & _
                "          when SO.Classe in ('AFIXAR','DEPOSITOS') " & vbCrLf & _
                "             then pxi.UnitarioOficialNormal " & vbCrLf & _
                "          when pxi.QtdeProgramada = 0 " & vbCrLf & _
                "	         then 0 " & vbCrLf & _
                "	       else case " & vbCrLf & _
                "	              when P.Moeda = 1 " & vbCrLf & _
                "	                then pxi.PedidoValorOficial " & vbCrLf & _
                "	                else pxi.PedidoValorMoeda " & vbCrLf & _
                "	            end / (pxi.QtdeProgramada) " & vbCrLf & _
                "		end) as unitario, " & vbCrLf & _
                "        case  " & vbCrLf & _
                "          when SO.Classe in ('GLOBAL', 'REMESSAS') " & vbCrLf & _
                "            then 1 " & vbCrLf & _
                "          when SO.Classe in ('AFIXAR','COMPLEMENTACOES') " & vbCrLf & _
                "            then 2 " & vbCrLf & _
                "          when SO.Classe in ('DEPOSITOS') " & vbCrLf & _
                "            then 3 " & vbCrLf & _
                "          else 4 " & vbCrLf & _
                "        end Tipo, " & vbCrLf & _
                "        pxi.QtdeProgramada, " & vbCrLf & _
                "        pxi.PedidoValorOficial AS PedidoValor, " & vbCrLf & _
                "        convert(Decimal(18,4),0)  as QtdeFixacao, " & vbCrLf & _
                "        convert(Decimal(21,10),0) as UntFixacao, " & vbCrLf & _
                "        convert(Decimal(18,2),0)  as VlrFixacao " & vbCrLf & _
                "    into #Pedido " & vbCrLf & _
                "   FROM Pedidos P " & vbCrLf & _
                "  Inner Join(select Empresa_id, EndEmpresa_id, " & vbCrLf & _
                "                    Pedido_id, " & vbCrLf & _
                "                    Produto_Id, " & vbCrLf & _
                "                    sum(Case " & vbCrLf & _
                "                          when TipoDeLancamento = 'N' " & vbCrLf & _
                "                           then unitariooficial " & vbCrLf & _
                "                           else 0 " & vbCrLf & _
                "                        end) UnitarioOficialNormal, " & vbCrLf & _
                "                    sum(Case " & vbCrLf & _
                "                          when TipoDeLancamento = 'N' " & vbCrLf & _
                "                           then unitarioMoeda " & vbCrLf & _
                "                           else 0 " & vbCrLf & _
                "                        end) UnitarioMoedaNormal, " & vbCrLf & _
                "                    sum(case  " & vbCrLf & _
                "						  when TipoDeLancamento = 'E' " & vbCrLf & _
                "							then Quantidade * -1 " & vbCrLf & _
                "							else Quantidade " & vbCrLf & _
                "						end) as QtdeProgramada, " & vbCrLf & _
                "				    sum(case  " & vbCrLf & _
                "						  when TipoDeLancamento = 'E' " & vbCrLf & _
                "							then TotalOficial * -1  " & vbCrLf & _
                "							else TotalOficial " & vbCrLf & _
                "						end) as PedidoValorOficial, " & vbCrLf & _
                "				    sum(case  " & vbCrLf & _
                "						  when TipoDeLancamento = 'E' " & vbCrLf & _
                "							then TotalMoeda * -1  " & vbCrLf & _
                "							else TotalMoeda " & vbCrLf & _
                "						end) as PedidoValorMoeda " & vbCrLf & _
                "			   from PedidoxitemXLancamento " & vbCrLf & _
                "		      where Movimento           <= '" & DataReferencia & "'" & vbCrLf & _
                "		      group by Empresa_id, EndEmpresa_id, " & vbCrLf & _
                "                       Pedido_id, " & vbCrLf & _
                "                       Produto_Id " & vbCrLf & _
                "           ) pxi " & vbCrLf & _
                "		ON P.Empresa_Id    = pxi.Empresa_Id " & vbCrLf & _
                "	   AND P.EndEmpresa_Id = pxi.EndEmpresa_Id  " & vbCrLf & _
                "	   AND P.Pedido_Id     = pxi.Pedido_Id  " & vbCrLf & _
                "	 Inner Join SubOperacoes SO " & vbCrLf & _
                "		ON SO.Operacao_id     = P.Operacao " & vbCrLf & _
                "	   AND SO.SubOperacoes_id = P.SubOperacao " & vbCrLf & _
                "	 INNER JOIN Produtos  " & vbCrLf & _
                "		ON Produtos.Produto_Id = pxi.Produto_Id " & vbCrLf & _
                "	 INNER JOIN Moedas " & vbCrLf & _
                "		ON Moedas.Moeda_Id = P.Moeda " & vbCrLf & _
                "  Where P.Situacao = 1 " & vbCrLf

        If Parametros.ContainsKey("CodEmpresa") Then
            If Parametros("CodEmpresa").ToString.Length = 8 Then
                sql &= "        And left(P.Empresa_id,8) = '" & Parametros("CodEmpresa") & "'" & vbCrLf
            Else
                sql &= "        And P.Empresa_id = '" & Parametros("CodEmpresa") & "'" & vbCrLf
                If Parametros.ContainsKey("EndEmpresa") Then
                    sql &= "        and P.EndEmpresa_Id = " & Parametros("EndEmpresa") & vbCrLf
                End If
            End If
        End If

        If Parametros.ContainsKey("FiscalAberto") Then
            sql &= " And isnull(P.FiscalAberto,1) = " & IIf(Parametros("FiscalAberto") = "S", 1, 0) & vbCrLf
        End If

        If Parametros.ContainsKey("CodCliente") Then
            If Parametros("ConsolidaCliente") Then
                'Consolida clientes dos pedidos de entrada - Saida tem que ser exatamente para o mesmo cliente do pedido
                sql &= "        And left(P.Cliente,8) = '" & Left(Parametros("CodCliente"), 8) & "'" & vbCrLf
                sql &= "        And SO.Classe not in('AFIXAR') " & vbCrLf 'AFIXAR NÃO PODE SER UTILIZADO EM CONSOLIDAÇÃO - FURLAN - 25-04-2014
            Else
                sql &= "        And P.Cliente    ='" & Parametros("CodCliente") & "'" & vbCrLf & _
                       "        and P.EndCliente = " & Parametros("EndCliente") & vbCrLf
            End If
        End If

        If Parametros.ContainsKey("DataInicial") And Parametros.ContainsKey("DataFinal") Then
            sql &= " And P.DataPedido between '" & Parametros("DataInicial") & "' and '" & Parametros("DataFinal") & "'" & vbCrLf
        ElseIf Parametros.ContainsKey("DataInicial") Then
            sql &= " And P.DataPedido >= '" & Parametros("DataInicial") & "'" & vbCrLf
        ElseIf Parametros.ContainsKey("DataFinal") Then
            sql &= " And P.DataPedido <= '" & Parametros("DataFinal") & "'" & vbCrLf
        End If

        If Parametros.ContainsKey("Safra") Then
            sql &= " And P.Safra = '" & Parametros("Safra") & "'" & vbCrLf
        End If

        If Parametros.ContainsKey("Pedido") Then
            sql &= " And P.Pedido_Id = " & Parametros("Pedido") & vbCrLf
        Else
            If Parametros.ContainsKey("Operacao") Then
                sql &= " And P.Operacao = " & Parametros("Operacao") & vbCrLf
            End If

            If Parametros.ContainsKey("SubOperacao") Then
                sql &= "  And P.SubOperacao = " & Parametros("SubOperacao") & vbCrLf
            End If
        End If

        '******************************************************************************************************************************
        '******************************************************************************************************************************
        '******************************************************************************************************************************
        sql &= ";update #Pedido set" & vbCrLf & _
               "      QtdeFixacao = fix.Quantidade" & vbCrLf & _
               "     ,VlrFixacao  = fix.TotalOficial" & vbCrLf & _
               "     ,UntFixacao  = fix.Unitario" & vbCrLf & _
               "   from #Pedido" & vbCrLf & _
               "  INNER JOIN (" & vbCrLf & _
               "               select Pxi.Empresa_id, Pxi.EndEmpresa_id, Pxi.Pedido_Id, Pxi.Produto_id," & vbCrLf & _
               "                      Sum(Pxi.Quantidade) as Quantidade, sum(Pxi.TotalOficial) as TotalOficial," & vbCrLf & _
               "                      sum(Pxi.TotalOficial) / (Sum(Pxi.Quantidade)) as unitario" & vbCrLf & _
               "                 from VW_pedidosxitensxfixacoes Pxi" & vbCrLf & _
               "                inner join Produtos p" & vbCrLf & _
               "                   on p.Produto_Id = Pxi.produto_Id" & vbCrLf & _
               "                Group by Pxi.Empresa_id, Pxi.EndEmpresa_id, Pxi.Pedido_Id, Pxi.Produto_id" & vbCrLf & _
               "              ) fix" & vbCrLf & _
               "      ON #Pedido.Empresa_Id    = fix.Empresa_Id" & vbCrLf & _
               "     AND #Pedido.EndEmpresa_Id = fix.EndEmpresa_Id" & vbCrLf & _
               "     AND #Pedido.Pedido_Id     = fix.Pedido_Id" & vbCrLf & _
               "     AND #Pedido.Produto_Id    = fix.Produto_id" & vbCrLf & _
               "     AND fix.Quantidade        > 0" & vbCrLf & _
               "   Where #Pedido.PrecoFixo = 'N'" & vbCrLf & vbCrLf


        '******************************************************************************************************************************
        '******************************************************************************************************************************
        '******************************************************************************************************************************


        sql &= " Select P.DataPedido, " & vbCrLf & _
               "        P.Empresa_Id,     P.EndEmpresa_Id," & vbCrLf & _
               "        P.Cliente,        P.EndCliente," & vbCrLf & _
               "        P.Pedido_Id," & vbCrLf & _
               "        P.FiscalAberto, " & vbCrLf & _
               "	    P.Produto_Id,     P.Nome,             P.Unidade," & vbCrLf & _
               "        P.Operacao,       P.SubOperacao, " & vbCrLf & _
               "        P.Descricao, " & vbCrLf & _
               "        isnull(Sb_Nf.tipo,P.Tipo) as tipo," & vbCrLf & _
               "        P.QtdeFixacao, P.VlrFixacao," & vbCrLf & _
               "        P.QtdeProgramada, P.unitario, P.PedidoValor," & vbCrLf & _
               "                                                         " & vbCrLf & _
               "        case isnull(Sb_Nf.tipo,P.Tipo)" & vbCrLf & _
               "          when 1 then isnull(Sb_Nf.GlobalNormalFiscal - sb_Nf.GlobalDevolucaoFiscal  ,0) " & vbCrLf & _
               "          when 2 then case " & vbCrLf & _
               "                        when P.QtdeProgramada = 0 or P.QtdeProgramada < isnull(Sb_Nf.AFixarNormalFiscal ,0) " & vbCrLf & _
               "                          then isnull(Sb_Nf.AFixarNormalFiscal ,0) - isnull(Sb_Nf.AFixarDevolucaoFiscal,0)" & vbCrLf & _
               "                          else P.QtdeProgramada " & vbCrLf & _
               "                      end " & vbCrLf & _
               "          when 3 then case " & vbCrLf & _
               "                        when P.QtdeProgramada = 0 " & vbCrLf & _
               "                          then isnull(Sb_Nf.DepositoFiscal ,0) " & vbCrLf & _
               "                          else P.QtdeProgramada " & vbCrLf & _
               "                      end " & vbCrLf & _
               "          else P.QtdeProgramada" & vbCrLf & _
               "        end as Contratado, " & vbCrLf & _
               "                                                         " & vbCrLf & _
               "        case isnull(Sb_Nf.tipo,P.Tipo)" & vbCrLf & _
               "          when 1 then isnull(Sb_Nf.GlobalNormalFiscal - sb_Nf.GlobalDevolucaoFiscal  ,0) " & vbCrLf & _
               "          when 2 then case " & vbCrLf & _
               "                        when P.QtdeProgramada = 0 or P.QtdeProgramada < isnull(Sb_Nf.AFixarNormalFisica ,0) " & vbCrLf & _
               "                         then isnull(Sb_Nf.AFixarNormalFisica ,0) - isnull(Sb_Nf.AFixarDevolucaoFisica ,0) " & vbCrLf & _
               "                         else P.QtdeProgramada " & vbCrLf & _
               "                      end " & vbCrLf & _
               "          when 3 then case " & vbCrLf & _
               "                        when P.QtdeProgramada = 0 " & vbCrLf & _
               "                          then isnull(Sb_Nf.DepositoFisica ,0) " & vbCrLf & _
               "                          else P.QtdeProgramada " & vbCrLf & _
               "                      end " & vbCrLf & _
               "          else P.QtdeProgramada" & vbCrLf & _
               "        end as ContratadoFisico, " & vbCrLf & _
               "                                                         " & vbCrLf & _
               "        case isnull(Sb_Nf.tipo,P.Tipo)" & vbCrLf & _
               "          when 1 then isnull(Sb_Nf.RemessaNormalFiscal,0)  - isnull(Sb_Nf.RemessaDevolucaoFiscal ,0)" & vbCrLf & _
               "          when 2 then isnull(Sb_Nf.AFixarNormalFiscal,0)   - isnull(Sb_Nf.AFixarDevolucaoFiscal  ,0)" & vbCrLf & _
               "          when 3 then isnull(Sb_Nf.DepositoFiscal ,0) - isnull(Sb_Nf.DepositoDevolucaoFiscal ,0) --isnull(Sb_Nf.DepositoDevolucaoFiscal ,0)" & vbCrLf & _
               "          else isnull(Sb_Nf.NormalOperacaoFiscal,0) - isnull(Sb_Nf.NormalOperacaoDevolucaoFiscal,0)" & vbCrLf & _
               "        end as EntregueFiscal, " & vbCrLf & _
               "                                                         " & vbCrLf & _
               "        case isnull(Sb_Nf.tipo,P.Tipo)" & vbCrLf & _
               "          when 1 then isnull(Sb_Nf.RemessaNormalFisica,0) - isnull(Sb_Nf.RemessaDevolucaoFisica ,0)" & vbCrLf & _
               "          when 2 then isnull(Sb_Nf.AFixarNormalFisica,0)  - isnull(Sb_Nf.AFixarDevolucaoFisica ,0)" & vbCrLf & _
               "          when 3 then isnull(Sb_Nf.DepositoFisica ,0) - isnull(Sb_Nf.DepositoDevolucaoFisica ,0) --isnull(Sb_Nf.DepositoDevolucaoFisica ,0) " & vbCrLf & _
               "          else isnull(Sb_Nf.NormalOperacaoFisica,0) - isnull(Sb_Nf.NormalOperacaoDevolucaoFisica,0)" & vbCrLf & _
               "        end as EntregueFisico," & vbCrLf & _
               "                                                         " & vbCrLf & _
               "        case isnull(Sb_Nf.tipo,P.Tipo)" & vbCrLf & _
               "          when 1 then isnull(Sb_Nf.RemessaValor - Sb_Nf.RemessaValorDevolucao,0)" & vbCrLf & _
               "          when 2 then isnull((Sb_Nf.AFixarValor  - Sb_Nf.AFixarValorDevolucao) + (Sb_Nf.ComplementacoesValor - Sb_Nf.ComplementacoesValorDevolucao),0)" & vbCrLf & _
               "          when 3 then isnull(Sb_Nf.DepositoValor - Sb_Nf.DepositoDevolucaoValor,0)" & vbCrLf & _
               "          else isnull(Sb_Nf.NormalOperacaoValor - sb_nf.NormalOperacaoValorDevolucao,0)" & vbCrLf & _
               "        end as NotaValor," & vbCrLf & _
               "                                                         " & vbCrLf & _
               "        isnull(Sb_Nf.GlobalNormalFiscal,0)            as GlobalNormal," & vbCrLf & _
               "        isnull(Sb_Nf.GlobalDevolucaoFiscal,0)         as GlobalDevolucao," & vbCrLf & _
               "        isnull(Sb_Nf.RemessaNormalFiscal,0)           as RemessaNormal," & vbCrLf & _
               "        isnull(Sb_Nf.RemessaDevolucaoFiscal,0)        as RemessaDevolucao," & vbCrLf & _
               "        isnull(Sb_Nf.DepositoFiscal,0)                as DepositadoFiscal," & vbCrLf & _
               "        isnull(Sb_Nf.DepositoFisica,0)                as DepositadoFisico," & vbCrLf & _
               "        isnull(Sb_Nf.DepositoDevolucaoFiscal,0)       as EntregueDeposito," & vbCrLf & _
               "        isnull(Sb_Nf.NormalOperacaoFiscal,0)          as OperacaoNormal," & vbCrLf & _
               "        isnull(Sb_Nf.NormalOperacaoDevolucaoFiscal,0) as OperacaoDevolucao" & vbCrLf & _
               "   into #PreSaldo" & vbCrLf & _
               "   from #Pedido P" & vbCrLf & _
               "   LEFT JOIN (" & vbCrLf & _
               "			  SELECT NotasFiscais.Empresa_Id," & vbCrLf & _
               "			  	     NotasFiscais.EndEmpresa_Id," & vbCrLf & _
               "					 NotasFiscais.Pedido," & vbCrLf & _
               "					 NxI.Produto_Id," & vbCrLf & _
               "					 case " & vbCrLf & _
               "					    when sb.Classe in ('GLOBAL', 'REMESSAS')" & vbCrLf & _
               "						  then 1" & vbCrLf & _
               "					    When sb.Classe in ('AFIXAR','COMPLEMENTACOES')" & vbCrLf & _
               "						  then 2" & vbCrLf & _
               "					    when sb.Classe in ('DEPOSITOS')" & vbCrLf & _
               "						  then 3" & vbCrLf & _
               "					    else 4" & vbCrLf & _
               "					 end Tipo," & vbCrLf & _
               "					 sum(case When sb.Classe = 'GLOBAL'          and sb.Devolucao = 'N' then NxI.QuantidadeFiscal Else 0 end) As GlobalNormalFiscal," & vbCrLf & _
               "					 sum(case When sb.Classe = 'GLOBAL'          and sb.Devolucao = 'S' then NxI.QuantidadeFiscal Else 0 end) As GlobalDevolucaoFiscal," & vbCrLf & _
               "					 Sum(case when sb.Classe = 'GLOBAL'          and Sb.Devolucao = 'N' then NxI.Valor            Else 0 end) As GlobalValor," & vbCrLf & _
               "					 Sum(case when sb.Classe = 'GLOBAL'          and Sb.Devolucao = 'S' then NxI.Valor            Else 0 end) As GlobalValorDevolucao," & vbCrLf & _
               "					 sum(case When sb.Classe = 'REMESSAS'        and Sb.Devolucao = 'N' then NxI.QuantidadeFiscal Else 0 end) As RemessaNormalFiscal," & vbCrLf & _
               "					 sum(case When sb.Classe = 'REMESSAS'        and Sb.Devolucao = 'S' then NxI.QuantidadeFiscal Else 0 end) As RemessaDevolucaoFiscal," & vbCrLf & _
               "					 sum(case When sb.Classe = 'REMESSAS'        and Sb.Devolucao = 'N' then NxI.QuantidadeFisica Else 0 end) As RemessaNormalFisica," & vbCrLf & _
               "					 sum(case When sb.Classe = 'REMESSAS'        and Sb.Devolucao = 'S' then NxI.QuantidadeFisica Else 0 end) As RemessaDevolucaoFisica," & vbCrLf & _
               "					 sum(case When sb.Classe = 'REMESSAS'        and Sb.Devolucao = 'N' then NxI.Valor            Else 0 end) As RemessaValor," & vbCrLf & _
               "					 sum(case When sb.Classe = 'REMESSAS'        and Sb.Devolucao = 'S' then NxI.Valor            Else 0 end) As RemessaValorDevolucao," & vbCrLf & _
               "					 sum(case When sb.Classe = 'AFIXAR'          and sb.Devolucao = 'N' then NxI.QuantidadeFiscal Else 0 end) As AFixarNormalFiscal," & vbCrLf & _
               "					 sum(case When sb.Classe = 'AFIXAR'          and sb.Devolucao = 'S' then NxI.QuantidadeFiscal Else 0 end) As AFixarDevolucaoFiscal," & vbCrLf & _
               "					 sum(case When sb.Classe = 'AFIXAR'          and sb.Devolucao = 'N' then NxI.QuantidadeFisica Else 0 end) As AFixarNormalFisica," & vbCrLf & _
               "					 sum(case When sb.Classe = 'AFIXAR'          and sb.Devolucao = 'S' then NxI.QuantidadeFisica Else 0 end) As AFixarDevolucaoFisica," & vbCrLf & _
               "					 Sum(case when sb.Classe = 'AFIXAR'          and Sb.Devolucao = 'N' then NxI.Valor            Else 0 end) As AFixarValor," & vbCrLf & _
               "					 Sum(case when sb.Classe = 'AFIXAR'          and Sb.Devolucao = 'S' then NxI.Valor            Else 0 end) As AFixarValorDevolucao," & vbCrLf & _
               "					 sum(case When sb.Classe = 'COMPLEMENTACOES' and Sb.Devolucao = 'N' then NxI.Valor            Else 0 end) As ComplementacoesValor," & vbCrLf & _
               "					 sum(case When sb.Classe = 'COMPLEMENTACOES' and Sb.Devolucao = 'S' then NxI.Valor            Else 0 end) As ComplementacoesValorDevolucao," & vbCrLf & _
               "					 sum(case When sb.Classe = 'DEPOSITOS'       and Sb.Devolucao = 'N' then NxI.QuantidadeFiscal else 0 end) As DepositoFiscal," & vbCrLf & _
               "					 sum(case When sb.Classe = 'DEPOSITOS'       and Sb.Devolucao = 'S' then NxI.QuantidadeFiscal else 0 end) As DepositoDevolucaoFiscal," & vbCrLf & _
               "					 sum(case When sb.Classe = 'DEPOSITOS'       and Sb.Devolucao = 'N' then NxI.QuantidadeFisica else 0 end) As DepositoFisica," & vbCrLf & _
               "					 sum(case When sb.Classe = 'DEPOSITOS'       and Sb.Devolucao = 'S' then NxI.QuantidadeFisica else 0 end) As DepositoDevolucaoFisica," & vbCrLf & _
               "					 sum(case When sb.Classe = 'DEPOSITOS'       and Sb.Devolucao = 'N' then NxI.Valor            else 0 end) As DepositoValor," & vbCrLf & _
               "					 sum(case When sb.Classe = 'DEPOSITOS'       and Sb.Devolucao = 'S' then NxI.Valor            else 0 end) As DepositoDevolucaoValor," & vbCrLf & _
               "					 sum(case When sb.classe not in ('GLOBAL','REMESSAS','AFIXAR','COMPLEMENTACOES','DEPOSITOS') and sb.Devolucao = 'N' then NxI.QuantidadeFiscal else 0 end) As NormalOperacaoFiscal," & vbCrLf & _
               "					 sum(case When sb.classe not in ('GLOBAL','REMESSAS','AFIXAR','COMPLEMENTACOES','DEPOSITOS') and Sb.Devolucao = 'N' then NxI.QuantidadeFisica else 0 end) As NormalOperacaoFisica," & vbCrLf & _
               "					 sum(case when sb.classe not in ('GLOBAL','REMESSAS','AFIXAR','COMPLEMENTACOES','DEPOSITOS') and sb.Devolucao = 'N' then NxI.Valor            else 0 end) As NormalOperacaoValor," & vbCrLf & _
               "					 sum(case when sb.classe not in ('GLOBAL','REMESSAS','AFIXAR','COMPLEMENTACOES','DEPOSITOS') and sb.Devolucao = 'S' then NxI.QuantidadeFiscal else 0 end) As NormalOperacaoDevolucaoFiscal," & vbCrLf & _
               "					 sum(case when sb.classe not in ('GLOBAL','REMESSAS','AFIXAR','COMPLEMENTACOES','DEPOSITOS') and sb.Devolucao = 'S' then NxI.QuantidadeFisica else 0 end) As NormalOperacaoDevolucaoFisica," & vbCrLf & _
               "					 sum(case when sb.classe not in ('GLOBAL','REMESSAS','AFIXAR','COMPLEMENTACOES','DEPOSITOS') and sb.Devolucao = 'S' then NxI.Valor            else 0 end) As NormalOperacaoValorDevolucao" & vbCrLf & _
               "                FROM NotasFiscais  " & vbCrLf & _
               "		       INNER JOIN NotasFiscaisXItens NxI " & vbCrLf & _
               "				  ON NotasFiscais.Empresa_Id      = NxI.Empresa_Id " & vbCrLf & _
               "				 AND NotasFiscais.EndEmpresa_Id   = NxI.EndEmpresa_Id  " & vbCrLf & _
               "				 AND NotasFiscais.Cliente_Id      = NxI.Cliente_Id  " & vbCrLf & _
               "				 AND NotasFiscais.EndCliente_Id   = NxI.EndCliente_Id  " & vbCrLf & _
               "				 AND NotasFiscais.EntradaSaida_Id = NxI.EntradaSaida_Id  " & vbCrLf & _
               "				 AND NotasFiscais.Serie_Id        = NxI.Serie_Id  " & vbCrLf & _
               "				 AND NotasFiscais.Nota_Id         = NxI.Nota_Id  " & vbCrLf & _
               "			--  LEFT JOIN NotasXDestinos NxD --#NFD 08 " & vbCrLf & _
               "			--	  ON NxI.Empresa_Id      = NxD.Empresa_Id " & vbCrLf & _
               "			--	 AND NxI.EndEmpresa_Id   = NxD.EndEmpresa_Id " & vbCrLf & _
               "			--	 AND NxI.Cliente_Id      = NxD.Cliente_Id " & vbCrLf & _
               "			--	 AND NxI.EndCliente_Id   = NxD.EndCliente_Id" & vbCrLf & _
               "			--	 AND NxI.EntradaSaida_Id = NxD.EntradaSaida_Id " & vbCrLf & _
               "			--	 AND NxI.Serie_Id        = NxD.Serie_Id " & vbCrLf & _
               "			--	 AND NxI.Nota_Id         = NxD.Nota_Id " & vbCrLf & _
               "			--	 AND NxI.Produto_Id      = NxD.Produto_Id" & vbCrLf & _
               "			   INNER JOIN SubOperacoes as Sb " & vbCrLf & _
               "				  ON NotasFiscais.Operacao    = Sb.Operacao_Id  " & vbCrLf & _
               "				 AND NotasFiscais.SubOperacao = Sb.SubOperacoes_Id " & vbCrLf & _
               "			   INNER JOIN Pedidos P" & vbCrLf & _
               "			 	  ON P.Empresa_Id    = NotasFiscais.Empresa_Id " & vbCrLf & _
               "				 AND P.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id " & vbCrLf & _
               "				 AND P.Pedido_Id     = NotasFiscais.Pedido" & vbCrLf & _
               "		       Where NotasFiscais.Situacao in (1,4,7)" & vbCrLf & _
               "			     and NotasFiscais.TipoDeDocumento = 1 " & vbCrLf & _
               "			     and NotasFiscais.Movimento <= '" & DataReferencia & "'" & vbCrLf & _
               "		       Group by NotasFiscais.Empresa_Id,  " & vbCrLf & _
               "	  				    NotasFiscais.EndEmpresa_Id, NotasFiscais.Pedido, NxI.Produto_Id," & vbCrLf & _
               "						case " & vbCrLf & _
               "						   when sb.Classe in ('GLOBAL', 'REMESSAS')" & vbCrLf & _
               "							 then 1" & vbCrLf & _
               "						   When sb.Classe in ('AFIXAR','COMPLEMENTACOES')" & vbCrLf & _
               "							 then 2" & vbCrLf & _
               "						   when sb.Classe in ('DEPOSITOS')" & vbCrLf & _
               "							 then 3" & vbCrLf & _
               "						   else 4" & vbCrLf & _
               "						end" & vbCrLf & _
               "		    ) AS Sb_Nf" & vbCrLf & _
               "	 on Sb_Nf.Empresa_Id    = P.Empresa_Id " & vbCrLf & _
               "    and Sb_Nf.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
               "    and Sb_Nf.Pedido        = P.Pedido_Id " & vbCrLf & _
               "    and Sb_Nf.Produto_Id    = P.Produto_Id;" & vbCrLf & vbCrLf

        '*********************************************************************************************************************************************************************
        '*********************************************************************************************************************************************************************
        '*********************************************************************************************************************************************************************


        sql &= " select Operacao, SubOperacao," & vbCrLf & _
               "        Empresa_Id, EndEmpresa_Id," & vbCrLf & _
               "        Cliente, EndCliente, " & vbCrLf & _
               "        convert(nvarchar(10),DataPedido,103) as DataPedido," & vbCrLf & _
               "        Pedido_Id, FiscalAberto, Tipo, Produto_Id,Unitario, " & vbCrLf & _
               "        max(QtdeFixacao) as QtdeFixacao, " & vbCrLf & _
               "        max(VlrFixacao)  as VlrFixacao, " & vbCrLf & _
               "        max(QtdeProgramada)      as QtdeProgramada," & vbCrLf & _
               "        Max(PedidoValor)         as PedidoValor," & vbCrLf & _
               "        Max(Contratado)          as Contratado," & vbCrLf & _
               "        Max(ContratadoFisico)    as ContratadoFisico," & vbCrLf & _
               "        sum(EntregueFiscal)      as EntregueFiscal, " & vbCrLf & _
               "        convert(numeric(18,4),0) as SaldoFiscal," & vbCrLf & _
               "        sum(EntregueFisico)      as EntregueFisico," & vbCrLf & _
               "        convert(numeric(18,4),0) as SaldoFisico," & vbCrLf & _
               "        sum(NotaValor)           as NotaValor, " & vbCrLf & _
               "        convert(numeric(18,2),0) as SaldoValor,  " & vbCrLf & _
               "        convert(numeric(18,6),0) as UnitarioValor" & vbCrLf & _
               "   Into #PreSaldoPedido " & vbCrLf & _
               "   from #PreSaldo" & vbCrLf & _
               "  group by Empresa_Id, EndEmpresa_Id, Cliente, EndCliente, convert(nvarchar(10),DataPedido,103), Pedido_Id, FiscalAberto," & vbCrLf & _
               "           Tipo, Operacao, SubOperacao, Produto_Id, Unitario " & vbCrLf

        'sql &= " update #PreSaldoPedido set" & vbCrLf & _
        '       "  SaldoFiscal   = case when tipo = 2 then EntregueFiscal - QtdeFixacao else Contratado       - EntregueFiscal end," & vbCrLf & _
        '       "  SaldoFisico   = case when tipo = 2 then EntregueFisico - QtdeFixacao else ContratadoFisico - EntregueFisico end," & vbCrLf & _
        '       "  SaldoValor    = case when tipo = 2 then NotaValor      - VlrFixacao  else PedidoValor      - NotaValor end," & vbCrLf & _
        '       "  UnitarioValor = 0" & vbCrLf & vbCrLf

        sql &= " Update #PreSaldoPedido set" & vbCrLf & _
               " SaldoFiscal = case" & vbCrLf & _
               "                 when tipo = 2 then EntregueFiscal - QtdeFixacao" & vbCrLf & _
               "                 when tipo = 3 then case" & vbCrLf & _
               "                                      when Contratado > 0" & vbCrLf & _
               "                                        then Contratado - EntregueFiscal" & vbCrLf & _
               "                                        else EntregueFiscal" & vbCrLf & _
               "                                    end" & vbCrLf & _
               "                 else Contratado       - EntregueFiscal end," & vbCrLf & _
               " SaldoFisico = case" & vbCrLf & _
               "                 when tipo = 2 then EntregueFisico - QtdeFixacao" & vbCrLf & _
               "                 when tipo = 3 then Case" & vbCrLf & _
               "                                      When ContratadoFisico > 0" & vbCrLf & _
               "                                        then ContratadoFisico - EntregueFisico" & vbCrLf & _
               "                                        else EntregueFisico" & vbCrLf & _
               "                                    end" & vbCrLf & _
               "                 else ContratadoFisico - EntregueFisico end," & vbCrLf & _
               " SaldoValor  = case" & vbCrLf & _
               "                   when tipo = 2" & vbCrLf & _
               "                     then NotaValor   - VlrFixacao" & vbCrLf & _
               "                     else PedidoValor - NotaValor" & vbCrLf & _
               "               end," & vbCrLf & _
               " UnitarioValor = 0" & vbCrLf

        sql &= " select Operacao,SubOperacao,Empresa_Id, EndEmpresa_Id, Cliente, EndCliente, convert(nvarchar(10),DataPedido,103) as DataPedido," & vbCrLf & _
               "        Pedido_Id, FiscalAberto, " & vbCrLf & _
               "        max(QtdeProgramada)   as QtdeProgramada," & vbCrLf & _
               "        sum(Contratado)       as ContratadoFiscal," & vbCrLf & _
               "        sum(ContratadoFisico) as ContratadoFisico," & vbCrLf & _
               "        sum(PedidoValor)      as PedidoValor," & vbCrLf & _
               "        sum(EntregueFiscal)   as EntregueFiscal, " & vbCrLf & _
               "        sum(SaldoFiscal)      as SaldoFiscal," & vbCrLf & _
               "        sum(EntregueFisico)   as EntregueFisico," & vbCrLf & _
               "        sum(SaldoFisico)      as SaldoFisico," & vbCrLf & _
               "        sum(NotaValor)        as NotaValor, " & vbCrLf & _
               "        Sum(SaldoValor)       as SaldoValor" & vbCrLf & _
               "   from #PreSaldoPedido" & vbCrLf & _
               "  group by Empresa_Id, EndEmpresa_Id, Cliente, EndCliente, convert(nvarchar(10),DataPedido,103), Pedido_Id, FiscalAberto, Operacao, SubOperacao" & vbCrLf

        If Parametros.ContainsKey("Situacao") Then
            Select Case Parametros("Situacao")
                Case 1
                    sql &= " Having sum(SaldoFiscal) > 0 or  sum(SaldoFisico) > 0 or Sum(SaldoValor) > 0" & vbCrLf & vbCrLf
                Case 2
                    sql &= " Having sum(SaldoFiscal) = 0 and sum(SaldoFisico) = 0 or Sum(SaldoValor) = 0" & vbCrLf & vbCrLf
            End Select
        End If


        sql &= " drop table #Pedido;" & vbCrLf & _
               " drop table #PreSaldo;" & vbCrLf & _
               " drop table #PreSaldoPedido;"

        Dim ds As DataSet
        Dim banco As New AcessaBanco
        ds = banco.ConsultaDataSet(sql, "Saldo")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim SP As New SaldoPedido
                SP.CodigoEmpresa = row("Empresa_Id") 'ok
                SP.EndEmpresa = row("EndEmpresa_Id") 'ok
                SP.CodigoCliente = row("Cliente") 'ok
                SP.EndCliente = row("EndCliente") 'ok
                SP.CodigoPedido = row("Pedido_id") 'ok
                SP.FiscalAberto = row("FiscalAberto") 'ok
                SP.DataPedido = row("DataPedido") 'ok
                SP.CodigoOperacao = row("Operacao") 'ok
                SP.CodigoSubOperacao = row("SubOperacao") 'ok

                SP.QtdeProgramadaPedido = row("QtdeProgramada") 'ok

                SP.Contratado = row("Contratadofiscal") 'ok
                SP.ContratadoFisico = row("ContratadoFisico") 'ok

                SP.EntregueFiscal = row("EntregueFiscal") 'ok
                SP.EntregueFisico = row("EntregueFisico") 'ok

                SP.SaldoFiscal = row("SaldoFiscal") 'ok
                SP.SaldoFisico = row("SaldoFisico") 'ok

                SP.PedidoValor = row("PedidoValor") 'ok
                SP.NotaValor = row("NotaValor") 'ok
                SP.SaldoValor = row("SaldoValor") 'ok
                Me.Add(SP)
            Next
        End If
    End Sub
#End Region

#Region "Fields"
    Private _ProdutosPedido As String
#End Region

#Region "Property"
    Public ReadOnly Property ProdutosPedido() As String
        Get
            Return _ProdutosPedido
        End Get
    End Property
#End Region

#Region "Methods"

#End Region

End Class

Public Class SaldoPedido
    Implements IBaseEntity

#Region "Fields"
    Private _Selecionado As Boolean = False
    Private _Empresa As Cliente
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer

    Private _Cliente As Cliente
    Private _CodigoCliente As String
    Private _EndCliente As Integer

    Private _CodigoPedido As String
    Private _FiscalAberto As Boolean
    Private _CodigoOperacao As Integer
    Private _CodigoSubOperacao As Integer
    Private _SubOperacao As SubOperacao

    Private _DescricaoEmpresa As String
    Private _DescricaoCliente As String
    Private _Pedido As Pedido
    Private _DataPedido As Date

    Private _ContratadoFiscal As Decimal
    Private _EntregueFiscal As Decimal
    Private _SaldoFiscal As Decimal

    Private _ContratadoFisico As Decimal
    Private _EntregueFisico As Decimal
    Private _SaldoFisico As Decimal

    Private _QtdeProgramadaPedido As Decimal
    Private _PedidoValor As Decimal
    Private _NotaValor As Decimal
    Private _SaldoValor As Decimal

    Private _Itens As ListSaldoPedidoxItem
#End Region

#Region "Property"
    Public Property Selecionado() As Boolean
        Get
            Return _Selecionado
        End Get
        Set(ByVal value As Boolean)
            _Selecionado = value
        End Set
    End Property

    '*************** PEDIDO ********************************
    Public Property CodigoPedido() As Integer
        Get
            Return _CodigoPedido
        End Get
        Set(ByVal value As Integer)
            _CodigoPedido = value
            _Pedido = Nothing
        End Set
    End Property 'ok

    Public Property Pedido() As Pedido
        Get
            If _Pedido Is Nothing And _CodigoPedido > 0 Then _Pedido = New Pedido(_CodigoEmpresa, _EndEmpresa, _CodigoPedido)
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
        End Set
    End Property 'ok

    Public Property FiscalAberto() As Boolean
        Get
            Return _FiscalAberto
        End Get
        Set(ByVal value As Boolean)
            _FiscalAberto = value
        End Set
    End Property 'ok

    Public Property DataPedido() As Date
        Get
            Return _DataPedido
        End Get
        Set(ByVal value As Date)
            _DataPedido = value
        End Set
    End Property 'ok

    '**************** OPERACAO *****************************
    Public Property CodigoOperacao() As Integer
        Get
            Return _CodigoOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoOperacao = value
            _SubOperacao = Nothing
        End Set
    End Property 'ok

    Public Property CodigoSubOperacao() As Integer
        Get
            Return _CodigoSubOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSubOperacao = value
            _SubOperacao = Nothing
        End Set
    End Property 'ok

    Public ReadOnly Property SubOperacao As SubOperacao
        Get
            If _SubOperacao Is Nothing And Me.CodigoOperacao > 0 And Me.CodigoSubOperacao > 0 Then _SubOperacao = New SubOperacao(Me.CodigoOperacao, Me.CodigoSubOperacao)
            Return _SubOperacao
        End Get
    End Property

    '**************** EMPRESA ******************************
    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            _Empresa = Nothing
            _DescricaoEmpresa = Nothing
        End Set
    End Property 'ok

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
            _Empresa = Nothing
            _DescricaoEmpresa = Nothing
        End Set
    End Property 'ok

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EndEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property 'ok

    Public ReadOnly Property DescricaoEmpresa()
        Get
            If _DescricaoEmpresa Is Nothing And Not Empresa Is Nothing Then
                _DescricaoEmpresa = Funcoes.AlinharEsquerda(Empresa.Nome, 28, ".") & " - " & Funcoes.AlinharEsquerda(Empresa.Cidade, 20, ".") & " " & Empresa.CodigoEstado

            End If
            Return _DescricaoEmpresa
        End Get
    End Property 'ok

    '**************** CLIENTE ******************************
    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            _Cliente = Nothing
            _DescricaoCliente = Nothing
        End Set
    End Property 'ok

    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
            _Cliente = Nothing
            _DescricaoCliente = Nothing
        End Set
    End Property 'ok

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And _CodigoEmpresa > 0 Then _Cliente = New Cliente(_CodigoCliente, _EndCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property 'ok

    Public ReadOnly Property DescricaoCliente()
        Get
            If _DescricaoCliente Is Nothing And Not Cliente Is Nothing Then
                _DescricaoCliente = Cliente.Nome
            End If
            Return _DescricaoCliente
        End Get
    End Property 'ok

    '************** Saldos ************************
    Public Property QtdeProgramadaPedido() As Decimal
        Get
            Return _QtdeProgramadaPedido
        End Get
        Set(ByVal value As Decimal)
            _QtdeProgramadaPedido = value
        End Set
    End Property 'ok

    Public Property Contratado() As Decimal
        Get
            Return _ContratadoFiscal
        End Get
        Set(ByVal value As Decimal)
            _ContratadoFiscal = value
        End Set
    End Property 'ok

    Public Property ContratadoFisico() As Decimal
        Get
            Return _ContratadoFisico
        End Get
        Set(ByVal value As Decimal)
            _ContratadoFisico = value
        End Set
    End Property 'ok

    Public Property EntregueFiscal() As Decimal
        Get
            Return _EntregueFiscal
        End Get
        Set(ByVal value As Decimal)
            _EntregueFiscal = value
        End Set
    End Property 'ok

    Public Property EntregueFisico() As Decimal
        Get
            Return _EntregueFisico
        End Get
        Set(ByVal value As Decimal)
            _EntregueFisico = value
        End Set
    End Property 'ok

    Public Property SaldoFiscal() As Decimal
        Get
            Return _SaldoFiscal
        End Get
        Set(ByVal value As Decimal)
            _SaldoFiscal = value
        End Set
    End Property 'ok

    Public Property SaldoFisico() As Decimal
        Get
            Return _SaldoFisico
        End Get
        Set(ByVal value As Decimal)
            _SaldoFisico = value
        End Set
    End Property 'ok

    '**************** Valores *********************
    Public Property PedidoValor() As Decimal
        Get
            Return _PedidoValor
        End Get
        Set(ByVal value As Decimal)
            _PedidoValor = value
        End Set
    End Property 'ok

    Public Property NotaValor() As Decimal
        Get
            Return _NotaValor
        End Get
        Set(ByVal value As Decimal)
            _NotaValor = value
        End Set
    End Property 'ok

    Public Property SaldoValor() As Decimal
        Get
            Return _SaldoValor
        End Get
        Set(ByVal value As Decimal)
            _SaldoValor = value
        End Set
    End Property 'ok

    '**********************************************
    Public Property Itens As ListSaldoPedidoxItem
        Get
            If _Itens Is Nothing Then _Itens = New ListSaldoPedidoxItem(Me)
            Return _Itens
        End Get
        Set(ByVal value As ListSaldoPedidoxItem)
            _Itens = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub RecarregarItens(Optional ByVal CodigoProduto As Integer = 0, Optional ByVal pDataReferencia As String = "", Optional ByVal suboperacao As SubOperacao = Nothing, Optional ByVal CodEmpresaDev As String = "", Optional ByVal EndEmpresaDev As Integer = 0)
        Me.Itens = New ListSaldoPedidoxItem(Me, CodigoProduto, pDataReferencia, suboperacao, CodEmpresaDev, EndEmpresaDev)
    End Sub
#End Region

End Class