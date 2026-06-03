Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'************************************************************************************************************************************************
'********************************************  Lista de Saldo dos ITENS do Pedidos  *************************************************************
'************************************************************************************************************************************************
Public Class ListSaldoPedidoxItem
    Inherits List(Of SaldoPedidoxItem)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pSaldoPedido As SaldoPedido, Optional ByVal CodigoProduto As Integer = 0, Optional ByVal pDataReferencia As String = "", Optional ByVal suboperacao As SubOperacao = Nothing, Optional ByVal CodEmpresaDev As String = "", Optional ByVal EndEmpresaDev As Integer = 0)
        _ObjPedido = pSaldoPedido.Pedido
        Listar(CodigoProduto, pDataReferencia, suboperacao, CodEmpresaDev, EndEmpresaDev)
    End Sub

    Public Sub New(ByVal CodEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, Optional ByVal CodigoProduto As Integer = 0, Optional ByVal pDataReferencia As String = "", Optional ByVal suboperacao As SubOperacao = Nothing, Optional ByVal CodEmpresaDev As String = "", Optional ByVal EndEmpresaDev As Integer = 0)
        _ObjPedido = New Pedido(CodEmpresa, EndEmpresa, CodigoPedido)
        Listar(CodigoProduto, pDataReferencia, suboperacao, CodEmpresaDev, EndEmpresaDev)
    End Sub

    'Feito pra ser usado no PedidoItem
    Public Sub New(pItemPedido As PedidoXItem)
        _ObjPedido = pItemPedido.Pedido
        Listar(pItemPedido.CodigoProduto)
    End Sub
#End Region

#Region "Fields"
    Private _ObjPedido As Pedido
    Private _ItemPedido As PedidoXItem
    Private _ProdutosPedido As String
#End Region

#Region "Property"
    Public ReadOnly Property ObjPedido As Pedido
        Get
            Return _ObjPedido
        End Get
    End Property

    Public ReadOnly Property ItemPedido As PedidoXItem
        Get
            Return _ItemPedido
        End Get
    End Property

    Public ReadOnly Property ProdutosPedido() As String
        Get
            Return _ProdutosPedido
        End Get
    End Property
#End Region

#Region "Methods"
    Private Sub Listar(Optional ByVal CodigoProduto As Integer = 0, Optional ByVal pDataReferencia As String = "", Optional ByVal suboperacao As SubOperacao = Nothing, Optional ByVal CodFilialDev As String = "", Optional ByVal EndFilialDev As Integer = 0)
        If suboperacao Is Nothing Then suboperacao = ObjPedido.SubOperacao
        Dim sql As String = ""
        sql = "SELECT P.DataPedido, " & vbCrLf & _
              "       P.Empresa_Id, " & vbCrLf & _
              "       P.EndEmpresa_Id, " & vbCrLf & _
              "       P.Cliente, " & vbCrLf & _
              "       P.EndCliente, " & vbCrLf & _
              "       P.Pedido_Id, " & vbCrLf & _
              "       isnull(P.FiscalAberto,1) AS FiscalAberto, " & vbCrLf & _
              "       Produtos.Nome, " & vbCrLf & _
              "       pxi.Produto_Id, " & vbCrLf & _
              "       Produtos.Unidade, " & vbCrLf & _
              "       P.Operacao, " & vbCrLf & _
              "       P.SubOperacao, " & vbCrLf & _
              "       SO.Descricao, " & vbCrLf & _
              "       SO.PrecoFixo, " & vbCrLf & _
              "       P.Moeda, " & vbCrLf & _
              "       CONVERT(NUMERIC(20,10), (case " & vbCrLf & _
              "                                  when SO.Classe in ('AFIXAR','DEPOSITOS') " & vbCrLf & _
              "                                    then CONVERT(NUMERIC(20,10),pxi.UnitarioOficialNormal) " & vbCrLf & _
              "                                  when pxi.QtdeProgramada = 0 " & vbCrLf & _
              "	                                   then CONVERT(NUMERIC(20,10),0) " & vbCrLf & _
              "	                                 else case " & vbCrLf & _
              "	                                        when P.Moeda = 1 " & vbCrLf & _
              "	                                          then CONVERT(NUMERIC(20,10),pxi.PedidoValorOficial) " & vbCrLf & _
              "	                                          else CONVERT(NUMERIC(20,10),pxi.PedidoValorMoeda) " & vbCrLf & _
              "	                                     end / convert(decimal(20,10),pxi.QtdeProgramada) " & vbCrLf & _
              "		                           end)) as unitario, " & vbCrLf & _
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
              "   into #Pedido " & vbCrLf & _
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
              "							then isnull(Quantidadecomercializacao,Quantidade) * -1 --Quantidade * -1 " & vbCrLf & _
              "							else isnull(Quantidadecomercializacao,Quantidade) --Quantidade " & vbCrLf & _
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
              "			   from PedidoxitemXLancamento" & vbCrLf & _
              "		      where Movimento <= '" & IIf(String.IsNullOrEmpty(pDataReferencia), Date.Now.ToString("yyyy-MM-dd"), pDataReferencia) & "'" & vbCrLf & _
              "		      group by Empresa_id, EndEmpresa_id, " & vbCrLf & _
              "                       Pedido_id, " & vbCrLf & _
              "                       Produto_Id " & vbCrLf & _
              "           ) pxi " & vbCrLf & _
              "		ON P.Empresa_Id    = pxi.Empresa_Id " & vbCrLf & _
              "	   AND P.EndEmpresa_Id = pxi.EndEmpresa_Id  " & vbCrLf & _
              "	   AND P.Pedido_Id     = pxi.Pedido_Id  " & vbCrLf

        If CodigoProduto > 0 Then sql &= "    and pxi.Produto_Id     = " & CodigoProduto & vbCrLf

        sql &= "	 Inner Join SubOperacoes SO " & vbCrLf & _
               "		ON SO.Operacao_id     = P.Operacao " & vbCrLf & _
               "	   AND SO.SubOperacoes_id = P.SubOperacao " & vbCrLf & _
               "	 INNER JOIN Produtos  " & vbCrLf & _
               "		ON Produtos.Produto_Id = pxi.Produto_Id " & vbCrLf & _
               "	 INNER JOIN Moedas " & vbCrLf & _
               "		ON Moedas.Moeda_Id = P.Moeda " & vbCrLf & _
               "     Where P.Situacao      = 1 " & vbCrLf & _
               "       and P.Empresa_id    ='" & ObjPedido.CodigoEmpresa & "'" & vbCrLf & _
               "       and P.EndEmpresa_Id = " & ObjPedido.EnderecoEmpresa & vbCrLf & _
               "       and P.Pedido_Id     = " & ObjPedido.Codigo & vbCrLf



        '******************************************************************************************************************************
        '******************************************************************************************************************************
        '******************************************************************************************************************************
        sql &= " Update #Pedido set" & vbCrLf & _
               "      QtdeFixacao = fix.Quantidade" & vbCrLf & _
               "     ,VlrFixacao  = fix.TotalOficial" & vbCrLf & _
               "     ,UntFixacao  = fix.Unitario" & vbCrLf & _
               "   from #Pedido" & vbCrLf & _
               "  INNER JOIN (" & vbCrLf & _
               "               select Pxi.Empresa_id, Pxi.EndEmpresa_id, Pxi.Pedido_Id, Pxi.Produto_id," & vbCrLf & _
               "                      Sum(Pxi.Quantidade) as Quantidade, sum(Pxi.TotalOficial) as TotalOficial," & vbCrLf & _
               "                      sum(Pxi.TotalOficial) / Sum(Pxi.Quantidade) as unitario" & vbCrLf & _
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
               "        isnull(Sb_Nf.Lote,'') as Lote, " & vbCrLf & _
               "        isnull(Sb_Nf.Classificacao,'') as Classificacao, " & vbCrLf & _
               "        isnull(Sb_Nf.Embalagem,0) as Embalagem," & vbCrLf & _
               "        isnull(sb_nf.TipoDeEmbalagem,'') as TipoDeEmbalagem," & vbCrLf & _
               "        isnull(sb_Nf.CapacidadeEmbalagem,0) as CapacidadeEmbalagem," & vbCrLf & _
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
               "					 NxI.Produto_Id," & vbCrLf

        Dim ColocaGroupby As Boolean = False
        If Not suboperacao Is Nothing _
           AndAlso (suboperacao.Devolucao _
                    OrElse suboperacao.Classe = eClassesOperacoes.REAJUSTES _
                    OrElse (suboperacao.QuantidadeFiscal = False And suboperacao.QuantidadeFisico = False) _
                    ) _
           AndAlso suboperacao.Classe <> eClassesOperacoes.GLOBAL Then
            sql &= "					 isnull(NxI.Lote,'') as Lote," & vbCrLf & _
                   "					 isnull(NxI.Classificacao,'') as Classificacao," & vbCrLf & _
                   "					 isnull(NxI.Embalagem,0) as Embalagem," & vbCrLf & _
                   "					 isnull(NxI.TipoDeEmbalagem,'') as TipoDeEmbalagem," & vbCrLf & _
                   "					 isnull(NxI.CapacidadeEmbalagem,0) as CapacidadeEmbalagem," & vbCrLf
            ColocaGroupby = True
        Else
            sql &= "					 '' as Lote," & vbCrLf & _
                   "					 '' as Classificacao," & vbCrLf & _
                   "					 0  as Embalagem," & vbCrLf & _
                   "					 '' as TipoDeEmbalagem," & vbCrLf & _
                   "					 0  as CapacidadeEmbalagem," & vbCrLf
        End If


        sql &= "					 case " & vbCrLf & _
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
               "			   INNER JOIN SubOperacoes as Sb " & vbCrLf & _
               "				  ON NotasFiscais.Operacao    = Sb.Operacao_Id  " & vbCrLf & _
               "				 AND NotasFiscais.SubOperacao = Sb.SubOperacoes_Id " & vbCrLf & _
               "			   INNER JOIN Pedidos P" & vbCrLf & _
               "			 	  ON P.Empresa_Id    = NotasFiscais.Empresa_Id " & vbCrLf & _
               "				 AND P.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id " & vbCrLf & _
               "				 AND P.Pedido_Id     = NotasFiscais.Pedido" & vbCrLf & _
               "		       Where NotasFiscais.Situacao in (1,4,7)" & vbCrLf & _
               "		         and NotasFiscais.TipoDeDocumento = 1 " & vbCrLf & _
               "			     and NotasFiscais.Movimento <= '" & IIf(String.IsNullOrEmpty(pDataReferencia), Date.Now.ToString("yyyy-MM-dd"), pDataReferencia) & "'" & vbCrLf

        If Not suboperacao Is Nothing AndAlso suboperacao.Devolucao Then
            sql &= "			     and notasfiscais.Cliente_Id    ='" & CodFilialDev & "'" & vbCrLf & _
                   "			     and NotasFiscais.EndCliente_Id = " & EndFilialDev & vbCrLf
        End If


        sql &= "		       Group by NotasFiscais.Empresa_Id,  " & vbCrLf & _
               "	  				    NotasFiscais.EndEmpresa_Id, NotasFiscais.Pedido, NxI.Produto_Id," & vbCrLf & _
               "						case " & vbCrLf & _
               "						   when sb.Classe in ('GLOBAL', 'REMESSAS')" & vbCrLf & _
               "							 then 1" & vbCrLf & _
               "						   When sb.Classe in ('AFIXAR','COMPLEMENTACOES')" & vbCrLf & _
               "							 then 2" & vbCrLf & _
               "						   when sb.Classe in ('DEPOSITOS')" & vbCrLf & _
               "							 then 3" & vbCrLf & _
               "						   else 4" & vbCrLf & _
               "						end" & vbCrLf

        If ColocaGroupby Then
            sql &= "						,isnull(NxI.Lote,''), isnull(NxI.Classificacao,''), isnull(NxI.Embalagem,0), isnull(NxI.TipoDeEmbalagem,''), isnull(NxI.CapacidadeEmbalagem,0)" & vbCrLf
        End If

        sql &= "		    ) AS Sb_Nf" & vbCrLf & _
               "	 on Sb_Nf.Empresa_Id    = P.Empresa_Id " & vbCrLf & _
               "    and Sb_Nf.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
               "    and Sb_Nf.Pedido        = P.Pedido_Id " & vbCrLf & _
               "    and Sb_Nf.Produto_Id    = P.Produto_Id;" & vbCrLf & vbCrLf

        '*********************************************************************************************************************************************************************
        '*********************************************************************************************************************************************************************
        '*********************************************************************************************************************************************************************

        If ((Not suboperacao.Devolucao Or suboperacao.Classe = eClassesOperacoes.GLOBAL) And _
            Not suboperacao.Classe = eClassesOperacoes.REAJUSTES) Then
            '((Not suboperacao.Devolucao Or suboperacao.Classe = eClassesOperacoes.GLOBAL) _
            '           And Not suboperacao.Classe = eClassesOperacoes.REAJUSTES _
            '           And Not suboperacao.Classe = eClassesOperacoes.REMESSAS) _
            '           Or (Not suboperacao.QuantidadeFisico _
            '           And Not suboperacao.QuantidadeFiscal) _
            '            Then

            sql &= " select Operacao, SubOperacao," & vbCrLf & _
                   "        Empresa_Id, EndEmpresa_Id," & vbCrLf & _
                   "        Cliente, EndCliente," & vbCrLf & _
                   "        convert(nvarchar(10),DataPedido,103) as DataPedido," & vbCrLf & _
                   "        Produto_Id, Nome, Unidade, " & vbCrLf & _
                   "        ''   as Lote," & vbCrLf & _
                   "        ''   as Classificacao," & vbCrLf & _
                   "        0    as Embalagem," & vbCrLf & _
                   "        ''   as TipoDeEmbalagem," & vbCrLf & _
                   "        0.00 as CapacidadeEmbalagem," & vbCrLf & _
                   "        QtdeProgramada," & vbCrLf & _
                   "        Unitario," & vbCrLf & _
                   "        Pedido_Id," & vbCrLf & _
                   "        FiscalAberto," & vbCrLf & _
                   "        Tipo," & vbCrLf & _
                   "        Contratado," & vbCrLf & _
                   "        ContratadoFisico," & vbCrLf & _
                   "        PedidoValor," & vbCrLf & _
                   "        max(QtdeFixacao)         as QtdeFixacao, " & vbCrLf & _
                   "        max(VlrFixacao)          as VlrFixacao, " & vbCrLf & _
                   "        sum(GlobalNormal)        as GlobalNormal," & vbCrLf & _
                   "        sum(GlobalDevolucao)     as GlobalDevolucao," & vbCrLf & _
                   "        sum(RemessaNormal)       as RemessaNormal," & vbCrLf & _
                   "        sum(RemessaDevolucao)    as RemessaDevolucao," & vbCrLf & _
                   "        sum(DepositadoFiscal)    as DepositadoFiscal," & vbCrLf & _
                   "        sum(DepositadoFisico)    as DepositadoFisico," & vbCrLf & _
                   "        sum(EntregueDeposito)    as EntregueDeposito," & vbCrLf & _
                   "        sum(OperacaoNormal)      as OperacaoNormal," & vbCrLf & _
                   "        sum(OperacaoDevolucao)   as OperacaoDevolucao," & vbCrLf & _
                   "        sum(EntregueFiscal)      as EntregueFiscal," & vbCrLf & _
                   "        convert(numeric(18,4),0) as SaldoFiscal," & vbCrLf & _
                   "        sum(EntregueFisico)      as EntregueFisico," & vbCrLf & _
                   "        convert(numeric(18,4),0) as SaldoFisico," & vbCrLf & _
                   "        sum(NotaValor)           as NotaValor," & vbCrLf & _
                   "        convert(numeric(18,2),0) as SaldoValor," & vbCrLf & _
                   "        convert(numeric(18,6),0) as UnitarioValor " & vbCrLf & _
                   "  Into #PreSaldoPedido " & vbCrLf & _
                   "  from #PreSaldo" & vbCrLf & _
                   " group by Operacao,SubOperacao,Empresa_Id, EndEmpresa_Id, Cliente, EndCliente, convert(nvarchar(10),DataPedido,103) , Produto_Id,Nome,Unidade, " & vbCrLf & _
                   "          QtdeProgramada,Unitario,Pedido_Id, FiscalAberto, Tipo, Contratado, ContratadoFisico, PedidoValor" & vbCrLf & vbCrLf

            'sql &= " update #PreSaldoPedido set" & vbCrLf & _
            '       "  SaldoFiscal   = case when tipo = 2 then EntregueFiscal - QtdeFixacao else Contratado       - EntregueFiscal end," & vbCrLf & _
            '       "  SaldoFisico   = case when tipo = 2 then EntregueFisico - QtdeFixacao else ContratadoFisico - EntregueFisico end," & vbCrLf & _
            '       "  SaldoValor    = case when tipo = 2 then NotaValor      - VlrFixacao  else PedidoValor      - NotaValor end," & vbCrLf & _
            '       "  UnitarioValor = 0" & vbCrLf & vbCrLf

            sql &= "update #PreSaldoPedido set" & vbCrLf & _
                   " SaldoFiscal   = case" & vbCrLf & _
                   "                   when tipo = 2 then EntregueFiscal - QtdeFixacao " & vbCrLf & _
                   "                   when tipo = 3 then case" & vbCrLf & _
                   "                                        when Contratado > 0" & vbCrLf & _
                   "                                          then Contratado - EntregueFiscal" & vbCrLf & _
                   "                                          else EntregueFiscal" & vbCrLf & _
                   "                                      end" & vbCrLf & _
                   "                   else Contratado       - EntregueFiscal end," & vbCrLf & _
                   " SaldoFisico   = case" & vbCrLf & _
                   "                   when tipo = 2 then EntregueFisico - QtdeFixacao" & vbCrLf & _
                   "                   when tipo = 3 then Case" & vbCrLf & _
                   "                                        When ContratadoFisico > 0" & vbCrLf & _
                   "                                          then ContratadoFisico - EntregueFisico" & vbCrLf & _
                   "                                          else EntregueFisico" & vbCrLf & _
                   "                                      end" & vbCrLf & _
                   "                   else ContratadoFisico - EntregueFisico end," & vbCrLf & _
                   " SaldoValor    = case" & vbCrLf & _
                   "                     when tipo = 2" & vbCrLf & _
                   "                       then NotaValor   - VlrFixacao" & vbCrLf & _
                   "                       else PedidoValor - NotaValor" & vbCrLf & _
                   "                 end," & vbCrLf & _
                   " UnitarioValor = 0" & vbCrLf

            sql &= " Select * from #PreSaldoPedido order by Produto_Id  " & vbCrLf
        Else
            sql &= " select Operacao,SubOperacao,Empresa_Id, EndEmpresa_Id, Cliente, EndCliente, convert(nvarchar(10),DataPedido,103) as DataPedido, Produto_Id,Nome,Unidade," & vbCrLf & _
                  "        Lote, Classificacao, Embalagem, TipoDeEmbalagem, CapacidadeEmbalagem," & vbCrLf & _
                  "        QtdeProgramada,Unitario,Pedido_Id, FiscalAberto, Tipo, Contratado, ContratadoFisico, PedidoValor," & vbCrLf & _
                  "        max(QtdeFixacao)         as QtdeFixacao, " & vbCrLf & _
                  "        max(VlrFixacao)          as VlrFixacao, " & vbCrLf & _
                  "        sum(GlobalNormal)        as GlobalNormal," & vbCrLf & _
                  "        Sum(GlobalDevolucao)     as GlobalDevolucao," & vbCrLf & _
                  "        sum(RemessaNormal)       as RemessaNormal," & vbCrLf & _
                  "        sum(RemessaDevolucao)    as RemessaDevolucao," & vbCrLf & _
                  "        sum(DepositadoFiscal)    as DepositadoFiscal," & vbCrLf & _
                  "        sum(DepositadoFisico)    as DepositadoFisico," & vbCrLf & _
                  "        sum(EntregueDeposito)    as EntregueDeposito," & vbCrLf & _
                  "        sum(OperacaoNormal)      as OperacaoNormal," & vbCrLf & _
                  "        sum(OperacaoDevolucao)   as OperacaoDevolucao," & vbCrLf & _
                  "        sum(EntregueFiscal)      as EntregueFiscal," & vbCrLf & _
                  "        convert(numeric(18,4),0) as SaldoFiscal," & vbCrLf & _
                  "        sum(EntregueFisico)      as EntregueFisico, " & vbCrLf & _
                  "        convert(numeric(18,4),0) as SaldoFisico," & vbCrLf & _
                  "        sum(NotaValor)           as NotaValor, " & vbCrLf & _
                  "        convert(numeric(18,2),0) as SaldoValor,  " & vbCrLf & _
                  "        convert(numeric(18,6),0) as UnitarioValor  " & vbCrLf & _
                  "  Into #PreSaldoPedido " & vbCrLf & _
                  "  from #PreSaldo" & vbCrLf & _
                  " group by Operacao,SubOperacao,Empresa_Id, EndEmpresa_Id, Cliente, EndCliente, convert(nvarchar(10),DataPedido,103), Produto_Id, Nome, Unidade, " & vbCrLf & _
                  "          Lote, Classificacao, Embalagem, TipoDeEmbalagem, CapacidadeEmbalagem, QtdeProgramada, Unitario, Pedido_Id, FiscalAberto, Tipo, Contratado, ContratadoFisico, PedidoValor" & vbCrLf & vbCrLf


            sql &= " update #PreSaldoPedido set" & vbCrLf & _
                   " Contratado    = sb.contratado," & vbCrLf & _
                   "SaldoFiscal   = case" & vbCrLf & _
                   "                  when tipo = 2 then sb.EntregueFiscal - QtdeFixacao" & vbCrLf & _
                   "                  when tipo = 3 then case" & vbCrLf & _
                   "                                       when sb.Contratado > 0" & vbCrLf & _
                   "                                         then sb.Contratado - sb.EntregueFiscal" & vbCrLf & _
                   "                                         else sb.EntregueFiscal" & vbCrLf & _
                   "                                     end" & vbCrLf & _
                   "                  else sb.Contratado       - sb.EntregueFiscal end," & vbCrLf & _
                   "SaldoFisico   = case" & vbCrLf & _
                   "                  when tipo = 2 then sb.EntregueFisico - QtdeFixacao" & vbCrLf & _
                   "                  when tipo = 3 then Case" & vbCrLf & _
                   "                                       When sb.ContratadoFisico > 0" & vbCrLf & _
                   "                                         then sb.ContratadoFisico - sb.EntregueFisico" & vbCrLf & _
                   "                                         else sb.EntregueFisico" & vbCrLf & _
                   "                                     end" & vbCrLf & _
                   "                  else sb.ContratadoFisico - sb.EntregueFisico end," & vbCrLf & _
                   "  SaldoValor  = case when tipo = 2 then sb.NotaValor      - VlrFixacao  else " & IIf(Not suboperacao.Devolucao Or (Not suboperacao.QuantidadeFiscal And Not suboperacao.QuantidadeFisico), "PedidoValor      - sb.NotaValor", "sb.NotaValor") & "      end" & vbCrLf & _
                   "   from (select Pedido_Id, Produto_Id, Lote," & vbCrLf & _
                   "                max(Contratado) as Contratado," & vbCrLf & _
                   "                max(ContratadoFisico) as ContratadoFisico," & vbCrLf & _
                   "                sum(EntregueFiscal)as EntregueFiscal," & vbCrLf & _
                   "                sum(EntregueFisico)as EntregueFisico," & vbCrLf & _
                   "                sum(NotaValor) as NotaValor" & vbCrLf & _
                   "           from #PreSaldoPedido" & vbCrLf & _
                   "          group by Pedido_Id, Produto_Id, Lote" & vbCrLf & _
                   "        ) sb" & vbCrLf & _
                   "   inner join #PreSaldoPedido" & vbCrLf & _
                   "      on #PreSaldoPedido.Pedido_Id  = sb.Pedido_Id" & vbCrLf & _
                   "     and #PreSaldoPedido.Produto_Id = sb.Produto_Id" & vbCrLf & _
                   "     and #PreSaldoPedido.Lote = sb.Lote" & vbCrLf

            sql &= " Select * from #PreSaldoPedido " & vbCrLf
            If suboperacao.Classe <> eClassesOperacoes.GLOBAL And suboperacao.Classe <> eClassesOperacoes.REMESSAS Then sql &= " where GlobalNormal = 0 " & vbCrLf
            sql &= "  order by Produto_Id" & vbCrLf
        End If


        sql &= " drop table #Pedido;" & vbCrLf & _
               " drop table #PreSaldo;" & vbCrLf & _
               " drop table #PreSaldoPedido;"

        Dim ds As DataSet
        Dim banco As New AcessaBanco
        ds = banco.ConsultaDataSet(sql, "PreSaldoPedido")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim SPxN As New SaldoPedidoxItem

                SPxN.QtdeProgramadaPedido = row("QtdeProgramada")
                SPxN.Contratado = row("Contratado")
                SPxN.ContratadoFisico = row("ContratadoFisico")


                SPxN.EntregueFiscal = row("EntregueFiscal")
                SPxN.SaldoFiscal = row("SaldoFiscal")

                SPxN.EntregueFiscalComercializacao = row("EntregueFiscal")
                SPxN.SaldoFiscalComercializacao = row("SaldoFiscal")


                SPxN.EntregueFisico = row("EntregueFisico")
                SPxN.SaldoFisico = row("SaldoFisico")

                SPxN.PedidoValor = row("PedidoValor")
                SPxN.NotaValor = row("NotaValor")
                SPxN.SaldoValor = row("SaldoValor")

                SPxN.UnitarioValor = row("UnitarioValor")

                SPxN.Produto = row("Produto_Id")
                SPxN.NomeProduto = row("Nome")
                SPxN.Unidade = row("Unidade")
                SPxN.Lote = row("Lote")
                SPxN.Classificacao = row("Classificacao")
                SPxN.CodigoEmbalagem = row("Embalagem")
                SPxN.TipoDeEmbalagem = row("TipoDeEmbalagem")
                SPxN.CapacidadeEmbalagem = row("CapacidadeEmbalagem")

                SPxN.Quantidade = row("QtdeProgramada")
                SPxN.Unitario = row("Unitario")

                SPxN.GlobalNormal = row("GlobalNormal")
                SPxN.GlobalDevolucao = row("GlobalDevolucao")
                SPxN.RemessaNormal = row("RemessaNormal")
                SPxN.RemessaDevolucao = row("RemessaDevolucao")
                SPxN.DepositadoFiscal = row("DepositadoFiscal")
                SPxN.DepositadoFisico = row("DepositadoFisico")
                SPxN.EntregueDeposito = row("EntregueDeposito")
                SPxN.OperacaoNormal = row("OperacaoNormal")
                SPxN.OperacaoDevolucao = row("OperacaoDevolucao")

                Me.Add(SPxN)
            Next
        End If

    End Sub
#End Region

End Class

'************************************************************************************************************************************************
'***************************************************  Saldo Do Item Do Pedido  ******************************************************************
'************************************************************************************************************************************************
Public Class SaldoPedidoxItem

#Region "Fields"
    Private _Selecionado As Boolean = False
    Private _CodigoPedido As String
    Private _Pedido As Pedido
    Private _Produto As String
    Private _NomeProduto As String

    Private _Lote As String
    Private _Classificacao As String

    Private _CodigoEmbalagem As Integer
    Private _Embalagem As Embalagem

    Private _TipoDeEmbalagem As String
    Private _CapacidadeEmbalagem As Decimal

    Private _Quantidade As Decimal
    Private _Unitario As String
    Private _Unidade As String

    Private _Contratado As Decimal
    Private _EntregueFiscal As Decimal
    Private _SaldoFiscal As Decimal

    Private _EntregueFiscalComercializacao As Decimal
    Private _SaldoFiscalComercializacao As Decimal

    Private _ContratadoFisico As Decimal
    Private _EntregueFisico As Decimal
    Private _SaldoFisico As Decimal

    Private _GlobalNormal As Decimal
    Private _GlobalDevolucao As Decimal

    Private _RemessaNormal As Decimal
    Private _RemessaDevolucao As Decimal

    Private _DepositadoFiscal As Decimal
    Private _DepositadoFisico As Decimal
    Private _EntregueDeposito As Decimal

    Private _OperacaoNormal As Decimal
    Private _OperacaoDevolucao As Decimal

    Private _QtdeProgramadaPedido As Decimal
    Private _UnitarioValor As Decimal

    'Valores 
    Private _PedidoValor As Decimal
    Private _NotaValor As Decimal
    Private _SaldoValor As Decimal
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

    Public Property Pedido() As Pedido
        Get
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
        End Set
    End Property

    Public Property Produto() As String
        Get
            Return _Produto
        End Get
        Set(ByVal value As String)
            _Produto = value
        End Set
    End Property

    Public Property NomeProduto() As String
        Get
            Return _NomeProduto
        End Get
        Set(ByVal value As String)
            _NomeProduto = value
        End Set
    End Property

    Public Property Unidade() As String
        Get
            Return _Unidade
        End Get
        Set(ByVal value As String)
            _Unidade = value
        End Set
    End Property

    Public Property Lote() As String
        Get
            Return _Lote
        End Get
        Set(ByVal value As String)
            _Lote = value
        End Set
    End Property

    Public Property Classificacao() As String
        Get
            Return _Classificacao
        End Get
        Set(ByVal value As String)
            _Classificacao = value
        End Set
    End Property

    Public Property CodigoEmbalagem() As Integer
        Get
            Return _CodigoEmbalagem
        End Get
        Set(ByVal value As Integer)
            _CodigoEmbalagem = value
        End Set
    End Property

    Public ReadOnly Property Embalagem() As Embalagem
        Get
            If _Embalagem Is Nothing And _CodigoEmbalagem > 0 Then _Embalagem = New Embalagem(_CodigoEmbalagem)
            Return _Embalagem
        End Get
    End Property

    Public ReadOnly Property EmbalagemIndea() As String
        Get
            If Embalagem Is Nothing Then
                Return ""
            Else
                Return Embalagem.EmbalagemIndea
            End If
        End Get
    End Property

    Public ReadOnly Property DescricaoEmbalagem() As String
        Get
            If Embalagem Is Nothing Then
                Return ""
            Else
                Return Embalagem.EmbalagemIndea & " - " & _TipoDeEmbalagem & " / " & _CapacidadeEmbalagem
            End If
        End Get
    End Property

    Public Property TipoDeEmbalagem() As String
        Get
            Return _TipoDeEmbalagem
        End Get
        Set(ByVal value As String)
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

    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
        End Set
    End Property

    'Unitario Medio do Pedido Calculado
    Public Property Unitario() As Decimal
        Get
            Return _Unitario
        End Get
        Set(ByVal value As Decimal)
            'If QtdeProgramadaPedido = 0 Then
            '    _Unidade = value
            'Else
            '    Dim Unt As Decimal

            '    Unt = (PedidoValor / QtdeProgramadaPedido) * BaseDeCalculo
            '    If Math.Round(QtdeProgramadaPedido * Math.Round(Unt, 0, MidpointRounding.AwayFromZero) / BaseDeCalculo, 2, MidpointRounding.AwayFromZero) = PedidoValor Then
            '        _Unitario = Math.Round(Unt, 0, MidpointRounding.AwayFromZero)
            '    Else
            '        _Unitario = (PedidoValor / QtdeProgramadaPedido) * BaseDeCalculo
            '    End If
            'End If

            _Unitario = value
        End Set
    End Property

    'Public Property BaseDeCalculo() As Integer
    '    Get
    '        Return _BaseDeCalculo
    '    End Get
    '    Set(ByVal value As Integer)
    '        _BaseDeCalculo = value
    '    End Set
    'End Property

    Public ReadOnly Property Total() As Decimal
        Get
            Return Math.Round(_Quantidade * _Unitario, 2)
        End Get
    End Property

    Public Property QtdeProgramadaPedido() As Decimal
        Get
            Return _QtdeProgramadaPedido
        End Get
        Set(ByVal value As Decimal)
            _QtdeProgramadaPedido = value
        End Set
    End Property

    Public Property Contratado() As Decimal
        Get
            Return _Contratado
        End Get
        Set(ByVal value As Decimal)
            _Contratado = value
        End Set
    End Property

    'Fiscal
    Public Property EntregueFiscal() As Decimal
        Get
            Return _EntregueFiscal
        End Get
        Set(ByVal value As Decimal)
            _EntregueFiscal = value
        End Set
    End Property

    Public Property SaldoFiscal() As Decimal
        Get
            Return _SaldoFiscal
        End Get
        Set(ByVal value As Decimal)
            _SaldoFiscal = value
        End Set
    End Property

    'Fiscal Unidade de Comercializacao
    Public Property EntregueFiscalComercializacao As Decimal
        Get
            Return _EntregueFiscalComercializacao
        End Get
        Set(value As Decimal)
            _EntregueFiscalComercializacao = value
        End Set
    End Property

    Public Property SaldoFiscalComercializacao As Decimal
        Get
            Return _SaldoFiscalComercializacao
        End Get
        Set(value As Decimal)
            _SaldoFiscalComercializacao = value
        End Set
    End Property

    Public Property ContratadoFisico() As Decimal
        Get
            Return _ContratadoFisico
        End Get
        Set(ByVal value As Decimal)
            _ContratadoFisico = value
        End Set
    End Property

    Public Property EntregueFisico() As Decimal
        Get
            Return _EntregueFisico
        End Get
        Set(ByVal value As Decimal)
            _EntregueFisico = value
        End Set
    End Property

    Public Property SaldoFisico() As Decimal
        Get
            Return _SaldoFisico
        End Get
        Set(ByVal value As Decimal)
            _SaldoFisico = value
        End Set
    End Property

    Public ReadOnly Property [Global]() As Decimal
        Get
            Return _GlobalNormal - _GlobalDevolucao
        End Get
    End Property

    Public Property GlobalNormal() As Decimal
        Get
            Return _GlobalNormal
        End Get
        Set(ByVal value As Decimal)
            _GlobalNormal = value
        End Set
    End Property

    Public Property GlobalDevolucao() As Decimal
        Get
            Return _GlobalDevolucao
        End Get
        Set(ByVal value As Decimal)
            _GlobalDevolucao = value
        End Set
    End Property

    Public Property RemessaNormal() As Decimal
        Get
            Return _RemessaNormal
        End Get
        Set(ByVal value As Decimal)
            _RemessaNormal = value
        End Set
    End Property

    Public Property RemessaDevolucao() As Decimal
        Get
            Return _RemessaDevolucao
        End Get
        Set(ByVal value As Decimal)
            _RemessaDevolucao = value
        End Set
    End Property

    Public Property DepositadoFiscal() As Decimal
        Get
            Return _DepositadoFiscal
        End Get
        Set(ByVal value As Decimal)
            _DepositadoFiscal = value
        End Set
    End Property

    Public Property DepositadoFisico() As Decimal
        Get
            Return _DepositadoFisico
        End Get
        Set(ByVal value As Decimal)
            _DepositadoFisico = value
        End Set
    End Property

    Public Property EntregueDeposito() As Decimal
        Get
            Return _EntregueDeposito
        End Get
        Set(ByVal value As Decimal)
            _EntregueDeposito = value
        End Set
    End Property

    Public Property OperacaoNormal() As Decimal
        Get
            Return _OperacaoNormal
        End Get
        Set(ByVal value As Decimal)
            _OperacaoNormal = value
        End Set
    End Property

    Public Property OperacaoDevolucao() As Decimal
        Get
            Return _OperacaoDevolucao
        End Get
        Set(ByVal value As Decimal)
            _OperacaoDevolucao = value
        End Set
    End Property

    'Unitario Medio Sujerido calculado com parametro no saldo de quantidade e valor
    Public Property UnitarioValor() As Decimal
        Get
            Return _UnitarioValor
        End Get
        Set(ByVal value As Decimal)
            'If SaldoFiscal = 0 Then
            '    _UnitarioValor = Unitario
            'Else
            '    Dim Unt As Decimal
            '    Unt = (SaldoValor / SaldoFiscal) * BaseDeCalculo
            '    If Math.Round(SaldoFiscal * Math.Round(Unt, 0, MidpointRounding.AwayFromZero) / BaseDeCalculo, 2, MidpointRounding.AwayFromZero) = SaldoValor Then
            '        _UnitarioValor = Math.Round(Unt, 0, MidpointRounding.AwayFromZero)
            '    Else
            '        _UnitarioValor = (SaldoValor / SaldoFiscal) * BaseDeCalculo
            '    End If
            'End If
            _UnitarioValor = value
        End Set
    End Property

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
#End Region

End Class
