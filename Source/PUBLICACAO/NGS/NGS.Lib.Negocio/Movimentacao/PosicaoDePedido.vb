Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListPosicaoDePedido
    Inherits List(Of PosicaoDePedido)

    Private Function getStrEmpresasConsolidadas(ByVal ListEmpresas As List(Of String)) As List(Of String)
        Dim list As New List(Of String)
        If ListEmpresas.Count > 0 Then
            For Each row As String In ListEmpresas
                list.Add(Left(row, 8))
            Next
        End If
        Return list
    End Function

    Public Function SqlPosicao(ByVal pData As Date, ByVal ListEmpresa As List(Of String), ByVal EmpresaConsolidado As Boolean, ByVal Cliente As Cliente, ByVal ClienteConsolidado As Boolean, ByVal TabelaTemporaria As String,
                               ByVal PProduto As String, ByVal PsafraInicial As String, ByVal PsafraFinal As String, Optional ByVal pSQLProduto As String = "", Optional ByVal pDataEntregaInicial As String = "",
                               Optional ByVal pDataEntregaFinal As String = "", Optional ByVal pDataInicialAberturaPedido As String = "", Optional ByVal pDataFinalAberturaPedido As String = "",
                               Optional ByVal pSituacaoDoPedidoAbertoFechado As String = "", Optional ByVal pFrete As Integer = 0, Optional ByVal pClasseOperacao As String = "", Optional ByVal pClasseSubOperacao As String = "",
                               Optional ByVal pCodigoOperacao As Integer = 0, Optional ByVal pCodigoSuboperacao As Integer = 0, Optional ByVal pSaldo As Integer = 0, Optional ByVal pCodigoMoeda As Integer = 0,
                               Optional ByVal DeletarContratoZerado As Boolean = True, Optional pSituacaoTroca As Integer = 0, Optional pOperacaoAntecipada As Integer = 0, Optional pRecompra As Integer = 0,
                               Optional Cessionario As Boolean = False, Optional ByVal Representante As Cliente = Nothing, Optional ByVal pSQLOperacoes As String = "") As String

        'pSituacaoDoPedidoAbertoFechado = A Aberto, F Fechado
        'pFrete = 0 Todos 1 CIF 2 FOB
        'pSaldo 0 = todos 1 = Com Saldo 2= liquidados
        'pSitucaoTroca         0=todos 1= sim  2= Nao
        'pOperacaoAntecipada   0=todos 1= sim  2= Nao
        'Recompra              0=todos 1= sim  2= nao

        Dim strEmpresas As String = String.Join("','", ListEmpresa)
        Dim strEmpresasConsolidadas As String = String.Join("','", getStrEmpresasConsolidadas(ListEmpresa))

        Dim Data As String = "'" & pData.ToString("yyyy-MM-dd") & "'"
        Dim Sql As String = String.Empty

        If Not String.IsNullOrEmpty(PsafraInicial) Then
            Sql &= " Select safra_Id " & vbCrLf &
                   "   into #Safra" & vbCrLf &
                   "   from safras " & vbCrLf &
                   "  where Vencimento >= (Select Vencimento " & vbCrLf &
                   "                         from Safras " & vbCrLf &
                   "                        where safra_id = '" & PsafraInicial & "')" & vbCrLf &
                   "    and Vencimento <= (Select Vencimento " & vbCrLf &
                   "                         from Safras " & vbCrLf &
                   "                        where safra_id = '" & IIf(String.IsNullOrEmpty(PsafraFinal), PsafraInicial, PsafraFinal) & "')" & vbCrLf
        End If

        '*********************************************************************************************************************
        '****************************************** Cessao de Credito ********************************************************
        '*********************************************************************************************************************
        Sql &= "select Empresa_Id, EndEmpresa_Id, Pedido_Id, DataPedido, Procuracao_ID, DataProcuracao, grupo, Produto_Id, ClienteCedente, EndCedente, ClienteCessionario, EndCessionario, QtdeProcuracao, QtdeEntregue, QtdeSaldo" & vbCrLf &
               "  Into #Procuracoes" & vbCrLf &
               "  from VW_CessaoDeCredito" & vbCrLf &
               " Where Situacao = 1" & vbCrLf

        If ListEmpresa IsNot Nothing AndAlso ListEmpresa.Count > 0 Then
            If EmpresaConsolidado Then
                Sql &= "   AND left(Empresa_Id,8) in ('" & strEmpresasConsolidadas & "')" & vbCrLf
            Else
                Sql &= "   AND Empresa_Id + Cast(EndEmpresa_Id as varchar) in ('" & strEmpresas & "')" & vbCrLf
            End If
        End If

        Sql &= "   AND DataProcuracao <= " & Data & vbCrLf &
               "   AND DataPedido     <= " & Data & vbCrLf

        'Selecao de Produtos User Control Selecao de Produto
        If Not String.IsNullOrEmpty(pSQLProduto) Then
            Sql &= "   AND " & pSQLProduto.Replace("Prd.", "") & vbCrLf
        Else
            Sql &= "   AND Produto_Id in (" & PProduto & ")" & vbCrLf
        End If

        '*********************************************************************************************************************
        '************************************  MONTA TABELA TEMPORARIA *******************************************************
        '*********************************************************************************************************************
        Sql &= "Select convert(varchar(3), '')   as Tipo," & vbCrLf &
               "       convert(varchar(18),  0)  as Empresa_Id," & vbCrLf &
               "       0                         as EndEmpresa_Id, " & vbCrLf &
               "       0                         as Pedido_Id," & vbCrLf &
               "       convert(varchar(20),  '') as Classe," & vbCrLf &
               "       0                         as Procuracao," & vbCrLf &
               "       convert(Decimal(18,4), 0) as Contratado," & vbCrLf &
               "       convert(Decimal(18,4), 0) as Laudo," & vbCrLf &
               "       convert(Decimal(18,4), 0) as Entregue," & vbCrLf &
               "       convert(Decimal(18,4), 0) as AEntregar," & vbCrLf &
               "       convert(Decimal(18,4), 0) as QuantidadeFixado," & vbCrLf &
               "       convert(Decimal(18,2), 0) as ValorFixadoOficial," & vbCrLf &
               "       convert(Decimal(18,2), 0) as ValorFixadoMoeda," & vbCrLf &
               "       convert(Decimal(18,4), 0) as AFixar," & vbCrLf &
               "       convert(Decimal(18,4), 0) as Pago," & vbCrLf &
               "       convert(Decimal(18,4), 0) as PagoNaoRecebido," & vbCrLf &
               "       convert(Decimal(18,4), 0) as RecebidoNaoPago," & vbCrLf &
               "       convert(nvarchar(30),  0) as Produto," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Adiantamento," & vbCrLf &
               "       convert(Decimal(18,2), 0) as AdiantamentoOficial," & vbCrLf &
               "       convert(Decimal(18,2), 0) as BaixaAdiantamentoOficial," & vbCrLf &
               "       convert(Decimal(18,2), 0) as AdiantamentoMoeda," & vbCrLf &
               "       convert(Decimal(18,2), 0) as BaixaAdiantamentoMoeda," & vbCrLf &
               "       convert(Decimal(18,9), 0) as UnitarioPedido," & vbCrLf &
               "       convert(Decimal(18,9), 0) as UnitarioOficial," & vbCrLf &
               "       convert(Decimal(18,9), 0) as UnitarioMoeda," & vbCrLf &
               "       convert(Decimal(18,2), 0) as TotalOficial," & vbCrLf &
               "       convert(Decimal(18,2), 0) as TotalMoeda," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Programado_PagarOficial," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Baixado_PagarOficial," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Programado_PagarMoeda," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Baixado_PagarMoeda," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Programado_ReceberOficial," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Baixado_ReceberOficial," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Programado_ReceberMoeda," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Baixado_ReceberMoeda," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Programado_SaldoOficial," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Baixado_SaldoOficial," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Programado_SaldoMoeda," & vbCrLf &
               "       convert(Decimal(18,2), 0) as Baixado_SaldoMoeda," & vbCrLf &
               "       convert(nvarchar(20),  0) as OrigemDestino," & vbCrLf &
               "       convert(nvarchar(50),  0) as Contrato," & vbCrLf &
               "       convert(nvarchar(1),   0) as EntradaSaida," & vbCrLf &
               "       convert(decimal(18,2), 0) as cedente," & vbCrLf &
               "       convert(decimal(18,2), 0) as cessionario," & vbCrLf &
               "       0                         as MoedaPedido" & vbCrLf &
               "  into #TPosicaoDeContratos;" & vbCrLf

        '****************************************************************************************************************************************************************
        Sql &= "Delete #TPosicaoDeContratos;" & vbCrLf
        '****************************************************************************************************************************************************************

        '****************************************************************************************************************************************************************
        '*************************************************  Inseri Dados preliminares empresa pedido produto ************************************************************
        '****************************************************************************************************************************************************************
        Sql &= "SELECT 'CED' as Tipo, P.Empresa_Id, P.EndEmpresa_Id, P.Pedido_Id," & vbCrLf &
               "       P.OrigemDestino, P.Contrato, SO.EntradaSaida, P.Moeda," & vbCrLf &
               "	   PxI.Produto_Id, so.classe," & vbCrLf &
               "	   0 as Procuracao_Id," & vbCrLf &
               "	   isnull(QtdeProcuracao,0) as QtdeProcuracao," & vbCrLf &
               "	   isnull(QtdeEntregue,0)   as QtdeEntregue," & vbCrLf &
               "	   isnull(QtdeSaldo,0)      as QtdeSaldo" & vbCrLf &
               "  into #Pedidos" & vbCrLf &
               "  FROM Pedidos P" & vbCrLf &
               " INNER JOIN PedidoXItem PxI" & vbCrLf &
               "    ON P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf &
               "   AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf &
               "   AND P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf &
               " INNER JOIN Produtos Prd" & vbCrLf &
               "    ON PxI.Produto_Id = Prd.Produto_Id" & vbCrLf &
               " Inner Join Operacoes OP" & vbCrLf &
               "    on OP.Operacao_ID = P.Operacao" & vbCrLf &
               " INNER JOIN SubOperacoes SO" & vbCrLf &
               "    ON P.Operacao    = SO.Operacao_Id" & vbCrLf &
               "   AND P.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
               "  Left Join(Select Empresa_id, EndEmpresa_id, Pedido_id," & vbCrLf &
               "                  sum(QtdeProcuracao) as QtdeProcuracao," & vbCrLf &
               "				  sum(QtdeEntregue)   as QtdeEntregue," & vbCrLf &
               "				  sum(QtdeSaldo)      as QtdeSaldo" & vbCrLf &
               "			 from #Procuracoes" & vbCrLf &
               "			Group by Empresa_id, EndEmpresa_id, Pedido_id" & vbCrLf &
               "          ) sbProc" & vbCrLf &
               "    on P.Empresa_Id    = sbProc.Empresa_id" & vbCrLf &
               "   and P.EndEmpresa_Id = sbProc.EndEmpresa_Id" & vbCrLf &
               "   and P.Pedido_Id     = sbProc.Pedido_id" & vbCrLf

        If Not String.IsNullOrEmpty(PsafraInicial) Then
            Sql &= " Inner join #Safra" & vbCrLf &
                   "    on #Safra.Safra_Id =  P.Safra" & vbCrLf
        End If

        If Not Representante Is Nothing Then
            Sql &= " left join Comissoes pXComissoes" & vbCrLf &
                    "       on pXComissoes.Empresa_Id    = P.Empresa_Id" & vbCrLf &
                    "      and pXComissoes.Endempresa_id = P.Endempresa_id" & vbCrLf &
                    "      and pXComissoes.Pedido_Id     = P.Pedido_Id" & vbCrLf
        End If

        Sql &= " WHERE P.Situacao     = 1 " & vbCrLf &
               "   And P.DataPedido  <=" & Data & vbCrLf &
               " --And pxi.Movimento <=" & Data & vbCrLf &
               "   AND Prd.Agrupar    ='N'" & vbCrLf &
               "   AND NOT EXISTS(SELECT 1 FROM NotasFiscais WHERE isnull(NFG,0) = 1 and Empresa_Id = P.Empresa_Id AND Endempresa_id = P.EndEmpresa_Id AND Pedido = P.Pedido_Id)" & vbCrLf

        If Not String.IsNullOrEmpty(pDataEntregaInicial) And Not String.IsNullOrEmpty(pDataEntregaFinal) Then
            If IsDate(pDataEntregaInicial) And IsDate(pDataEntregaFinal) Then
                Sql &= "   And P.DataEntrega between '" & pDataEntregaInicial & "' and '" & pDataEntregaFinal & "'" & vbCrLf
            End If
        End If

        If Not String.IsNullOrEmpty(pDataInicialAberturaPedido) And Not String.IsNullOrEmpty(pDataFinalAberturaPedido) Then
            If IsDate(pDataInicialAberturaPedido) And IsDate(pDataFinalAberturaPedido) Then
                Sql &= "   And P.DataPedido between '" & pDataInicialAberturaPedido & "' and '" & pDataFinalAberturaPedido & "'" & vbCrLf
            End If
        End If

        If pSituacaoDoPedidoAbertoFechado = "A" Then
            Sql &= "   AND (P.FiscalAberto = 'true' OR P.FinanceiroAberto = 'true') " & vbCrLf
        ElseIf pSituacaoDoPedidoAbertoFechado = "F" Then
            Sql &= "   AND P.FiscalAberto = 'false' AND P.FinanceiroAberto = 'false' " & vbCrLf
        End If

        Select Case pFrete
            Case 1 : Sql &= "   AND isnull(P.FreteCIFFOB,'FOB') = 'CIF'" & vbCrLf
            Case 2 : Sql &= "   AND isnull(P.FreteCIFFOB,'FOB') = 'FOB'" & vbCrLf
        End Select


        Select Case pSituacaoTroca
            Case 1 : Sql &= "   AND isnull(p.troca,0) = 1" & vbCrLf
            Case 2 : Sql &= "   AND isnull(p.troca,0) = 0" & vbCrLf
        End Select

        Select Case pOperacaoAntecipada
            Case 1 : Sql &= "   AND isnull(p.Antecipada,0) = 1" & vbCrLf
            Case 2 : Sql &= "   AND isnull(p.Antecipada,0) = 0" & vbCrLf
        End Select

        Select Case pRecompra
            Case 1 : Sql &= "   AND isnull(p.Recompra,0) = 1" & vbCrLf
            Case 2 : Sql &= "   AND isnull(p.Recompra,0) = 0" & vbCrLf
        End Select

        If Not String.IsNullOrEmpty(pClasseOperacao) Then Sql &= "   AND OP.Classe = '" & pClasseOperacao & "'" & vbCrLf
        If Not String.IsNullOrEmpty(pClasseSubOperacao) Then Sql &= "   AND SO.Classe = '" & pClasseSubOperacao & "'" & vbCrLf

        If ListEmpresa IsNot Nothing AndAlso ListEmpresa.Count > 0 Then
            If EmpresaConsolidado Then
                Sql &= " AND left(P.Empresa_Id,8) in ('" & strEmpresasConsolidadas & "')" & vbCrLf
            Else
                Sql &= "   AND P.Empresa_Id + Cast(P.EndEmpresa_Id as varchar) in ('" & strEmpresas & "')" & vbCrLf
            End If
        End If

        If Not Cliente Is Nothing Then
            If ClienteConsolidado Then
                Sql &= "   AND (left(P.Cliente,8) ='" & Left(Cliente.Codigo, 8) & "'" & vbCrLf &
                       "        OR" & vbCrLf &
                       "        exists(select 1 from procuracoes where pedidocedente = p.pedido_Id and left(cessionario,8) = '" & Left(Cliente.Codigo, 8) & "'))"
            Else
                Sql &= "   AND ((P.Cliente    ='" & Cliente.Codigo & "' AND P.EndCliente = " & Cliente.CodigoEndereco & ")" & vbCrLf &
                       "        OR" & vbCrLf &
                       "        exists(select 1 from procuracoes where pedidocedente = p.pedido_Id and cessionario = '" & Cliente.Codigo & "' and endcessionario = " & Cliente.CodigoEndereco & "))"

            End If
        End If

        If Not Representante Is Nothing Then
            Sql &= "   AND (pXComissoes.Representante_Id = '" & Representante.Codigo & "' AND pXComissoes.EndRepresentante_Id = " & Representante.CodigoEndereco & ")" & vbCrLf
        End If

        'Selecao de SubOperações User Control Selecao de Operações
        If Not String.IsNullOrEmpty(pSQLOperacoes) Then

            pSQLOperacoes = pSQLOperacoes.Replace("Operacao_Id", "P.Operacao").Replace("SubOperacoes_Id", "P.SubOperacao")

            Sql &= "   AND " & pSQLOperacoes & vbCrLf
        Else
            If pCodigoOperacao > 0 Then
                Sql &= "   AND P.Operacao = " & pCodigoOperacao & " " & vbCrLf
            End If
            If pCodigoSuboperacao > 0 Then
                Sql &= "   AND P.SubOperacao = " & pCodigoSuboperacao & vbCrLf
            End If
        End If

        If pCodigoMoeda > 0 Then
            Sql &= "   AND P.Moeda = " & pCodigoMoeda & vbCrLf
        End If

        'Selecao de Produtos User Control Selecao de Produto
        If Not String.IsNullOrEmpty(pSQLProduto) Then
            Sql &= "   AND " & pSQLProduto & vbCrLf
        Else
            Sql &= "   AND Prd.Produto_Id in (" & PProduto & ")" & vbCrLf
        End If



        Sql &= "INSERT INTO #TPosicaoDeContratos(Tipo, Empresa_Id, EndEmpresa_Id, Pedido_Id, Classe, OrigemDestino, Contrato, EntradaSaida, MoedaPedido, Produto, Procuracao, Cedente)" & vbCrLf &
               "Select Tipo, Empresa_Id, EndEmpresa_Id, Pedido_Id, classe, OrigemDestino, Contrato, EntradaSaida, Moeda, Produto_Id,  Procuracao_Id, QtdeSaldo" & vbCrLf &
               "  from #pedidos;" & vbCrLf

        '****************************************************************************************************************************************************************
        '*************************************************  Atualiza Quatidade Fixada e fixada Produracao ***************************************************************
        '****************************************************************************************************************************************************************
        '---- Atualiza Quantidade Fixada Pedidos X Itens X Fixacoes ----
        Sql &= " Update #TPosicaoDeContratos SET" & vbCrLf &
               "     QuantidadeFixado   = consulta.Quantidade" & vbCrLf &
               "    ,ValorFixadoOficial = Consulta.TotalOficial" & vbCrLf &
               "    ,ValorFixadoMoeda   = Consulta.TotalMoeda" & vbCrLf &
               " FROM (SELECT PIF.Empresa_Id," & vbCrLf &
               "              PIF.EndEmpresa_Id," & vbCrLf &
               "              PIF.Pedido_Id,  " & vbCrLf &
               "              PIF.Produto_Id,  " & vbCrLf &
               "              SUM(PIF.Quantidade)    As Quantidade," & vbCrLf &
               "              SUM(PIFE.ValorOficial) As TotalOficial," & vbCrLf &
               "              SUM(PIFE.ValorMoeda)   As TotalMoeda" & vbCrLf &
               "         FROM #TPosicaoDeContratos AS PC " & vbCrLf &
               "        INNER JOIN VW_PedidosXItensXFixacoes PIF" & vbCrLf &
               "           ON PIF.Empresa_Id		= PC.Empresa_Id  " & vbCrLf &
               "           AND PIF.EndEmpresa_Id	= PC.EndEmpresa_Id " & vbCrLf &
               "           AND PIF.Pedido_Id		= PC.Pedido_Id  " & vbCrLf &
               "           AND PIF.Produto_Id		= PC.Produto " & vbCrLf &
               "        INNER JOIN VW_PedidosXItensXFixacoesXEncargos PIFE" & vbCrLf &
               "           ON PIF.Empresa_Id        = PIFE.Empresa_Id " & vbCrLf &
               "          AND PIF.EndEmpresa_Id     = PIFE.EndEmpresa_Id " & vbCrLf &
               "          AND PIF.Pedido_Id         = PIFE.Pedido_Id " & vbCrLf &
               "          AND PIF.Produto_Id        = PIFE.Produto_Id" & vbCrLf &
               "          AND PIF.Fixacao_Id        = PIFE.Fixacao_Id" & vbCrLf &
               "        Where PIFE.Encargo_Id       = 'LIQUIDO'" & vbCrLf &
               "          AND PIF.Movimento         <=" & Data & vbCrLf &
               "        GROUP BY PIF.Empresa_Id, PIF.EndEmpresa_Id, PIF.Pedido_Id, PIF.Produto_Id " & vbCrLf &
               "       ) AS consulta" & vbCrLf &
               " INNER JOIN #TPosicaoDeContratos " & vbCrLf &
               "    ON Consulta.Pedido_Id     = #TPosicaoDeContratos.Pedido_Id " & vbCrLf &
               "   AND consulta.Empresa_Id    = #TPosicaoDeContratos.Empresa_Id " & vbCrLf &
               "   AND consulta.EndEmpresa_Id = #TPosicaoDeContratos.EndEmpresa_Id" & vbCrLf &
               "   AND consulta.Produto_Id    = #TPosicaoDeContratos.Produto" & vbCrLf

        Sql &= vbCrLf

        '****************************************************************************************************************************************************************
        '****************************************************************  Atualiza Contratado **************************************************************************
        '****************************************************************************************************************************************************************
        '----Atualiza Quantidade Contratada Itens do Pedido----
        Sql &= " Update #TPosicaoDeContratos SET" & vbCrLf &
               "     Contratado      = consulta.Quantidade" & vbCrLf &
               "    ,UnitarioPedido  = consulta.unitarioPedido" & vbCrLf &
               "    ,UnitarioOficial = consulta.UnitarioOficial" & vbCrLf &
               "    ,UnitarioMoeda   = consulta.UnitarioMoeda" & vbCrLf &
               "    ,TotalOficial    = consulta.TotalOficial" & vbCrLf &
               "    ,TotalMoeda      = consulta.TotalMoeda" & vbCrLf &
               " FROM (SELECT PxI.Empresa_Id," & vbCrLf &
               "              PxI.EndEmpresa_Id," & vbCrLf &
               "              PxI.Pedido_Id," & vbCrLf &
               "              PxI.Produto_Id, " & vbCrLf &
               "              convert(Decimal(18,5), 0) as UnitarioPedido," & vbCrLf &
               "              convert(Decimal(18,5), 0) as UnitarioOficial," & vbCrLf &
               "              convert(Decimal(18,5), 0) as UnitarioMoeda,  " & vbCrLf &
               "              SUM(CASE WHEN PxI.TipoDeLancamento = 'E' THEN PxI.Quantidade   * - 1 ELSE PxI.Quantidade   END) AS Quantidade,  " & vbCrLf &
               "              SUM(CASE WHEN PxI.TipoDeLancamento = 'E' THEN PxI.TotalOficial * - 1 ELSE PxI.TotalOficial END) AS TotalOficial,  " & vbCrLf &
               "              SUM(CASE WHEN PxI.TipoDeLancamento = 'E' THEN PxI.TotalMoeda   * - 1 ELSE PxI.TotalMoeda   END) AS TotalMoeda  " & vbCrLf &
               "         FROM PedidoXItemXLancamento PxI " & vbCrLf &
               "        Where PXI.Movimento <=" & Data & vbCrLf &
               "        GROUP BY PxI.Empresa_Id, PxI.EndEmpresa_Id, PxI.Pedido_Id, PxI.Produto_Id" & vbCrLf &
               "       ) AS Consulta" & vbCrLf &
               " INNER JOIN #TPosicaoDeContratos P " & vbCrLf &
               "    ON consulta.Pedido_Id     = P.Pedido_Id" & vbCrLf &
               "   AND consulta.Empresa_Id    = P.Empresa_Id" & vbCrLf &
               "   AND consulta.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf &
               "   AND consulta.produto_Id    = P.Produto;" & vbCrLf

        Sql &= vbCrLf

        '****************************************************************************************************************************************************************
        '****************************************************************  Atualiza Entregue ****************************************************************************
        '****************************************************************************************************************************************************************
        '----Atualiza Quantidade Entregue Notas Fiscais X Itens ----
        Sql &= " Update #TPosicaoDeContratos Set " & vbCrLf &
               "     #TPosicaoDeContratos.Entregue = Consulta.Entregue" & vbCrLf &
               " from (SELECT NF.Empresa_Id," & vbCrLf &
               "              NF.EndEmpresa_Id," & vbCrLf &
               "              NF.Pedido," & vbCrLf &
               "              SUM(CASE" & vbCrLf &
               "                    WHEN SO.Devolucao = 'S'" & vbCrLf &
               "                      THEN NFxI.QuantidadeFiscal * - 1" & vbCrLf &
               "                      ELSE NFxI.QuantidadeFiscal" & vbCrLf &
               "                   END) AS Entregue" & vbCrLf &
               "         FROM NotasFiscais NF" & vbCrLf &
               "        INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf &
               "           ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
               "          AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
               "          AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
               "          AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
               "          And NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
               "          AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
               "          AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
               "        INNER JOIN SubOperacoes SO" & vbCrLf &
               "           ON SO.Operacao_Id     = NF.Operacao" & vbCrLf &
               "          AND SO.SubOperacoes_Id = NF.SubOperacao   " & vbCrLf &
               "        INNER JOIN #TPosicaoDeContratos PC" & vbCrLf &
               "           ON NF.Pedido        = PC.Pedido_Id  " & vbCrLf &
               "          AND NF.Empresa_Id    = PC.Empresa_Id  " & vbCrLf &
               "          AND NF.EndEmpresa_Id = PC.EndEmpresa_Id  " & vbCrLf &
               "        WHERE NF.Situacao in (1,4,7) " & vbCrLf &
               "          AND NF.TipoDeDocumento = 1 " & vbCrLf &
               "          AND SO.Classe NOT IN ('" & eClassesOperacoes.CONTAEORDEM.ToString & "', '" & eClassesOperacoes.GLOBAL.ToString & "', '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "')  " & vbCrLf &
               "          AND NF.Movimento <=" & Data & vbCrLf &
               "        GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido" & vbCrLf &
               "       )as Consulta  " & vbCrLf &
               "  Where #TPosicaoDeContratos.Pedido_Id = Consulta.Pedido  " & vbCrLf &
               "    AND #TPosicaoDeContratos.Empresa_Id = Consulta.Empresa_Id  " & vbCrLf &
               "    AND #TPosicaoDeContratos.EndEmpresa_Id = Consulta.EndEmpresa_Id;" & vbCrLf

        Sql &= vbCrLf

        '****************************************************************************************************************************************************************
        '****************************************************************  Atualiza Laudo *******************************************************************************
        '****************************************************************************************************************************************************************
        '----Atualiza Quantidade Entregue Romaneios ----
        Sql &= " Update #TPosicaoDeContratos Set " & vbCrLf &
               "     #TPosicaoDeContratos.Laudo = Consulta.Laudo" & vbCrLf &
               " from (SELECT NF.Empresa_Id," & vbCrLf &
               "              NF.EndEmpresa_Id," & vbCrLf &
               "              NF.Pedido," & vbCrLf &
               "              SUM(CASE" & vbCrLf &
               "                    WHEN SO.Devolucao = 'S'" & vbCrLf &
               "                      THEN R.PesoLiquido * - 1" & vbCrLf &
               "                      ELSE R.PesoLiquido" & vbCrLf &
               "                  END) AS Laudo  " & vbCrLf &
               "         FROM NotasFiscais NF" & vbCrLf &
               "        INNER JOIN SubOperacoes SO" & vbCrLf &
               "           ON NF.Operacao    = SO.Operacao_Id" & vbCrLf &
               "          AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
               "        INNER JOIN NotasFiscaisXRomaneios NFxR" & vbCrLf &
               "           ON NF.Empresa_Id      = NFxR.Empresa_Id" & vbCrLf &
               "          AND NF.EndEmpresa_Id   = NFxR.EndEmpresa_Id" & vbCrLf &
               "          AND NF.Cliente_Id      = NFxR.Cliente_Id" & vbCrLf &
               "          AND NF.EndCliente_Id   = NFxR.EndCliente_Id" & vbCrLf &
               "          AND NF.EntradaSaida_Id = NFxR.EntradaSaida_Id" & vbCrLf &
               "          AND NF.Serie_Id        = NFxR.Serie_Id" & vbCrLf &
               "          AND NF.Nota_Id         = NFxR.Nota_Id   " & vbCrLf &
               "        INNER JOIN Romaneios R" & vbCrLf &
               "           ON NFxR.Empresa_Id    = R.Empresa_Id" & vbCrLf &
               "          AND NFxR.EndEmpresa_Id = R.EndEmpresa_Id   " & vbCrLf &
               "          AND NFxR.Romaneio_Id   = R.Romaneio_Id   " & vbCrLf &
               "        INNER JOIN #TPosicaoDeContratos PC" & vbCrLf &
               "           ON NF.Pedido        = PC.Pedido_Id  " & vbCrLf &
               "          AND NF.Empresa_Id    = PC.Empresa_Id  " & vbCrLf &
               "          AND NF.EndEmpresa_Id = PC.EndEmpresa_Id  " & vbCrLf &
               "        Where NF.Situacao in (1,4,7)  " & vbCrLf &
               "          AND NF.TipoDeDocumento = 1 " & vbCrLf &
               "          AND NF.Movimento <=" & Data & vbCrLf &
               "        GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido" & vbCrLf &
               "       )as Consulta" & vbCrLf &
               "  Where #TPosicaoDeContratos.Pedido_Id     = Consulta.Pedido  " & vbCrLf &
               "    AND #TPosicaoDeContratos.Empresa_Id    = Consulta.Empresa_Id  " & vbCrLf &
               "    AND #TPosicaoDeContratos.EndEmpresa_Id = Consulta.EndEmpresa_Id;" & vbCrLf

        Sql &= vbCrLf

        '********************************************************************************************************************************************************************************
        '********************************************************************************************************************************************************************************        
        Sql &= " UPDATE #TPosicaoDeContratos SET" & vbCrLf &
               "     Contratado = isnull(Laudo,0) + isnull(Cessionario,0)" & vbCrLf &
               "  WHERE isnull(Contratado,0) = 0;" & vbCrLf

        Sql &= vbCrLf

        Sql &= "UPDATE #TPosicaoDeContratos SET" & vbCrLf &
               "    Contratado = Case" & vbCrLf &
               "                   when isnull(Laudo,0) + isnull(Cessionario,0) > isnull(Contratado,0)" & vbCrLf &
               "                     then isnull(Laudo,0) + isnull(Cessionario,0)" & vbCrLf &
               "                     else isnull(Contratado,0)" & vbCrLf &
               "                 end;" & vbCrLf

        Sql &= vbCrLf

        If Left(strEmpresas, 8) = "04440724" Then
            Sql &= " UPDATE #TPosicaoDeContratos SET " & vbCrLf &
                   "    AEntregar = isnull(Contratado,0) - (Case " & vbCrLf &
                   "                                          when #TPosicaoDeContratos.EntradaSaida = 'E' " & vbCrLf &
                   "                                            then Case " & vbCrLf &
                   "                                                   when p.FreteCifFob = 'FOB' " & vbCrLf &
                   "                                                     then ISNULL(#TPosicaoDeContratos.Entregue, 0) " & vbCrLf &
                   "                                                     else ISNULL(#TPosicaoDeContratos.Laudo, 0) " & vbCrLf &
                   "                                                End " & vbCrLf &
                   "                                            else Case " & vbCrLf &
                   "                                                   when p.FreteCifFob = 'FOB' " & vbCrLf &
                   "                                                     then ISNULL(#TPosicaoDeContratos.Entregue, 0) " & vbCrLf &
                   "                                                     else ISNULL(#TPosicaoDeContratos.Entregue, 0) " & vbCrLf &
                   "                                                 End " & vbCrLf &
                   "                                        end),  " & vbCrLf &
                   "    AFixar = Case " & vbCrLf &
                   "               When isnull(#TPosicaoDeContratos.QuantidadeFixado,0) > case  " & vbCrLf &
                   "                                                                        when isnull(#TPosicaoDeContratos.Laudo,0) = 0 " & vbCrLf &
                   "                                                                          then isnull(#TPosicaoDeContratos.Entregue,0)  " & vbCrLf &
                   "                                                                          else isnull(#TPosicaoDeContratos.Laudo,0)  " & vbCrLf &
                   "                                                                       End " & vbCrLf &
                   "                then 0 " & vbCrLf &
                   "                else case " & vbCrLf &
                   "                       when isnull(#TPosicaoDeContratos.Laudo,0) = 0 " & vbCrLf &
                   "                         then isnull(#TPosicaoDeContratos.Entregue,0) " & vbCrLf &
                   "                         else isnull(#TPosicaoDeContratos.Laudo,0) " & vbCrLf &
                   "                    end - isnull(#TPosicaoDeContratos.QuantidadeFixado,0) " & vbCrLf &
                   "             End " & vbCrLf &
                   "  FROM #TPosicaoDeContratos " & vbCrLf &
                   " INNER JOIN Pedidos p     " & vbCrLf &
                   "    ON p.Empresa_Id    = #TPosicaoDeContratos.Empresa_Id" & vbCrLf &
                   "   AND p.EndEmpresa_Id = #TPosicaoDeContratos.EndEmpresa_Id" & vbCrLf &
                   "   AND p.Pedido_Id     = #TPosicaoDeContratos.Pedido_Id" & vbCrLf
        Else
            Sql &= " UPDATE #TPosicaoDeContratos SET" & vbCrLf &
               "    AEntregar = isnull(Contratado,0) - (case " & vbCrLf &
               "											when isnull(#TPosicaoDeContratos.Laudo,0) = 0 " & vbCrLf &
               "												then isnull(#TPosicaoDeContratos.Entregue,0) " & vbCrLf &
               "												else isnull(#TPosicaoDeContratos.Laudo,0) " & vbCrLf &
               "											end + isnull(Cessionario,0)), " & vbCrLf &
               "       AFixar = Case " & vbCrLf &
               "                  When isnull(#TPosicaoDeContratos.QuantidadeFixado,0) > case " & vbCrLf &
               "																			when isnull(#TPosicaoDeContratos.Laudo,0) = 0 " & vbCrLf &
               "																				then isnull(#TPosicaoDeContratos.Entregue,0) " & vbCrLf &
               "																				else isnull(#TPosicaoDeContratos.Laudo,0) " & vbCrLf &
               "																			end " & vbCrLf &
               "                    then 0 " & vbCrLf &
               "                    else " & vbCrLf &
               "						case " & vbCrLf &
               "							when isnull(#TPosicaoDeContratos.Laudo,0) = 0 " & vbCrLf &
               "								then isnull(#TPosicaoDeContratos.Entregue,0) " & vbCrLf &
               "								else isnull(#TPosicaoDeContratos.Laudo,0) " & vbCrLf &
               "							end - isnull(#TPosicaoDeContratos.QuantidadeFixado,0) " & vbCrLf &
               "                end; " & vbCrLf
        End If

        Sql &= vbCrLf

        'Contratos zerados'
        If DeletarContratoZerado Then
            Sql &= " Delete #TPosicaoDeContratos" & vbCrLf &
                   "  where isnull(Contratado,0) = 0" & vbCrLf &
                   "    and isnull(Entregue, 0)  = 0;" & vbCrLf
        End If

        Sql &= vbCrLf


        '**********************************************************************************************************
        '*************** CONTAS A RECEBER QUE NAO SEJA ADIANTAMENTO BAIXA OU TENHA A SEGUNDA CARTEIRA   ***********
        '**********************************************************************************************************
        Sql &= "Select Empresa_id, EndEmpresa_id, Pedido_id, cr, ParcelaOficial, ParcelaMoeda" & vbCrLf &
               "  into #VW_TituloVirtual" & vbCrLf &
               "  from VW_TituloVirtual" & vbCrLf

        Sql &= "Update #TPosicaoDeContratos SET" & vbCrLf &
               "    #TPosicaoDeContratos.Programado_ReceberOficial = Consulta.ProgramadoOficial," & vbCrLf &
               "    #TPosicaoDeContratos.Baixado_ReceberOficial    = Consulta.BaixadoOficial," & vbCrLf &
               "    #TPosicaoDeContratos.Programado_ReceberMoeda   = Consulta.ProgramadoMoeda," & vbCrLf &
               "    #TPosicaoDeContratos.Baixado_ReceberMoeda      = Consulta.BaixadoMoeda" & vbCrLf &
               "  FROM (" & vbCrLf &
               "        Select Empresa, EndEmpresa, Pedido," & vbCrLf &
               "               sum(ProgramadoOficial) as ProgramadoOficial ," & vbCrLf &
               "               sum(BaixadoOficial)    as BaixadoOficial," & vbCrLf &
               "               sum(ProgramadoMoeda)   as ProgramadoMoeda," & vbCrLf &
               "               sum(BaixadoMoeda)      as BaixadoMoeda" & vbCrLf &
               "          From (" & vbCrLf &
               "                 SELECT CR.Empresa," & vbCrLf &
               "                        CR.EndEmpresa," & vbCrLf &
               "                        CR.Pedido," & vbCrLf &
               "                        ISNULL((CASE WHEN CR.Provisao = 2" & vbCrLf &
               "                                          OR ( CR.Provisao = 1" & vbCrLf &
               "                                               and CR.Baixa    > " & Data & vbCrLf &
               "                                               )" & vbCrLf &
               "                                          THEN CR.ValorDoDocumento" & vbCrLf &
               "                                          ELSE 0" & vbCrLf &
               "                                   END), 0) AS ProgramadoOficial," & vbCrLf &
               "                       ISNULL((CASE WHEN  CR.Provisao = 1" & vbCrLf &
               "                                        AND CR.Baixa   <= " & Data & vbCrLf &
               "                                         THEN CR.ValorDoDocumento" & vbCrLf &
               "                                         ELSE 0" & vbCrLf &
               "                                  END), 0) AS BaixadoOficial," & vbCrLf &
               "                        ISNULL((CASE WHEN CR.Provisao = 2" & vbCrLf &
               "                                          OR (CR.Provisao = 1" & vbCrLf &
               "                                               and CR.Baixa    > " & Data & vbCrLf &
               "                                               )" & vbCrLf &
               "                                          THEN CR.MoedaValorDoDocumento" & vbCrLf &
               "                                          ELSE 0" & vbCrLf &
               "                                   END), 0) AS ProgramadoMoeda," & vbCrLf &
               "                       ISNULL((CASE WHEN CR.Provisao = 1" & vbCrLf &
               "                                        AND CR.Baixa   <= " & Data & vbCrLf &
               "                                         THEN CR.MoedaValorDoDocumento" & vbCrLf &
               "                                         ELSE 0" & vbCrLf &
               "                                  END), 0) AS BaixadoMoeda" & vbCrLf &
               "                   FROM ContasAReceber CR" & vbCrLf &
               "                  INNER JOIN #TPosicaoDeContratos PC" & vbCrLf &
               "                     ON CR.Pedido     = PC.Pedido_Id" & vbCrLf &
               "                    AND CR.Empresa    = PC.Empresa_Id" & vbCrLf &
               "                    AND CR.EndEmpresa = PC.EndEmpresa_Id" & vbCrLf &
               "                  inner Join ComprasXProdutos Cart" & vbCrLf &
               "                     on Cart.Produto_Id = CR.Carteira" & vbCrLf &
               "                  WHERE CR.Situacao                     = 1" & vbCrLf &
               "		            AND Cart.Adiantamento               ='N'" & vbCrLf &
               "		            AND cart.BaixaAdiantamento          = 0" & vbCrLf &
               "		            AND len(isnull(cr.CarteiraAdto,'')) = 0" & vbCrLf &
               "                    and isnull(cr.grupado,'')          <>'M'" & vbCrLf &
               "                  UNION ALL" & vbCrLf &
               "                 Select Empresa_id, EndEmpresa_id, Pedido_id, ParcelaOficial, 0, ParcelaMoeda, 0" & vbCrLf &
               "                   from #VW_TituloVirtual" & vbCrLf &
               "                  Where CR = 'R'" & vbCrLf &
               "                ) sb" & vbCrLf &
               "           Group by Empresa, EndEmpresa, Pedido" & vbCrLf &
               "        ) AS Consulta" & vbCrLf &
               "  INNER JOIN #TPosicaoDeContratos" & vbCrLf &
               "     ON Consulta.Pedido     = #TPosicaoDeContratos.Pedido_Id" & vbCrLf &
               "    AND Consulta.Empresa    = #TPosicaoDeContratos.Empresa_Id" & vbCrLf &
               "    AND Consulta.EndEmpresa = #TPosicaoDeContratos.EndEmpresa_Id;" & vbCrLf


        '**********************************************************************************************************
        '*************** CONTAS A RECEBER ADIANTAMENTO ou BAIXA De Adiantamento ou Tenha A SEGUNDA CARTEIRA   *****
        '**********************************************************************************************************
        Sql &= "Update #TPosicaoDeContratos SET" & vbCrLf &
               "	#TPosicaoDeContratos.AdiantamentoOficial        = Consulta.AdiantamentoOficial," & vbCrLf &
               "    #TPosicaoDeContratos.AdiantamentoMoeda          = Consulta.AdiantamentoMoeda," & vbCrLf &
               "    #TPosicaoDeContratos.BaixaAdiantamentoOficial   = Consulta.BaixaAdtoOficial," & vbCrLf &
               "    #TPosicaoDeContratos.BaixaAdiantamentoMoeda     = Consulta.BaixaAdtoMoeda" & vbCrLf &
               "  FROM (" & vbCrLf &
               "        SELECT CR.Empresa," & vbCrLf &
               "               CR.EndEmpresa," & vbCrLf &
               "               CR.Pedido," & vbCrLf &
               "              ISNULL(SUM(CASE WHEN (cart.BaixaAdiantamento = 0 AND len(isnull(cr.CarteiraAdto,'')) = 0)" & vbCrLf &
               "                               AND CR.Provisao = 1" & vbCrLf &
               "                               AND CR.Baixa   <= " & Data & vbCrLf &
               "                                THEN CR.ValorDoDocumento" & vbCrLf &
               "                                ELSE 0" & vbCrLf &
               "                         END), 0) AS AdiantamentoOficial," & vbCrLf &
               "              ISNULL(SUM(CASE WHEN (cart.BaixaAdiantamento = 0 AND len(isnull(cr.CarteiraAdto,'')) = 0)" & vbCrLf &
               "                               AND CR.Provisao = 1" & vbCrLf &
               "                               AND CR.Baixa   <= " & Data & vbCrLf &
               "                                THEN CR.MoedaValorDoDocumento" & vbCrLf &
               "                                ELSE 0" & vbCrLf &
               "                         END), 0) AS AdiantamentoMoeda," & vbCrLf &
               "              ISNULL(SUM(CASE WHEN cart.BaixaAdiantamento = 1 or (CartBaixa.baixaadiantamento is not null and CartBaixa.baixaadiantamento = 1)" & vbCrLf &
               "                               AND CR.Provisao = 1" & vbCrLf &
               "                               AND CR.Baixa   <= " & Data & vbCrLf &
               "                                THEN CR.ValorLiquido" & vbCrLf &
               "                                ELSE 0" & vbCrLf &
               "                         END), 0) AS BaixaAdtoOficial," & vbCrLf &
               "              ISNULL(SUM(CASE WHEN cart.BaixaAdiantamento = 1 or (CartBaixa.baixaadiantamento is not null and CartBaixa.baixaadiantamento = 1)" & vbCrLf &
               "                               AND CR.Provisao = 1" & vbCrLf &
               "                               AND CR.Baixa   <= " & Data & vbCrLf &
               "                                THEN CR.MoedaValorLiquido" & vbCrLf &
               "                                ELSE 0" & vbCrLf &
               "                         END), 0) AS BaixaAdtoMoeda" & vbCrLf &
               "          FROM ContasAReceber CR" & vbCrLf &
               "         INNER JOIN #TPosicaoDeContratos PC" & vbCrLf &
               "            ON CR.Pedido     = PC.Pedido_Id" & vbCrLf &
               "           AND CR.Empresa    = PC.Empresa_Id " & vbCrLf &
               "           AND CR.EndEmpresa = PC.EndEmpresa_Id" & vbCrLf &
               "         inner Join ComprasXProdutos Cart" & vbCrLf &
               "            on Cart.Produto_Id = CR.Carteira" & vbCrLf &
               "         Left Join ComprasXProdutos CartBaixa" & vbCrLf &
               "            on CartBaixa.Produto_Id = CR.CarteiraAdto" & vbCrLf &
               "         WHERE CR.Situacao     = 1" & vbCrLf &
               "		   AND (   Cart.Adiantamento               = 'S'" & vbCrLf &
               "		        OR cart.BaixaAdiantamento          = 1 " & vbCrLf &
               "		        OR len(isnull(cr.CarteiraAdto,'')) > 0)" & vbCrLf &
               "         GROUP BY CR.Empresa, CR.EndEmpresa, CR.Pedido" & vbCrLf &
               "        ) AS Consulta" & vbCrLf &
               "  INNER JOIN #TPosicaoDeContratos" & vbCrLf &
               "     ON Consulta.Pedido     = #TPosicaoDeContratos.Pedido_Id" & vbCrLf &
               "    AND Consulta.Empresa    = #TPosicaoDeContratos.Empresa_Id" & vbCrLf &
               "    AND Consulta.EndEmpresa = #TPosicaoDeContratos.EndEmpresa_Id;" & vbCrLf


        '**********************************************************************************************************
        '*************** CONTAS A PAGAR QUE NAO SEJA ADIANTAMENTO BAIXA OU TENHA A SEGUNDA CARTEIRA   ***********
        '**********************************************************************************************************

        Sql &= "Update #TPosicaoDeContratos SET" & vbCrLf &
               "    #TPosicaoDeContratos.Programado_PagarOficial  = Consulta.ProgramadoOficial," & vbCrLf &
               "    #TPosicaoDeContratos.Baixado_PagarOficial     = Consulta.BaixadoOficial," & vbCrLf &
               "    #TPosicaoDeContratos.Programado_PagarMoeda    = Consulta.ProgramadoMoeda," & vbCrLf &
               "    #TPosicaoDeContratos.Baixado_PagarMoeda       = Consulta.BaixadoMoeda" & vbCrLf &
               "  FROM (" & vbCrLf &
               "        Select Empresa, EndEmpresa, Pedido," & vbCrLf &
               "               sum(ProgramadoOficial) as ProgramadoOficial ," & vbCrLf &
               "               sum(BaixadoOficial)    as BaixadoOficial," & vbCrLf &
               "               sum(ProgramadoMoeda)   as ProgramadoMoeda," & vbCrLf &
               "               sum(BaixadoMoeda)      as BaixadoMoeda" & vbCrLf &
               "          From (" & vbCrLf &
               "                 SELECT CR.Empresa," & vbCrLf &
               "                        CR.EndEmpresa," & vbCrLf &
               "                        CR.Pedido," & vbCrLf &
               "                        ISNULL((CASE WHEN CR.Provisao = 2" & vbCrLf &
               "                                          OR ( CR.Provisao = 1" & vbCrLf &
               "                                               and CR.Baixa > " & Data & vbCrLf &
               "                                               )" & vbCrLf &
               "                                          THEN CR.ValorDoDocumento" & vbCrLf &
               "                                          ELSE 0" & vbCrLf &
               "                                   END), 0) AS ProgramadoOficial," & vbCrLf &
               "                       ISNULL((CASE WHEN CR.Provisao = 1" & vbCrLf &
               "                                        AND CR.Baixa <= " & Data & vbCrLf &
               "                                         THEN CR.ValorDoDocumento" & vbCrLf &
               "                                         ELSE 0" & vbCrLf &
               "                                  END), 0)  AS BaixadoOficial," & vbCrLf &
               "                        ISNULL((CASE WHEN CR.Provisao = 2" & vbCrLf &
               "                                          OR (" & vbCrLf &
               "                                                CR.Provisao = 1" & vbCrLf &
               "                                               and CR.Baixa > " & Data & vbCrLf &
               "                                               )" & vbCrLf &
               "                                          THEN CR.MoedaValorDoDocumento" & vbCrLf &
               "                                          ELSE 0" & vbCrLf &
               "                                   END), 0) AS ProgramadoMoeda," & vbCrLf &
               "                       ISNULL((CASE WHEN  CR.Provisao = 1 " & vbCrLf &
               "                                        AND CR.Baixa <= " & Data & vbCrLf &
               "                                         THEN CR.MoedaValorDoDocumento" & vbCrLf &
               "                                         ELSE 0" & vbCrLf &
               "                                  END), 0) AS BaixadoMoeda" & vbCrLf &
               "                   FROM ContasAPagar CR" & vbCrLf &
               "                  INNER JOIN #TPosicaoDeContratos PC" & vbCrLf &
               "                     ON CR.Pedido     = PC.Pedido_Id" & vbCrLf &
               "                    AND CR.Empresa    = PC.Empresa_Id" & vbCrLf &
               "                    AND CR.EndEmpresa = PC.EndEmpresa_Id" & vbCrLf &
               "                  inner Join ComprasXProdutos Cart" & vbCrLf &
               "                     on Cart.Produto_Id = CR.Carteira" & vbCrLf &
               "                   Left Join ComprasXProdutos CartBaixa" & vbCrLf &
               "                     on CartBaixa.Produto_Id = CR.CarteiraAdto" & vbCrLf &
               "                  WHERE CR.Situacao           = 1" & vbCrLf &
               "		            and Cart.Adiantamento     = 'N'" & vbCrLf &
               "		            AND (cart.BaixaAdiantamento = 0 or (CartBaixa.baixaadiantamento is not null and CartBaixa.baixaadiantamento = 0))" & vbCrLf &
               "                    and isnull(cr.grupado,'') <> 'M'" & vbCrLf &
               "                  UNION ALL" & vbCrLf &
               "                 Select Empresa_id, EndEmpresa_id, Pedido_id, ParcelaOficial, 0, ParcelaMoeda, 0" & vbCrLf &
               "                   from #VW_TituloVirtual" & vbCrLf &
               "                  Where CR = 'P'" & vbCrLf &
               "                ) sb" & vbCrLf &
               "           Group by Empresa, EndEmpresa, Pedido" & vbCrLf &
               "        ) AS Consulta" & vbCrLf &
               "  INNER JOIN #TPosicaoDeContratos" & vbCrLf &
               "     ON Consulta.Pedido     = #TPosicaoDeContratos.Pedido_Id " & vbCrLf &
               "    AND Consulta.Empresa    = #TPosicaoDeContratos.Empresa_Id " & vbCrLf &
               "    AND Consulta.EndEmpresa = #TPosicaoDeContratos.EndEmpresa_Id;" & vbCrLf



        '**********************************************************************************************************
        '*************** CONTAS A Pagar ADIANTAMENTO ou BAIXA De Adiantamento ou Tenha A SEGUNDA CARTEIRA   *****
        '**********************************************************************************************************
        Sql &= "Update #TPosicaoDeContratos SET" & vbCrLf &
               "    #TPosicaoDeContratos.AdiantamentoOficial      = isnull(#TPosicaoDeContratos.AdiantamentoOficial,0)      + Consulta.AdiantamentoOficial," & vbCrLf &
               "    #TPosicaoDeContratos.AdiantamentoMoeda        = isnull(#TPosicaoDeContratos.AdiantamentoMoeda,0)        + Consulta.AdiantamentoMoeda," & vbCrLf &
               "	#TPosicaoDeContratos.BaixaAdiantamentoOficial = isnull(#TPosicaoDeContratos.BaixaAdiantamentoOficial,0) + Consulta.BaixaAdtoOficial," & vbCrLf &
               "    #TPosicaoDeContratos.BaixaAdiantamentoMoeda   = isnull(#TPosicaoDeContratos.BaixaAdiantamentoMoeda,0)   + Consulta.BaixaAdtoMoeda" & vbCrLf &
               "  FROM (" & vbCrLf &
               "        SELECT CR.Empresa," & vbCrLf &
               "               CR.EndEmpresa," & vbCrLf &
               "               CR.Pedido," & vbCrLf &
               "              ISNULL(SUM(CASE WHEN (cart.BaixaAdiantamento = 0 AND len(isnull(cr.CarteiraAdto,'')) = 0)" & vbCrLf &
               "                               AND CR.Provisao = 1" & vbCrLf &
               "                               AND CR.Baixa <= " & Data & vbCrLf &
               "                                THEN CR.ValorDoDocumento" & vbCrLf &
               "                                ELSE 0" & vbCrLf &
               "                         END), 0) AS AdiantamentoOficial," & vbCrLf &
               "              ISNULL(SUM(CASE WHEN (cart.BaixaAdiantamento = 0 AND len(isnull(cr.CarteiraAdto,'')) = 0)" & vbCrLf &
               "                               AND CR.Provisao = 1" & vbCrLf &
               "                               AND CR.Baixa <= " & Data & vbCrLf &
               "                                THEN CR.MoedaValorDoDocumento" & vbCrLf &
               "                                ELSE 0" & vbCrLf &
               "                         END), 0) AS AdiantamentoMoeda," & vbCrLf &
               "              ISNULL(SUM(CASE WHEN cart.BaixaAdiantamento = 1 or (CartBaixa.baixaadiantamento is not null and CartBaixa.baixaadiantamento = 1)" & vbCrLf &
               "                               AND CR.Provisao = 1" & vbCrLf &
               "                               AND CR.Baixa <= " & Data & vbCrLf &
               "                                THEN CR.ValorLiquido" & vbCrLf &
               "                                ELSE 0" & vbCrLf &
               "                         END), 0) AS BaixaAdtoOficial," & vbCrLf &
               "              ISNULL(SUM(CASE WHEN cart.BaixaAdiantamento = 1 or (CartBaixa.baixaadiantamento is not null and CartBaixa.baixaadiantamento = 1)" & vbCrLf &
               "                               AND CR.Provisao = 1" & vbCrLf &
               "                               AND CR.Baixa <= " & Data & vbCrLf &
               "                                THEN CR.MoedaValorLiquido" & vbCrLf &
               "                                ELSE 0" & vbCrLf &
               "                         END), 0) AS BaixaAdtoMoeda" & vbCrLf &
               "          FROM ContasAPagar CR" & vbCrLf &
               "         INNER JOIN #TPosicaoDeContratos PC" & vbCrLf &
               "            ON CR.Pedido     = PC.Pedido_Id" & vbCrLf &
               "           AND CR.Empresa    = PC.Empresa_Id" & vbCrLf &
               "           AND CR.EndEmpresa = PC.EndEmpresa_Id" & vbCrLf &
               "         inner Join ComprasXProdutos Cart" & vbCrLf &
               "            on Cart.Produto_Id = CR.Carteira" & vbCrLf &
               "         Left Join ComprasXProdutos CartBaixa" & vbCrLf &
               "            on CartBaixa.Produto_Id = CR.CarteiraAdto" & vbCrLf &
               "         WHERE CR.Situacao     = 1" & vbCrLf &
               "		   and (Cart.Adiantamento = 'S' or cart.BaixaAdiantamento = 1 or len(isnull(cr.CarteiraAdto,'')) > 0)" & vbCrLf &
               "         GROUP BY CR.Empresa, CR.EndEmpresa, CR.Pedido" & vbCrLf &
               "        ) AS Consulta" & vbCrLf &
               "  INNER JOIN #TPosicaoDeContratos" & vbCrLf &
               "     ON Consulta.Pedido     = #TPosicaoDeContratos.Pedido_Id" & vbCrLf &
               "    AND Consulta.Empresa    = #TPosicaoDeContratos.Empresa_Id" & vbCrLf &
               "    AND Consulta.EndEmpresa = #TPosicaoDeContratos.EndEmpresa_Id;" & vbCrLf

        '---- Ajusta Saldo Financeiro ----------------
        Sql &= " Update #TPosicaoDeContratos Set " & vbCrLf &
               "     Programado_SaldoOficial = isnull(Programado_PagarOficial, 0) - isnull(Programado_ReceberOficial, 0)" & vbCrLf &
               "    ,Programado_SaldoMoeda   = isnull(Programado_PagarMoeda, 0)   - isnull(Programado_ReceberMoeda, 0)" & vbCrLf &
               "    ,Baixado_SaldoOficial    = isnull(Baixado_PagarOficial, 0)    - isnull(Baixado_ReceberOficial, 0)" & vbCrLf &
               "    ,Baixado_SaldoMoeda      = IsNull(Baixado_PagarMoeda, 0)      - IsNull(Baixado_ReceberMoeda, 0)" & vbCrLf &
               "  from #TPosicaoDeContratos;   " & vbCrLf

        Sql &= vbCrLf

        Sql &= " Update #TPosicaoDeContratos Set" & vbCrLf &
               "     Programado_SaldoOficial = (Programado_SaldoOficial * -1)" & vbCrLf &
               "   from #TPosicaoDeContratos where Programado_SaldoOficial < 0;" & vbCrLf

        Sql &= vbCrLf

        Sql &= " Update #TPosicaoDeContratos Set" & vbCrLf &
               "    Programado_SaldoMoeda   = (Programado_SaldoMoeda * -1)" & vbCrLf &
               "   from #TPosicaoDeContratos" & vbCrLf &
               "  where Programado_SaldoMoeda < 0;" & vbCrLf

        Sql &= vbCrLf

        Sql &= " Update #TPosicaoDeContratos Set" & vbCrLf &
               "     Baixado_SaldoOficial = (Baixado_SaldoOficial * -1)" & vbCrLf &
               "   from #TPosicaoDeContratos" & vbCrLf &
               "  where Baixado_SaldoOficial < 0;" & vbCrLf

        Sql &= vbCrLf

        Sql &= " Update #TPosicaoDeContratos Set" & vbCrLf &
               "     Baixado_SaldoMoeda   = (Baixado_SaldoMoeda * -1)" & vbCrLf &
               "   from #TPosicaoDeContratos" & vbCrLf &
               "  where Baixado_SaldoMoeda < 0;" & vbCrLf

        Sql &= vbCrLf

        '-- Calcula Quantidade Paga ---------

        Sql &= " Update #TPosicaoDeContratos Set" & vbCrLf &
               "     Pago = CASE" & vbCrLf &
               "              WHEN MoedaPedido = 1" & vbCrLf &
               "                THEN (QuantidadeFixado / ValorFixadoOficial) * Baixado_SaldoOficial  " & vbCrLf &
               "                ELSE (QuantidadeFixado / ValorFixadoMoeda) * Baixado_SaldoMoeda" & vbCrLf &
               "            END  " & vbCrLf &
               "    WHERE  CASE" & vbCrLf &
               "        WHEN MoedaPedido = 1" & vbCrLf &
               "          THEN ValorFixadoOficial" & vbCrLf &
               "          ELSE ValorFixadoMoeda" & vbCrLf &
               "        END > 0" & vbCrLf
        Sql &= vbCrLf

        Sql &= "   Update #TPosicaoDeContratos Set" & vbCrLf &
               "       Adiantamento = CASE" & vbCrLf &
            "   				        WHEN MoedaPedido = 1" & vbCrLf &
            "   				          THEN (isnull(AdiantamentoOficial,0) - isnull(BaixaAdiantamentoOficial,0))" & vbCrLf &
            "   				          ELSE (isnull(AdiantamentoMoeda,0)   - isnull(BaixaAdiantamentoMoeda,0))" & vbCrLf &
            "   				      END" & vbCrLf &
               "      WHERE case" & vbCrLf &
               "              when MoedaPedido = 1" & vbCrLf &
               "                then isnull(AdiantamentoOficial,0) - ISNULL(BaixaAdiantamentoOficial,0)" & vbCrLf &
               "                else isnull(AdiantamentoMoeda,0)   - ISNULL(BaixaAdiantamentoMoeda,0)" & vbCrLf &
               "            end > 0;" & vbCrLf

        Sql &= vbCrLf

        Sql &= " UPDATE #TPosicaoDeContratos Set" & vbCrLf &
               "     RecebidoNaoPago = (Entregue - Pago)" & vbCrLf &
               "  WHERE isnull(Entregue, 0) > Pago;" & vbCrLf

        Sql &= vbCrLf

        Sql &= " UPDATE #TPosicaoDeContratos Set" & vbCrLf &
               "     PagoNaoRecebido = (Pago - isnull(Entregue, 0))" & vbCrLf &
               "  WHERE Pago > isnull(Entregue, 0);" & vbCrLf

        Sql &= vbCrLf

        Sql &= " Update #TPosicaoDeContratos SET" & vbCrLf &
               "     TotalOficial =  (SELECT SUM(CASE" & vbCrLf &
               "                                   WHEN PxI.TipoDeLancamento = 'E'" & vbCrLf &
               "                                     THEN PxI.TotalOficial * - 1" & vbCrLf &
               "                                     ELSE PxI.TotalOficial" & vbCrLf &
               "                                 END) AS Oficial  " & vbCrLf &
               "                        FROM PedidoXItemXLancamento PxI  " & vbCrLf &
               "                       WHERE PxI.Pedido_Id     = #TPosicaoDeContratos.Pedido_Id" & vbCrLf &
               "                         AND PxI.Empresa_Id    = #TPosicaoDeContratos.Empresa_Id" & vbCrLf &
               "                         AND PxI.EndEmpresa_Id = #TPosicaoDeContratos.EndEmpresa_Id" & vbCrLf &
               "                         AND PxI.Movimento    <=" & Data & vbCrLf &
               "                       )" & vbCrLf &
               "    ,TotalMoeda   =  (SELECT SUM(CASE" & vbCrLf &
               "                                   WHEN PxI.TipoDeLancamento = 'E'" & vbCrLf &
               "                                     THEN PxI.TotalMoeda * - 1" & vbCrLf &
               "                                     ELSE PxI.TotalMoeda" & vbCrLf &
               "                                 END) AS Moeda  " & vbCrLf &
               "                        FROM PedidoXItemXLancamento PxI  " & vbCrLf &
               "                       WHERE PxI.Pedido_Id     = #TPosicaoDeContratos.Pedido_Id" & vbCrLf &
               "                         AND PxI.Empresa_Id    = #TPosicaoDeContratos.Empresa_Id" & vbCrLf &
               "                         AND PxI.EndEmpresa_Id = #TPosicaoDeContratos.EndEmpresa_Id" & vbCrLf &
               "                         AND PxI.Movimento    <=" & Data & vbCrLf &
               "                      );" & vbCrLf

        Sql &= vbCrLf

        Sql &= " Update #TPosicaoDeContratos Set" & vbCrLf &
               "     UnitarioOficial = (PC.TotalOficial / PC.Contratado)" & vbCrLf &
               "   FROM #TPosicaoDeContratos PC" & vbCrLf &
               "  INNER JOIN Produtos prd" & vbCrLf &
               "     ON PC.Produto = prd.Produto_Id  " & vbCrLf &
               "  Where PC.TotalOficial   > 0" & vbCrLf &
               "    And PC.Contratado     > 0;  " & vbCrLf

        Sql &= vbCrLf

        Sql &= " Update #TPosicaoDeContratos Set" & vbCrLf &
               "     UnitarioPedido = case when pc.MoedaPedido = 1 then PC.TotalOficial else PC.TotalMoeda end / PC.Contratado" & vbCrLf &
               "   FROM #TPosicaoDeContratos PC" & vbCrLf &
               "  INNER JOIN Produtos prd" & vbCrLf &
               "     ON PC.Produto = Prd.Produto_Id  " & vbCrLf &
               "  Where case when pc.MoedaPedido = 1 then PC.TotalOficial else PC.TotalMoeda end > 0" & vbCrLf &
               "    And PC.Contratado     > 0;" & vbCrLf

        Sql &= vbCrLf

        Sql &= " Update #TPosicaoDeContratos Set" & vbCrLf &
               "     UnitarioMoeda = (PC.TotalMoeda / PC.Contratado)" & vbCrLf &
               "   FROM #TPosicaoDeContratos PC" & vbCrLf &
               "  INNER JOIN Produtos prd" & vbCrLf &
               "     ON PC.Produto = Prd.Produto_Id  " & vbCrLf &
               "  Where PC.TotalMoeda     > 0" & vbCrLf &
               "    And PC.Contratado     > 0;" & vbCrLf

        Sql &= vbCrLf
        If Cessionario Then
            Sql &= "INSERT INTO #TPosicaoDeContratos(Tipo, Empresa_Id, EndEmpresa_Id, Pedido_Id, Classe, OrigemDestino, Contrato, EntradaSaida, MoedaPedido, Produto, Procuracao,  Cessionario)" & vbCrLf &
                   "SELECT 'CES', P.Empresa_Id, P.EndEmpresa_Id, P.Pedido_Id, P.Classe, P.OrigemDestino, P.Contrato, P.EntradaSaida, P.Moeda, P.Produto_Id, PR.Procuracao_Id, pr.QtdeSaldo" & vbCrLf &
                   "  from #Pedidos P" & vbCrLf &
                   " inner join #Procuracoes PR" & vbCrLf &
                   "    on P.Empresa_Id    = PR.Empresa_ID" & vbCrLf &
                   "   and P.EndEmpresa_Id = PR.EndEmpresa_ID" & vbCrLf &
                   "   and P.Pedido_Id     = PR.Pedido_ID" & vbCrLf
        End If

        '------------------------------------------------------------------------------------------------

        Sql &= "Select E.Reduzido    as ReduzidoEmpresa," & vbCrLf &
               "       E.Cliente_Id  as Empresa," & vbCrLf &
               "       E.Endereco_Id as EndEmpresa," & vbCrLf &
               "       E.Nome        as NomeEmpresa," & vbCrLf &
               "       E.Cidade      as CidadeEmpresa," & vbCrLf &
               "       E.Estado      as EstadoEmpresa," & vbCrLf &
               "       C.Cliente_Id  as Cliente," & vbCrLf &
               "       C.Endereco_Id as EndCliente," & vbCrLf &
               "       Case" & vbCrLf &
               "         when len(C.Complemento) > 0" & vbCrLf &
               "           then substring(C.nome,0,20) + ' - ' + replace(C.Complemento,'FAZENDA', 'FAZ.')" & vbCrLf &
               "           else C.nome" & vbCrLf &
               "       end NomeCliente," & vbCrLf &
               "       C.Cidade      as CidadeCliente," & vbCrLf &
               "       C.Estado      as EstadoCliente," & vbCrLf &
               "       PC.Tipo," & vbCrLf &
               "       PC.Procuracao," & vbCrLf &
               "       P.OrigemDestino," & vbCrLf &
               "       P.Contrato," & vbCrLf &
               "       PC.Pedido_Id       as Pedido," & vbCrLf &
               "       P.DataPedido," & vbCrLf &
               "       P.DataEntrega," & vbCrLf &
               "       SO.Operacao_Id     as Operacao," & vbCrLf &
               "       SO.SubOperacoes_Id as SubOperacao," & vbCrLf &
               "       SO.Descricao       as NomeOperacao," & vbCrLf &
               "       P.Moeda," & vbCrLf &
               "       M.Descricao AS NomeMoeda," & vbCrLf &
               "       P.Indexador," & vbCrLf &
               "       I.Descricao AS NomeIndexador," & vbCrLf &
               "       P.Finalidade," & vbCrLf &
               "       F.Descricao AS NomeFinalidade," & vbCrLf &
               "       P.FreteCIFFOB," & vbCrLf &
               "       P.CondicaoPagamento," & vbCrLf &
               "       Pg.Descricao AS NomeCondicaoDePagamento," & vbCrLf &
               "       PC.Produto as Produto," & vbCrLf &
               "       Prd.Nome AS NomeProduto," & vbCrLf &
               "       isnull(PC.Contratado,0)      as Contratado," & vbCrLf &
               "       isnull(PC.UnitarioPedido,0)  as UnitarioPedido," & vbCrLf &
               "       isnull(PC.UnitarioOficial,0) as UnitarioOficial," & vbCrLf &
               "       isnull(PC.TotalOficial,0)    as ValorTotalContratado," & vbCrLf &
               "       isnull(PC.UnitarioMoeda,0)   as UnitarioMoeda," & vbCrLf &
               "       isnull(PC.TotalMoeda,0)      as TotalMoeda," & vbCrLf &
               "       isnull(PC.Laudo,0)           as Laudo," & vbCrLf &
               "       isnull(PC.Entregue,0)        as Entregue," & vbCrLf &
               "       isnull(PC.AEntregar,0)        as AEntregar," & vbCrLf &
               "       isnull(PC.QuantidadeFixado,0) as Fixada," & vbCrLf &
               "       isnull(ValorFixadoOficial,0)  as ValorFixadoOficial," & vbCrLf &
               "       isnull(ValorFixadoMoeda,0)    as ValorFixadoMoeda," & vbCrLf &
               "       isnull(PC.AFixar,0)           as AFixar," & vbCrLf &
               "       isnull(PC.Pago,0)             as Pago," & vbCrLf &
               "       isnull(PC.PagoNaoRecebido,0)  as PagoNaoRecebido," & vbCrLf &
               "       isnull(PC.RecebidoNaoPago,0)  as RecebidoNaoPago," & vbCrLf &
               "       isnull(PC.Adiantamento,0)     as Adiantamento," & vbCrLf &
               "       isnull(PC.Cedente,0)          as Cedente," & vbCrLf &
               "       isnull(PC.Cessionario,0)      as Cessionario," & vbCrLf &
               "       isnull(PDC.Titulo, 'Nenhum') + case when cliod.estado = 'EX' then ' - ME' else ' - MI' end  as TipoDeCliente," & vbCrLf &
               "       NULL as AbrirProcuracoes," & vbCrLf &
               "       case" & vbCrLf &
               "         when (OP.Classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and P.FreteCIFFOB = 'FOB') or (OP.Classe = '" & eClassesOperacoes.VENDAS.ToString & "' and P.FreteCIFFOB = 'CIF')" & vbCrLf &
               "           then P.FreteCIFFOB" & vbCrLf &
               "           else '*'+P.FreteCIFFOB" & vbCrLf &
               "       end FreteCIFFOBRel," & vbCrLf &
               "       case" & vbCrLf &
               "         when (OP.Classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and P.FreteCIFFOB = 'FOB') or (OP.Classe = '" & eClassesOperacoes.VENDAS.ToString & "' and P.FreteCIFFOB = 'CIF')" & vbCrLf &
               "           then convert(nvarchar(10),P.DataEntrega,103)" & vbCrLf &
               "           else convert(nvarchar(10),P.DataEntrega,103)" & vbCrLf &
               "       end as DataEntregaRel," & vbCrLf &
               "       CONVERT(DECIMAL(18,4), PxR.UnitarioMedioDoFrete) AS UnitarioMedioDofrete," & vbCrLf &
               "       CONVERT(DECIMAL(18,4), PxR.UnitarioMedioDoFrete * 1000) AS ValorDoFrete,  " & vbCrLf &
               "       cliod.cidade as CidadeOD, " & vbCrLf &
               "       P.UsuarioInclusao," & vbCrLf &
               "       P.UsuarioAlteracao," & vbCrLf &
               "       P.UsuarioLiberacao" & vbCrLf

        If Not String.IsNullOrEmpty(TabelaTemporaria) Then
            Sql &= " into " & TabelaTemporaria
        End If

        Sql &= "  FROM #TPosicaoDeContratos PC" & vbCrLf &
               " inner join Pedidos P" & vbCrLf &
               "    ON P.Pedido_Id     = PC.Pedido_Id" & vbCrLf &
               "   AND P.Empresa_Id    = PC.Empresa_Id" & vbCrLf &
               "   AND P.EndEmpresa_Id = PC.EndEmpresa_Id" & vbCrLf &
               "  LEFT JOIN (SELECT Empresa_Id, EndEmpresa_Id, Pedido_Id, AVG(Valor) AS UnitarioMedioDoFrete" & vbCrLf &
               "		       FROM PedidoXRoteiro " & vbCrLf &
               "		      GROUP BY Empresa_Id, EndEmpresa_Id, Pedido_Id" & vbCrLf &
               "		     ) as PxR" & vbCrLf &
               "    ON P.Empresa_Id = PxR.Empresa_Id" & vbCrLf &
               "   AND P.EndEmpresa_Id = PxR.EndEmpresa_Id" & vbCrLf &
               "   AND P.Pedido_Id = PxR.Pedido_Id" & vbCrLf &
               " Inner JOIN Indexadores I" & vbCrLf &
               "    ON I.Indexador_Id  = P.Indexador" & vbCrLf &
               " Inner JOIN Moedas M" & vbCrLf &
               "    ON P.Moeda = M.Moeda_Id" & vbCrLf &
               " Inner JOIN Finalidades F" & vbCrLf &
               "    ON P.Finalidade = F.Finalidade_Id" & vbCrLf &
               " Inner join Operacoes OP" & vbCrLf &
               "    ON OP.Operacao_id = P.operacao" & vbCrLf &
               " Inner JOIN SubOperacoes SO" & vbCrLf &
               "    ON P.Operacao     = SO.Operacao_Id" & vbCrLf &
               "   AND P.SubOperacao  = SO.SubOperacoes_Id" & vbCrLf &
               "  left JOIN PlanoDeContas PDC" & vbCrLf &
               "    ON PDC.Conta_Id = SO.GrupoDeContas" & vbCrLf &
               " Inner JOIN Clientes AS E" & vbCrLf &
               "    ON P.Empresa_Id    = E.Cliente_Id" & vbCrLf &
               "   AND P.EndEmpresa_Id = E.Endereco_Id" & vbCrLf &
               "  Left Join #Procuracoes Pro" & vbCrLf &
               "    ON Pro.Empresa_Id    = PC.Empresa_Id" & vbCrLf &
               "   AND Pro.EndEmpresa_Id = PC.EndEmpresa_Id" & vbCrLf &
               "   AND Pro.Procuracao_ID = PC.Procuracao" & vbCrLf &
               " inner Join Clientes C" & vbCrLf &
               "    ON C.Cliente_Id  = isnull(pro.ClienteCessionario, P.Cliente)" & vbCrLf &
               "   AND C.Endereco_Id = isnull(pro.EndCessionario, P.EndCliente)" & vbCrLf &
               " inner JOIN Produtos Prd" & vbCrLf &
               "    ON PC.Produto = Prd.Produto_Id" & vbCrLf &
               "   LEFT JOIN Pagamentos Pg" & vbCrLf &
               "    ON P.CondicaoPagamento = Pg.Pagamento_Id" & vbCrLf &
               "  LEFT Join PedidosXDepositos pxd" & vbCrLf &
               "    ON pxd.pedido_id     = P.Pedido_Id" & vbCrLf &
               "   AND pxd.Empresa_Id    = P.Empresa_Id" & vbCrLf &
               "   AND Pxd.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf &
               "   and pxd.Tipo          = 'OD'" & vbCrLf &
               "  Left Join Clientes cliod" & vbCrLf &
               "    on cliod.Cliente_Id  = ISNULL(pxd.Deposito_Id, p.Cliente)" & vbCrLf &
               "   and cliod.Endereco_Id = ISNULL(pxd.EndDeposito_Id, p.endCliente)" & vbCrLf &
               " Where Prd.Agrupar = 'N'" & vbCrLf

        If Not Cliente Is Nothing Then
            If ClienteConsolidado Then
                Sql &= "   AND left(C.Cliente_Id,8)    ='" & Left(Cliente.Codigo, 8) & "'" & vbCrLf
            Else
                Sql &= "   AND C.Cliente_Id  ='" & Cliente.Codigo & "'" & vbCrLf &
                       "   AND C.Endereco_id = " & Cliente.CodigoEndereco & vbCrLf
            End If
        End If

        If pSaldo = 1 Then
            Sql &= " AND  (round(PC.AEntregar,0) <> 0 or round(PC.AFixar,0) <> 0 or round(PC.PagoNaoRecebido,0) <> 0 or round(PC.RecebidoNaoPago,0) <> 0 or round(PC.Adiantamento,0) <> 0  or round(PC.cedente,0) <> 0  or round(PC.cessionario,0) <> 0) " & vbCrLf
        ElseIf pSaldo = 2 Then
            Sql &= " AND  (round(PC.AEntregar,0) = 0 and round(PC.AFixar,0) = 0 and round(PC.PagoNaoRecebido,0) = 0 and round(PC.RecebidoNaoPago,0)  = 0 and round(PC.Adiantamento,0) = 0   and round(PC.cedente,0)  = 0  and round(PC.cessionario,0) = 0) " & vbCrLf
        End If

        Sql &= " Order By ReduzidoEmpresa, Empresa, NomeCliente;" & vbCrLf

        '*****************************************************************************************************
        '*********************  Cessao de Credito / Procuracao ***********************************************
        '*****************************************************************************************************

        If String.IsNullOrEmpty(TabelaTemporaria) Then
            Sql &= "Select 'CED'                as tipo," & vbCrLf &
                   "       P.Empresa_id         as Empresa," & vbCrLf &
                   "       P.EndEmpresa_id      as EndEmpresa," & vbCrLf &
                   "	   P.Pedido_id          as Pedido," & vbCrLf &
                   "	   0                    as Procuracao," & vbCrLf &
                   "       P.ClienteCessionario as cliente," & vbCrLf &
                   "	   P.EndCessionario     as EndCliente," & vbCrLf &
                   "	   C.Nome," & vbCrLf &
                   "	   P.QtdeSaldo  as Cedente," & vbCrLf &
                   "       0            as Cessionario" & vbCrLf &
                   "  from #Procuracoes P" & vbCrLf &
                   " Inner join Clientes C" & vbCrLf &
                   "    on C.Cliente_Id  = P.ClienteCessionario" & vbCrLf &
                   "   and C.Endereco_Id = P.EndCessionario" & vbCrLf
            If Cessionario Then
                Sql &= " Union All" & vbCrLf &
                       " select 'CES' as tipo," & vbCrLf &
                       "       P.Empresa_id," & vbCrLf &
                       "       P.EndEmpresa_id," & vbCrLf &
                       "	   P.Pedido_id," & vbCrLf &
                       "	   P.Procuracao_id," & vbCrLf &
                       "       P.ClienteCedente," & vbCrLf &
                       "	   P.EndCedente," & vbCrLf &
                       "	   C.Nome," & vbCrLf &
                       "       0," & vbCrLf &
                       "	   P.QtdeSaldo" & vbCrLf &
                       "  from #Procuracoes P" & vbCrLf &
                       " Inner join Clientes C" & vbCrLf &
                       "   on C.Cliente_Id  = P.ClienteCedente" & vbCrLf &
                       "  and C.Endereco_Id = P.EndCedente" & vbCrLf
            End If
        End If

        Return Sql
    End Function

End Class

<Serializable()> _
Public Class PosicaoDePedido

End Class