Imports System.Web
Imports System.Web.UI
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ExtratoPedido

    Dim Page As System.Web.UI.Page

#Region "Fields"

    Private _Pedido As Pedido
    Dim Banco As New AcessaBanco

#End Region

    Public Sub ExibirExtrato(ByVal pPage As System.Web.UI.Page, ByVal pPedido As Pedido)
        Page = pPage
        _Pedido = pPedido
        ImprimirPedidos()
    End Sub

    Private Sub ImprimirPedidos()
        Try
            Dim param As New Dictionary(Of String, Object)
            GetParamCabecalho(param)

            Dim ds As DataSet = getDataSetPedido()

            If ds Is Nothing OrElse ds.Tables("Pedidos").Rows.Count = 0 Then
                MsgBox(Page, "Nenhum resultado encontrado.")
                Exit Sub
            End If

            ds.Merge(getDataSetNotasFiscais())
            ds.Merge(getDataSetResumoNotasFiscaisPrd())
            ds.Merge(getDataSetResumoNotasFiscais())
            ds.Merge(getDataSetContratos())
            ds.Merge(getDataSetFixacoes())

            ds.Merge(getDataSetResumoFinanceiro())
            ds.Merge(getDataSetFinanceiro())
            ds.Merge(getDataSetRazao())


            ds.Merge(getDataSetAdiantamentos())
            ds.Merge(getDataSetProcuracoes(True))
            ds.Merge(getDataSetProcuracoes(False))
            getDataSetResumoPedido(ds)

            getSumNota(ds, param)

            Funcoes.BindReport(Me.Page, ds, "Cr_ExtratoDePedido_Novo", eExportType.PDF, param)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Function getDataSetResumoPedido(ByRef ds As DataSet) As DataSet
        Try
            ds.Tables.Add("ResumoPedido")

            ds.Tables("ResumoPedido").Columns.Add("Pedido", GetType(System.Int32))
            ds.Tables("ResumoPedido").Columns.Add("QtdeContratada", GetType(System.Int32))
            ds.Tables("ResumoPedido").Columns.Add("QtdeFixada", GetType(System.Int32))
            ds.Tables("ResumoPedido").Columns.Add("SaldoAFixar", GetType(System.Int32))

            ds.Tables("ResumoPedido").Columns.Add("QtdeEntregue", GetType(System.Int32))
            ds.Tables("ResumoPedido").Columns.Add("QtdeDevolvida", GetType(System.Int32))
            ds.Tables("ResumoPedido").Columns.Add("SaldoAEntregar", GetType(System.Int32))

            ds.Tables("ResumoPedido").Columns.Add("QtdeCessionario", GetType(System.Int32))
            ds.Tables("ResumoPedido").Columns.Add("QtdeCedente", GetType(System.Int32))
            ds.Tables("ResumoPedido").Columns.Add("SaldoProcuracao", GetType(System.Int32))

            For Each rowPedido As DataRow In ds.Tables("Pedidos").Rows

                Dim rowResumo As DataRow = ds.Tables("ResumoPedido").NewRow
                rowResumo("Pedido") = rowPedido("Pedido")
                rowResumo("QtdeContratada") = IIf(IsNumeric(ds.Tables("Contratos").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido"))), ds.Tables("Contratos").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido")), 0)
                rowResumo("QtdeFixada") = IIf(IsNumeric(ds.Tables("Fixacoes").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido"))), ds.Tables("Fixacoes").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido")), 0)
                rowResumo("SaldoAFixar") = rowResumo("QtdeContratada") - rowResumo("QtdeFixada")

                rowResumo("QtdeCessionario") = IIf(IsNumeric(ds.Tables("Procuracoes").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido"))), ds.Tables("Procuracoes").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido")), 0)
                rowResumo("QtdeCedente") = IIf(IsNumeric(ds.Tables("ProcuracoesCedente").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido"))), ds.Tables("ProcuracoesCedente").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido")), 0)
                rowResumo("SaldoProcuracao") = rowResumo("QtdeCessionario") - rowResumo("QtdeCedente")

                Dim devolucao As Integer
                Dim entregue As Integer

                For Each rowResumoNotasFiscais As DataRow In ds.Tables("ResumoNotasFiscais").Rows
                    If rowResumoNotasFiscais("Pedido") = rowPedido("Pedido") AndAlso rowResumoNotasFiscais("EntradaSaida") <> rowPedido("EntradaSaida") Then
                        devolucao += rowResumoNotasFiscais("PesoFiscal")
                    End If
                    If rowResumoNotasFiscais("Pedido") = rowPedido("Pedido") AndAlso rowResumoNotasFiscais("EntradaSaida") = rowPedido("EntradaSaida") Then
                        entregue += rowResumoNotasFiscais("PesoFiscal")
                    End If
                Next

                rowResumo("QtdeEntregue") = entregue
                rowResumo("QtdeDevolvida") = devolucao

                If rowPedido("Classe") = "DEPOSITOS" OrElse rowPedido("Classe") = "MUTUO" Then
                    rowResumo("SaldoAEntregar") = rowResumo("QtdeEntregue") - rowResumo("QtdeDevolvida")
                Else
                    rowResumo("SaldoAEntregar") = rowResumo("QtdeContratada") - rowResumo("QtdeEntregue") + rowResumo("QtdeDevolvida")
                End If

                ds.Tables("ResumoPedido").Rows.Add(rowResumo)
            Next

            Return ds
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetPedido() As DataSet
        Try
            Dim strSQL As String = "SELECT P.Empresa_Id as Empresa , P.EndEmpresa_Id as EndEmpresa, P.Pedido_Id as Pedido, so.EntradaSaida, so.Descricao as DescricaoOperacao," & vbCrLf & _
                                   "       P.UnidadeDeNegocio, P.Cliente, P.EndCliente, P.Praca, P.EndPraca, P.PedidoEfetivo, P.Safra, " & vbCrLf & _
                                   "       P.Moeda, P.Indexador, P.Operacao, P.SubOperacao, P.Situacao, P.DataPedido, P.DataEntrega, P.PedidoOrigem, " & vbCrLf & _
                                   "       P.FreteCIFFOB, isnull(P.OrigemDestino,'') AS OrigemDestino, P.Solicitacao, P.UsuarioInclusao, P.UsuarioInclusaoData, isnull(P.UsuarioAlteracao,'') AS UsuarioAlteracao, " & vbCrLf & _
                                   "       P.UsuarioAlteracaoData, isnull(P.UsuarioCancelamento,'') AS UsuarioCancelamento, P.UsuarioCancelamentoData, isnull(P.UsuarioLiberacao,'') AS UsuarioLiberacao, " & vbCrLf & _
                                   "       P.UsuarioLiberacaoData, P.Observacoes, isnull(P.CondicaoPagamento,0) as CondicaoPagamento,  P.BancoCliente, P.AgenciaCliente, " & vbCrLf & _
                                   "       P.DigitoAgenciaCliente, P.ContaCliente, P.DigitoContaCliente, isnull(P.Comercializacao,'') as Comercializacao," & vbCrLf & _
                                   "       isnull(P.Finalidade,1) AS Finalidade, isnull(P.Contrato,'') as Contrato, isnull(Carteira,'') as Carteira, " & vbCrLf & _
                                   "       isnull(P.Taxa,0) as Taxa, isnull(P.VencimentoPedido,P.DataEntrega) as VencimentoPedido, " & vbCrLf & _
                                   "       isnull(LocalEmbarque,'') as LocalEmbarque, isnull(EndLocalEmbarque,0) as EndLocalEmbarque," & vbCrLf & _
                                   "       isnull(MomentoFinanceiro,0) as MomentoFinanceiro, isnull(AgruparFinanceiro,0) as AgruparFinanceiro, " & vbCrLf & _
                                   "       isnull(IndiceFixado, 0) AS IndiceFixado, isnull(EstadoEntrega,'') AS EstadoEntrega, isnull(CidadeEntrega,'') AS CidadeEntrega, " & vbCrLf & _
                                   "       isnull(FiscalAberto,1) as FiscalAberto, isnull(FinanceiroAberto,1) as FinanceiroAberto, " & vbCrLf & _
                                   "       isnull(EmpresaTroca, '') AS EmpresaTroca, isnull(EndEmpresaTroca, 0) AS EndEmpresaTroca, " & vbCrLf & _
                                   "       isnull(PedidoTroca, 0) AS PedidoTroca, isnull(ContaAdiantamentoTroca, '') AS ContaAdiantamentoTroca, " & vbCrLf & _
                                   "       isnull(CondicaoPagamentoEntrega, 0) AS CondicaoPagamentoEntrega, isnull(QuotaEntrega, 0) AS QuotaEntrega, " & vbCrLf & _
                                   "       isnull(PeriodicidadeEntrega, 0) AS PeriodicidadeEntrega, isnull(PedidoBloqueado,0) AS PedidoBloqueado " & vbCrLf & _
                                   "       ,isnull(VersaoPedido,0) as VersaoPedido," & vbCrLf & _
                                   "       isnull(VersaoUsuario,isnull(P.UsuarioAlteracao, P.UsuarioInclusao)) as VersaoUsuario," & vbCrLf & _
                                   "       isnull(VersaoHorarioBloqueio,CAST(CAST(GETDATE() AS DATE) AS DATETIME)) as VersaoHorarioBloqueio, isnull(IndexadorFixo,1) AS IndexadorFixo, " & vbCrLf & _
                                   "       Cli.Cliente_Id + ' - ' +  Cli.Nome + ' - ' + Cli.Cidade + '/' + Cli.Estado as NomeCliente, SO.PrecoFixo, SO.Classe" & vbCrLf & _
                                   "  FROM Pedidos P " & vbCrLf & _
                                   "        INNER JOIN SubOperacoes SO " & vbCrLf & _
                                   "            ON SO.Operacao_Id = P.Operacao " & vbCrLf & _
                                   "            AND SO.SubOperacoes_Id = P.SubOperacao " & vbCrLf & _
                                   "         Inner Join Clientes cli" & vbCrLf & _
                                   "             ON cli.Cliente_Id = P.Cliente" & vbCrLf & _
                                   "             And cli.Endereco_Id = P.EndCliente" & vbCrLf & _
                                   "  WHERE P.UnidadeDeNegocio IS NOT NULL " & vbCrLf & _
                                   "        AND P.PedidoOrigem = 0 " & vbCrLf & _
                                   "        AND P.Situacao = 1" & vbCrLf

            If Not String.IsNullOrWhiteSpace(_Pedido.Codigo) Then
                strSQL &= "AND P.Pedido_Id in (" & _Pedido.Codigo & ")"
            End If

            If Not String.IsNullOrWhiteSpace(_Pedido.CodigoEmpresa) Then
                strSQL &= "AND P.Empresa_Id = '" & _Pedido.CodigoEmpresa & "' " & vbCrLf & _
                          "AND P.EndEmpresa_Id = " & _Pedido.EnderecoEmpresa & vbCrLf
            End If

            Dim strSQLProduto As String = "AND EXISTS (SELECT NULL " & vbCrLf & _
                                                      "FROM PedidoXItemxLancamento PIT " & vbCrLf & _
                                                      "[INNER] " & vbCrLf & _
                                                      "WHERE P.Empresa_Id  = PIT.Empresa_Id " & vbCrLf & _
                                                      "AND P.EndEmpresa_Id = PIT.EndEmpresa_Id " & vbCrLf & _
                                                      "AND P.Pedido_Id     = PIT.Pedido_Id " & vbCrLf & _
                                                      "[WHERE]) " & vbCrLf

            Dim strSQLProdFinal As String = ""

            strSQL &= strSQLProdFinal

            strSQL &= "ORDER BY DataPedido" & vbCrLf

            Dim dsPedidos As DataSet = Banco.ConsultaDataSet(strSQL, "Pedidos")

            Return dsPedidos
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetNotasFiscais() As DataSet
        Try
            Dim strSQL As String = "SELECT NF.Empresa_Id as Empresa, NF.EndEmpresa_Id as EndEmpresa, NF.Cliente_Id as Cliente, NF.EndCliente_Id as EndCliente, NF.Movimento, NFI.Produto_Id as Produto, Produtos.Nome AS NomeProduto," & vbCrLf & _
                     "       NF.Operacao, NF.SubOperacao, SubOperacoes.Classe, SubOperacoes.Devolucao, NFI.CFOP_Id AS CFOP, NULL AS Reduzido, NF.EntradaSaida_Id as EntradaSaida, " & vbCrLf & _
                     "       NF.Serie_Id as Serie, NF.Nota_Id as Nota, Produtos.Agrupar," & vbCrLf & _
                     "       ISNULL(RomaneiosXPesagens.Pesagem_Id, 0) AS NumeroTicket," & vbCrLf & _
                     "       NF.Pedido AS NumeroContrato, " & vbCrLf & _
                     "       case" & vbCrLf & _
                     "         when produtos.ControlarEmbalagem = 'S'  OR produtos.ControlarLote = 'S'" & vbCrLf & _
                     "           then NFI.QuantidadeFisica" & vbCrLf & _
                     "           else ISNULL(Romaneios.PesoBruto, 0)" & vbCrLf & _
                     "       END AS PesoBalanca," & vbCrLf & _
                     "       ISNULL(NXD.PesoLiquido,0) as PesoChegada, " & vbCrLf & _
                     "       ISNULL(sb_Descontos.Umidade, 0) AS Umidade, " & vbCrLf & _
                     "       ISNULL(sb_Descontos.Impureza, 0) AS Impureza," & vbCrLf & _
                     "       ISNULL(sb_Descontos.Avariados, 0) AS Avariados, " & vbCrLf & _
                     "       ISNULL(sb_Descontos.Verde, 0) AS Verde," & vbCrLf & _
                     "       ISNULL(sb_Descontos.PH, 0) AS PH," & vbCrLf & _
                     "       ISNULL(sb_Descontos.GMO, 0) AS GMO, " & vbCrLf & _
                     "       case" & vbCrLf & _
                     "         when produtos.ControlarEmbalagem = 'S'  OR produtos.ControlarLote = 'S'" & vbCrLf & _
                     "           then NFI.QuantidadeFisica" & vbCrLf & _
                     "           else ISNULL(Romaneios.PesoLiquido, 0)" & vbCrLf & _
                     "       end AS PesoLiquido, " & vbCrLf & _
                     "       case when ISNULL(NFI.PesoFiscal,0) = 0 " & vbCrLf & _
                     "          then NFI.QuantidadeFisica " & vbCrLf & _
                     "          else NFI.PesoFiscal " & vbCrLf & _
                     "        end as PesoFiscal, " & vbCrLf & _
                     "       NFI.Unitario," & vbCrLf & _
                     "       NFI.Valor," & vbCrLf & _
                     "       NF.Pedido, ISNULL(NF.Finalidade, 0) AS Finalidade, NFENC.Valor AS Liquido, ISNULL(NFT.Placa,'') AS Placa " & vbCrLf & _
                     "  FROM NotasFiscais AS NF " & vbCrLf & _
                     " INNER JOIN NotasFiscaisXItens AS NFI " & vbCrLf & _
                     "    ON NFI.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                     "   AND NFI.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                     "   AND NFI.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                     "   AND NFI.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                     "   AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                     "   AND NFI.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                     "   AND NFI.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                     " INNER JOIN NotasFiscaisXEncargos AS NFENC " & vbCrLf & _
                     "    ON NFENC.Empresa_Id      = NFI.Empresa_Id " & vbCrLf & _
                     "   AND NFENC.EndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf & _
                     "   AND NFENC.Cliente_Id      = NFI.Cliente_Id " & vbCrLf & _
                     "   AND NFENC.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf & _
                     "   AND NFENC.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
                     "   AND NFENC.Serie_Id        = NFI.Serie_Id " & vbCrLf & _
                     "   AND NFENC.Nota_Id         = NFI.Nota_Id " & vbCrLf & _
                     "   AND NFENC.Produto_Id      = NFI.Produto_Id " & vbCrLf & _
                     "   AND NFENC.CFOP_Id         = NFI.CFOP_Id " & vbCrLf & _
                     "   AND NFENC.Sequencia_Id    = NFI.Sequencia_Id " & vbCrLf & _
                     "   AND NFENC.Encargo_Id      = 'LIQUIDO' " & vbCrLf & _
                     " INNER JOIN Produtos " & vbCrLf & _
                     "    ON Produtos.Produto_Id   = NFI.Produto_Id " & vbCrLf & _
                     " INNER JOIN SubOperacoes " & vbCrLf & _
                     "    ON SubOperacoes.Operacao_Id     = NF.Operacao " & vbCrLf & _
                     "   AND SubOperacoes.SubOperacoes_Id = NF.SubOperacao " & vbCrLf & _
                     " LEFT JOIN NotasXDestinos NXD" & vbCrLf & _
                     "    ON NXD.Empresa_Id      = NFI.Empresa_Id      " & vbCrLf & _
                     "   AND NXD.EndEmpresa_Id   = NFI.EndEmpresa_Id   " & vbCrLf & _
                     "   AND NXD.Cliente_Id      = NFI.Cliente_Id      " & vbCrLf & _
                     "   AND NXD.EndCliente_Id   = NFI.EndCliente_Id   " & vbCrLf & _
                     "   AND NXD.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
                     "   AND NXD.Serie_Id        = NFI.Serie_Id        " & vbCrLf & _
                     "   AND NXD.Nota_Id         = NFI.Nota_Id         " & vbCrLf & _
                     "  -- AND NXD.Produto_Id      = NFI.Produto_Id      " & vbCrLf & _
                     "  LEFT JOIN NotasFiscaisXTransportadores AS NFT " & vbCrLf & _
                     "    ON NFT.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                     "   AND NFT.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                     "   AND NFT.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                     "   AND NFT.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                     "   AND NFT.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                     "   AND NFT.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                     "   AND NFT.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                     "  LEFT JOIN NotasFiscaisXRomaneios AS NFR " & vbCrLf & _
                     "    ON NFR.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                     "   AND NFR.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                     "   AND NFR.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                     "   AND NFR.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                     "   AND NFR.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                     "   AND NFR.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                     "   AND NFR.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                     "  LEFT JOIN Romaneios " & vbCrLf & _
                     "    ON Romaneios.Empresa_Id    = NFR.Empresa_Id " & vbCrLf & _
                     "   AND Romaneios.EndEmpresa_Id = NFR.EndEmpresa_Id " & vbCrLf & _
                     "   AND Romaneios.Romaneio_Id   = NFR.Romaneio_Id " & vbCrLf & _
                     "  LEFT JOIN RomaneiosXPesagens " & vbCrLf & _
                     "    ON RomaneiosXPesagens.Empresa_Id    = Romaneios.Empresa_Id " & vbCrLf & _
                     "   AND RomaneiosXPesagens.EndEmpresa_Id = Romaneios.EndEmpresa_Id " & vbCrLf & _
                     "   AND RomaneiosXPesagens.Romaneio_Id   = Romaneios.Romaneio_Id " & vbCrLf & _
                     "  LEFT JOIN " & vbCrLf & _
                     "	 		  (SELECT Empresa_Id, EndEmpresa_Id, Romaneio_id," & vbCrLf & _
                     "                    SUM(CASE WHEN ANALISE_ID = 1 THEN Desconto ELSE 0 END) AS Umidade, " & vbCrLf & _
                     "	 				  SUM(CASE WHEN ANALISE_ID = 2 THEN Desconto ELSE 0 END) AS Impureza," & vbCrLf & _
                     "                    SUM(CASE WHEN ANALISE_ID = 3 THEN Desconto ELSE 0 END) AS Avariados, " & vbCrLf & _
                     "					  SUM(CASE WHEN ANALISE_ID = 4 THEN Desconto ELSE 0 END) AS Verde," & vbCrLf & _
                     "                    SUM(CASE WHEN ANALISE_ID = 5 THEN Desconto ELSE 0 END) AS PH, " & vbCrLf & _
                     "					  SUM(CASE WHEN ANALISE_ID = 6 THEN Desconto ELSE 0 END) AS GMO " & vbCrLf & _
                     "				 FROM RomaneiosXDescontos " & vbCrLf & _
                     "				GROUP BY Empresa_Id, EndEmpresa_Id, Romaneio_ID" & vbCrLf & _
                     "             ) sb_Descontos " & vbCrLf & _
                     "    ON sb_Descontos.Empresa_Id    = Romaneios.Empresa_Id " & vbCrLf & _
                     "   AND sb_Descontos.EndEmpresa_Id = Romaneios.EndEmpresa_Id " & vbCrLf & _
                     "   AND sb_Descontos.Romaneio_Id   = Romaneios.Romaneio_Id " & vbCrLf & _
                     " WHERE NF.Empresa_Id    ='" & _Pedido.CodigoEmpresa & "'" & vbCrLf & _
                     "   AND NF.EndEmpresa_Id = " & _Pedido.EnderecoEmpresa & vbCrLf

            If Not String.IsNullOrWhiteSpace(_Pedido.Codigo) Then
                strSQL &= "   AND NF.Pedido        in (" & _Pedido.Codigo & ")" & vbCrLf
            End If

            strSQL &= " ORDER BY NF.Movimento "

            Dim dsNotasFiscais As DataSet = Banco.ConsultaDataSet(strSQL, "NotasFiscais")
            Return dsNotasFiscais '#NFD 07
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetResumoNotasFiscaisPrd() As DataSet
        Try
            Dim strSQL As String
            strSQL = "SELECT NFI.Produto_Id as Produto, PRD.Nome, " & vbCrLf & _
                     "       NF.Cliente_Id as Cliente, " & vbCrLf & _
                     "       NF.EndCliente_Id as EndCliente, " & vbCrLf & _
                     "       NF.Pedido, " & vbCrLf & _
                     "       sum(case " & vbCrLf & _
                     "       	   when SB.Devolucao = 'S' " & vbCrLf & _
                     "               then ISNULL(Romaneios.PesoLiquido, 0) * -1 " & vbCrLf & _
                     "               else ISNULL(Romaneios.PesoLiquido, 0) " & vbCrLf & _
                     "             end) AS PesoLiquido, " & vbCrLf & _
                     "       sum(case " & vbCrLf & _
                     "             when SB.Devolucao = 'S' " & vbCrLf & _
                     "               then ISNULL(NFI.PesoFiscal, 0) * -1 " & vbCrLf & _
                     "               else ISNULL(NFI.PesoFiscal, 0) " & vbCrLf & _
                     "             end) AS PesoFiscal, " & vbCrLf & _
                     "       sum(case " & vbCrLf & _
                     "       	   when SB.Devolucao = 'S' " & vbCrLf & _
                     "               then ISNULL(NFI.Valor, 0) * -1 " & vbCrLf & _
                     "               else ISNULL(NFI.Valor, 0) " & vbCrLf & _
                     "           end) AS Valor, " & vbCrLf & _
                     "       sum(case " & vbCrLf & _
                     "       	   when SB.Devolucao = 'S' " & vbCrLf & _
                     "               then ISNULL(NFENC.Valor, 0) * -1 " & vbCrLf & _
                     "               else ISNULL(NFENC.Valor, 0) " & vbCrLf & _
                     "           end) AS Liquido " & vbCrLf & _
                     "  FROM NotasFiscais AS NF " & vbCrLf & _
                     " INNER JOIN NotasFiscaisXItens AS NFI " & vbCrLf & _
                     "    ON NFI.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                     "   AND NFI.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                     "   AND NFI.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                     "   AND NFI.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                     "   AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                     "   AND NFI.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                     "   AND NFI.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                     " INNER JOIN NotasFiscaisXEncargos AS NFENC " & vbCrLf & _
                     "    ON NFENC.Empresa_Id      = NFI.Empresa_Id " & vbCrLf & _
                     "   AND NFENC.EndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf & _
                     "   AND NFENC.Cliente_Id      = NFI.Cliente_Id " & vbCrLf & _
                     "   AND NFENC.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf & _
                     "   AND NFENC.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
                     "   AND NFENC.Serie_Id        = NFI.Serie_Id " & vbCrLf & _
                     "   AND NFENC.Nota_Id         = NFI.Nota_Id " & vbCrLf & _
                     "   AND NFENC.Produto_Id      = NFI.Produto_Id " & vbCrLf & _
                     "   AND NFENC.CFOP_Id         = NFI.CFOP_Id " & vbCrLf & _
                     "   AND NFENC.Sequencia_Id    = NFI.Sequencia_Id " & vbCrLf & _
                     "   AND NFENC.Encargo_Id      = 'LIQUIDO' " & vbCrLf & _
                     "  LEFT JOIN NotasFiscaisXRomaneios AS NFR " & vbCrLf & _
                     "    ON NFR.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                     "   AND NFR.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                     "   AND NFR.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                     "   AND NFR.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                     "   AND NFR.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                     "   AND NFR.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                     "   AND NFR.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                     " INNER JOIN Produtos AS PRD " & vbCrLf & _
                     "    ON PRD.Produto_Id      = NFI.Produto_Id " & vbCrLf & _
                     " INNER JOIN SubOperacoes AS SB " & vbCrLf & _
                     "    ON SB.Operacao_Id      = NF.Operacao " & vbCrLf & _
                     "   AND SB.SubOperacoes_Id  = NF.SubOperacao " & vbCrLf & _
                     "  LEFT JOIN Romaneios " & vbCrLf & _
                     "    ON Romaneios.Empresa_Id    = NFR.Empresa_Id " & vbCrLf & _
                     "   AND Romaneios.EndEmpresa_Id = NFR.EndEmpresa_Id " & vbCrLf & _
                     "   AND Romaneios.Romaneio_Id   = NFR.Romaneio_Id " & vbCrLf & _
                     " WHERE (NF.Empresa_Id = '" & _Pedido.CodigoEmpresa & "') " & vbCrLf & _
                     "   AND (NF.EndEmpresa_Id = " & _Pedido.EnderecoEmpresa & ") " & vbCrLf

            If Not String.IsNullOrWhiteSpace(_Pedido.Codigo) Then
                strSQL &= "   AND NF.Pedido IN (" & _Pedido.Codigo & ") " & vbCrLf
            End If

            strSQL &= "   group by NF.Cliente_Id, NF.EndCliente_Id, NFI.Produto_Id, PRD.Nome, NF.Pedido  " & vbCrLf & _
                      "   ORDER BY 1"

            Dim dsResumoNotasFiscaisPrd As DataSet = Banco.ConsultaDataSet(strSQL, "ResumoNotasFiscaisPrd")
            Return dsResumoNotasFiscaisPrd
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetResumoNotasFiscais() As DataSet
        Try
            Dim strSQL As String = "SELECT prd.Agrupar, NF.Operacao, NF.SubOperacao, NF.EntradaSaida_Id as EntradaSaida, SO.Descricao, NF.Cliente_Id as Cliente, NF.EndCliente_Id as EndCliente, NF.Pedido," & vbCrLf & _
                                   "       sum(ISNULL(Romaneios.PesoBruto, 0)) AS PesoBalanca, sum(ISNULL(sb_Descontos.Umidade, 0)) AS Umidade, " & vbCrLf & _
                                   "       sum(ISNULL(sb_Descontos.Impureza, 0)) AS Impureza, sum(ISNULL(sb_Descontos.Avariados, 0)) AS Avariados, " & vbCrLf & _
                                   "       sum(ISNULL(sb_Descontos.Verde, 0)) AS Verde, sum(ISNULL(sb_Descontos.PH, 0)) AS PH, sum(ISNULL(sb_Descontos.GMO, 0)) AS GMO, " & vbCrLf & _
                                   "       sum(ISNULL(Romaneios.PesoLiquido, 0)) AS PesoLiquido, sum(NFI.PesoFiscal) AS PesoFiscal, sum(NFI.Unitario) AS Unitario, sum(NFI.Valor) AS Valor, sum(NFENC.Valor) AS Liquido " & vbCrLf & _
                                   "  FROM NotasFiscais AS NF " & vbCrLf & _
                                   " INNER JOIN NotasFiscaisXItens AS NFI " & vbCrLf & _
                                   "    ON NFI.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                                   "   AND NFI.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                                   "   AND NFI.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                                   "   AND NFI.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                                   "   AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                                   "   AND NFI.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                                   "   AND NFI.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                                   " INNER JOIN NotasFiscaisXEncargos AS NFENC " & vbCrLf & _
                                   "    ON NFENC.Empresa_Id      = NFI.Empresa_Id " & vbCrLf & _
                                   "   AND NFENC.EndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf & _
                                   "   AND NFENC.Cliente_Id      = NFI.Cliente_Id " & vbCrLf & _
                                   "   AND NFENC.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf & _
                                   "   AND NFENC.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
                                   "   AND NFENC.Serie_Id        = NFI.Serie_Id " & vbCrLf & _
                                   "   AND NFENC.Nota_Id         = NFI.Nota_Id " & vbCrLf & _
                                   "   AND NFENC.Produto_Id      = NFI.Produto_Id " & vbCrLf & _
                                   "   AND NFENC.CFOP_Id         = NFI.CFOP_Id " & vbCrLf & _
                                   "   AND NFENC.Sequencia_Id    = NFI.Sequencia_Id " & vbCrLf & _
                                   "   AND NFENC.Encargo_Id      = 'LIQUIDO' " & vbCrLf & _
                                   " INNER JOIN PRODUTOS PRD" & vbCrLf & _
                                   "     ON PRD.PRODUTO_ID = NFI.PRODUTO_ID" & vbCrLf & _
                                   "  LEFT OUTER JOIN NotasFiscaisXRomaneios AS NFR " & vbCrLf & _
                                   "    ON NFR.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                                   "   AND NFR.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                                   "   AND NFR.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                                   "   AND NFR.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                                   "   AND NFR.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                                   "   AND NFR.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                                   "   AND NFR.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                                   "  LEFT OUTER JOIN Romaneios " & vbCrLf & _
                                   "    ON Romaneios.Empresa_Id    = NFR.Empresa_Id " & vbCrLf & _
                                   "   AND Romaneios.EndEmpresa_Id = NFR.EndEmpresa_Id " & vbCrLf & _
                                   "   AND Romaneios.Romaneio_Id   = NFR.Romaneio_Id " & vbCrLf & _
                                   "  LEFT OUTER JOIN RomaneiosXPesagens " & vbCrLf & _
                                   "    ON RomaneiosXPesagens.Empresa_Id    = Romaneios.Empresa_Id " & vbCrLf & _
                                   "   AND RomaneiosXPesagens.EndEmpresa_Id = Romaneios.EndEmpresa_Id " & vbCrLf & _
                                   "   AND RomaneiosXPesagens.Romaneio_Id   = Romaneios.Romaneio_Id " & vbCrLf & _
                                   "  LEFT JOIN (SELECT Empresa_Id, EndEmpresa_Id, Romaneio_id, SUM(CASE WHEN ANALISE_ID = 1 THEN Desconto ELSE 0 END) AS Umidade, " & vbCrLf & _
                                   "	 				SUM(CASE WHEN ANALISE_ID = 2 THEN Desconto ELSE 0 END) AS Impureza, SUM(CASE WHEN ANALISE_ID = 3 THEN Desconto ELSE 0 END) AS Avariados, " & vbCrLf & _
                                   "					SUM(CASE WHEN ANALISE_ID = 4 THEN Desconto ELSE 0 END) AS Verde, SUM(CASE WHEN ANALISE_ID = 5 THEN Desconto ELSE 0 END) AS PH, " & vbCrLf & _
                                   "					SUM(CASE WHEN ANALISE_ID = 6 THEN Desconto ELSE 0 END) AS GMO " & vbCrLf & _
                                   "			   FROM RomaneiosXDescontos " & vbCrLf & _
                                   "			  GROUP BY Empresa_Id, EndEmpresa_Id, Romaneio_ID" & vbCrLf & _
                                   "             ) sb_Descontos " & vbCrLf & _
                                   "    ON sb_Descontos.Empresa_Id    = Romaneios.Empresa_Id " & vbCrLf & _
                                   "   AND sb_Descontos.EndEmpresa_Id = Romaneios.EndEmpresa_Id " & vbCrLf & _
                                   "   AND sb_Descontos.Romaneio_Id   = Romaneios.Romaneio_Id " & vbCrLf & _
                                   " INNER JOIN SubOperacoes AS SO " & vbCrLf & _
                                   "    ON SO.Operacao_Id     = NF.Operacao " & vbCrLf & _
                                   "   AND SO.SubOperacoes_Id = NF.SubOperacao " & vbCrLf & _
                                   " WHERE (NF.Empresa_Id = '" & _Pedido.CodigoEmpresa & "') " & vbCrLf & _
                                   "   AND (NF.EndEmpresa_Id = " & _Pedido.EnderecoEmpresa & ") " & vbCrLf

            If Not String.IsNullOrWhiteSpace(_Pedido.Codigo) Then
                strSQL &= "   AND NF.Pedido IN (" & _Pedido.Codigo & ") " & vbCrLf
            End If

            strSQL &= "   group by prd.Agrupar, NF.Operacao, NF.SubOperacao, NF.Cliente_Id, NF.EndCliente_Id, NF.EntradaSaida_Id, SO.Descricao, NF.Pedido " & vbCrLf & _
                      "   ORDER BY 1"

            Dim dsResumoNotasFiscais As DataSet = Banco.ConsultaDataSet(strSQL, "ResumoNotasFiscais")
            Return dsResumoNotasFiscais
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetContratos() As DataSet
        Try
            Dim strSQL As String = "SELECT P.DataPedido, PIT.TipoDeLancamento, PIT.Pedido_Id as Pedido, PIT.Quantidade, " & vbCrLf & _
                           "M.Descricao AS NomeMoeda, PIT.UnitarioOficial, PIT.UnitarioMoeda, " & vbCrLf & _
                           "PIT.TotalOficial, PIT.TotalMoeda, PIT.Produto_Id + '-' + Produtos.Nome AS NomeProduto, PIT.Produto_Id as Produto, " & vbCrLf & _
                           "isnull(P.IndiceFixado,0) AS IndiceFixado, P.Cliente, P.EndCliente " & vbCrLf & _
                           "FROM PedidoXItemxLancamento PIT " & vbCrLf & _
                           "INNER JOIN Pedidos P " & vbCrLf & _
                           "ON P.Empresa_Id = PIT.Empresa_Id " & vbCrLf & _
                           "AND P.EndEmpresa_Id = PIT.EndEmpresa_Id " & vbCrLf & _
                           "AND P.Pedido_Id = PIT.Pedido_Id " & vbCrLf & _
                           "INNER JOIN Moedas M " & vbCrLf & _
                           "ON M.Moeda_Id = P.Moeda INNER JOIN " & vbCrLf & _
                           "Produtos ON PIT.Produto_Id = Produtos.Produto_Id " & vbCrLf & _
                           "WHERE P.Empresa_Id = '" & _Pedido.CodigoEmpresa & "' " & vbCrLf & _
                           "AND P.EndEmpresa_Id = " & _Pedido.EnderecoEmpresa & " " & vbCrLf

            If Not String.IsNullOrWhiteSpace(_Pedido.Codigo) Then
                strSQL &= "   AND P.Pedido_Id IN (" & _Pedido.Codigo & ") " & vbCrLf
            End If

            strSQL &= "ORDER BY PIT.PedidoItem_Id " & vbCrLf

            Dim dsContratos As DataSet = Banco.ConsultaDataSet(strSQL, "Contratos")
            Return dsContratos
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetFixacoes() As DataSet
        Dim strSQL As String
        strSQL = "SELECT f.Pedido_id as Pedido, f.Fixacao_Id as Fixacao, f.Procuracao, f.Operacao, f.SubOperacao, f.Documento, f.Movimento, f.Quantidade," & vbCrLf & _
                 "       f.UnitarioOficial, f.UnitarioMoeda, f.TotalOficial, f.TotalMoeda, f.CondicaoPagamento, f.IndiceFixado," & vbCrLf & _
                 "       f.Produto_Id + '-' + p.Descricao as Produto, pxi.TipoDeLancamento, m.Descricao as Moeda, FE.ValorOficial AS Liquido, FE.ValorMoeda AS LiquidoMoeda, " & vbCrLf & _
                 "       Pd.Cliente, Pd.EndCliente" & vbCrLf & _
                 "  FROM VW_PedidosXItensXFixacoes f" & vbCrLf & _
                 " Inner Join Produtos p" & vbCrLf & _
                 "    On p.Produto_Id = f.Produto_Id" & vbCrLf & _
                 " Inner Join Pedidos Pd" & vbCrLf & _
                 "    on Pd.empresa_Id    = f.Empresa_Id" & vbCrLf & _
                 " 	 and pd.EndEmpresa_Id = f.EndEmpresa_Id" & vbCrLf & _
                 " 	 and pd.Pedido_Id     = f.Pedido_Id" & vbCrLf & _
                 " Inner JOin PedidoXItemxLancamento pxi" & vbCrLf & _
                 " 	  On pxi.Empresa_Id    = f.Empresa_Id" & vbCrLf & _
                 " 	 and pxi.EndEmpresa_Id = f.EndEmpresa_Id" & vbCrLf & _
                 " 	 and pxi.Pedido_Id     = f.Pedido_Id" & vbCrLf & _
                 "   and pxi.Produto_Id    = f.Produto_Id" & vbCrLf & _
                 " Inner Join VW_PedidosXItensXFixacoesXEncargos fe" & vbCrLf & _
                 " 	  on fe.Empresa_Id    = f.Empresa_Id" & vbCrLf & _
                 " 	 and fe.EndEmpresa_Id = f.EndEmpresa_Id" & vbCrLf & _
                 " 	 and fe.Pedido_Id     = f.Pedido_Id" & vbCrLf & _
                 "   and fe.Produto_Id    = f.Produto_Id" & vbCrLf & _
                 " 	 and fe.Fixacao_Id    = f.Fixacao_Id" & vbCrLf & _
                 " INNER JOIN Moedas M" & vbCrLf & _
                 " 	  ON M.Moeda_Id = Pd.Moeda" & vbCrLf & _
                 " inner join SubOperacoes SO" & vbCrLf & _
                 " 	  ON SO.Operacao_Id     = f.Operacao" & vbCrLf & _
                 " 	 and so.SubOperacoes_Id = f.SubOperacao" & vbCrLf & _
                 " 	 AND SO.CLASSE IN('" & eClassesOperacoes.AFIXAR.ToString & "','" & eClassesOperacoes.COMPLEMENTACOES.ToString & "') " & vbCrLf & _
                 " WHERE fe.Encargo_Id   = 'LIQUIDO'" & vbCrLf & _
                 "   AND f.Empresa_Id    = '" & _Pedido.CodigoEmpresa & "'" & vbCrLf & _
                 "   AND f.EndEmpresa_Id = " & _Pedido.EnderecoEmpresa & vbCrLf

        If Not String.IsNullOrWhiteSpace(_Pedido.Codigo) Then
            strSQL &= "AND f.Pedido_Id IN (" & _Pedido.Codigo & ") " & vbCrLf
        End If

        strSQL &= "    ORDER BY f.Pedido_id, f.Fixacao_Id                                                                              " & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "Fixacoes")
        Return ds
    End Function

    Private Function getDataSetFinanceiro() As DataSet
        Try
            Dim strSQL As String = ""


            strSQL = "SELECT T.Pedido, T.Movimento, T.Reprogramacao AS Vencimento, " & vbCrLf & _
                    "		isnull(TC_PROD.ValorOficial,0) AS ValorDoDocumento, T.CliFor as Cliente, T.EnderecoCliFor as EndCliente, " & vbCrLf & _
                    "		case " & vbCrLf & _
                    "			when isnull(TC_PROD.ValorOficial,0) > isnull(TC_LIQ.ValorOficial ,0) " & vbCrLf & _
                    "			  then isnull(TC_PROD.ValorOficial,0) - isnull(TC_LIQ.ValorOficial ,0) " & vbCrLf & _
                    "			  else 0 " & vbCrLf & _
                    "			end AS Descontos, " & vbCrLf & _
                    "       convert(decimal(18,2),0) as Deducoes, " & vbCrLf & _
                    "		case " & vbCrLf & _
                    "			when isnull(TC_LIQ.ValorOficial,0) > isnull(TC_PROD.ValorOficial ,0) " & vbCrLf & _
                    "			  then isnull(TC_LIQ.ValorOficial,0) - isnull(TC_PROD.ValorOficial ,0) " & vbCrLf & _
                    "			  else 0 " & vbCrLf & _
                    "			end AS Juros, " & vbCrLf & _
                    "       convert(decimal(18,2),0) as Acrescimos, " & vbCrLf & _
                    "       case " & vbCrLf & _
                    "           when T.Provisao = 1 " & vbCrLf & _
                    "			 then isnull(TC_LIQ.ValorOficial ,0) " & vbCrLf & _
                    "           else " & vbCrLf & _
                    "             case when M.Classificacao <> 'O' " & vbCrLf & _
                    "                then convert(numeric(18,2), isnull(TC_PROD.ValorMoeda,0) * Cotacoes.indice) " & vbCrLf & _
                    "             else isnull(TC_PROD.ValorOficial,0) " & vbCrLf & _
                    "           end " & vbCrLf & _
                    "        end ValorLiquido, " & vbCrLf & _
                    "		isnull(TC_PROD.ValorMoeda,0) AS MoedaValorDoDocumento, " & vbCrLf & _
                    "		case " & vbCrLf & _
                    "			when isnull(TC_PROD.ValorMoeda,0) > isnull(TC_LIQ.ValorMoeda ,0) " & vbCrLf & _
                    "			  then isnull(TC_PROD.ValorMoeda,0) - isnull(TC_LIQ.ValorMoeda ,0) " & vbCrLf & _
                    "			  else 0 " & vbCrLf & _
                    "			end AS MoedaDescontos, " & vbCrLf & _
                    "       convert(decimal(18,2),0) as MoedaDeducoes, " & vbCrLf & _
                    "		case " & vbCrLf & _
                    "			when isnull(TC_LIQ.ValorMoeda,0) > isnull(TC_PROD.ValorMoeda ,0) " & vbCrLf & _
                    "			  then isnull(TC_LIQ.ValorMoeda,0) - isnull(TC_PROD.ValorMoeda ,0) " & vbCrLf & _
                    "			  else 0 " & vbCrLf & _
                    "			end AS MoedaJuros, " & vbCrLf & _
                    "       convert(decimal(18,2),0) as MoedaAcrescimos, " & vbCrLf & _
                    "        case " & vbCrLf & _
                    "           when T.Provisao = 1 " & vbCrLf & _
                    "			 then isnull(TC_LIQ.ValorMoeda,0) " & vbCrLf & _
                    "           else " & vbCrLf & _
                    "             case when M.Classificacao = 'O' " & vbCrLf & _
                    "                then convert(numeric(18,2),isnull(TC_PROD.ValorOficial,0) / Cotacoes.indice) " & vbCrLf & _
                    "             else isnull(TC_PROD.ValorMoeda,0) " & vbCrLf & _
                    "           end " & vbCrLf & _
                    "        end MoedaValorLiquido, " & vbCrLf & _
                    "       CASE WHEN (SELECT Lancamento FROM vw_AdiantamentosXBaixas WHERE TituloBaixa = T.Titulo_Id) = 'B' " & vbCrLf & _
                    "           THEN 5 " & vbCrLf & _
                    "           Else T.Provisao " & vbCrLf & _
                    "       END AS Provisao, " & vbCrLf & _
                    "		T.Titulo_Id AS Registro, " & vbCrLf & _
                    "		case " & vbCrLf & _
                    "			when Pla.Adiantamento = 1 " & vbCrLf & _
                    "				then " & vbCrLf & _
                    "					case " & vbCrLf & _
                    "						when T.RecPag = 'P' " & vbCrLf & _
                    "							then " & vbCrLf & _
                    "								case " & vbCrLf & _
                    "									when LEFT(Pla.Conta_Id,1) = '1' " & vbCrLf & _
                    "										then 'A' " & vbCrLf & _
                    "										else 'RA' " & vbCrLf & _
                    "									end " & vbCrLf & _
                    "							else " & vbCrLf & _
                    "								case " & vbCrLf & _
                    "									when LEFT(Pla.Conta_Id,1) = '2' " & vbCrLf & _
                    "										then 'RA' " & vbCrLf & _
                    "										else 'R' " & vbCrLf & _
                    "									end	" & vbCrLf & _
                    "						end	" & vbCrLf & _
                    "				else T.RecPag " & vbCrLf & _
                    "			end AS Tipo, " & vbCrLf & _
                    "T.Moeda  " & vbCrLf & _
                    "FROM Titulos T " & vbCrLf & _
                    "  Left Join TitulosxContaContabil TC_LIQ " & vbCrLf & _
                    "    on TC_LIQ.Titulo_Id = T.Titulo_Id " & vbCrLf & _
                    "   and TC_LIQ.Conta_Id  = T.ContaContabilRecPag " & vbCrLf & _
                    "   and TC_LIQ.DC_Id     = case " & vbCrLf & _
                    "                            when T.RecPag in ('P','C') " & vbCrLf & _
                    "                              then 'C' " & vbCrLf & _
                    "                              else 'D' " & vbCrLf & _
                    "                          end " & vbCrLf & _
                    "  Left Join TitulosxContaContabil TC_PROD " & vbCrLf & _
                    "    on TC_PROD.Titulo_Id = T.Titulo_Id " & vbCrLf & _
                    "   and TC_PROD.Conta_Id  = T.ContaContabilCliFor " & vbCrLf & _
                    "   and TC_PROD.DC_Id     = case " & vbCrLf & _
                    "							 when T.RecPag in ('P','C') " & vbCrLf & _
                    "							   then 'D' " & vbCrLf & _
                    "							   else 'C' " & vbCrLf & _
                    "						   end " & vbCrLf & _
                    "   Inner Join cotacoes  " & vbCrLf & _
                    "       on Cotacoes.Data_id      = T.Reprogramacao " & vbCrLf & _
                    "      and Cotacoes.Indexador_Id = T.Indexador  " & vbCrLf & _
                    "   Inner Join Moedas M " & vbCrLf & _
                    "     ON M.Moeda_Id = T.Moeda  " & vbCrLf & _
                    "   Inner Join PlanoDeContas Pla " & vbCrLf & _
                    "     ON Pla.Conta_Id = T.ContaContabilCliFor  " & vbCrLf & _
                    "WHERE T.SITUACAO = 1" & vbCrLf & _
                    "   AND T.Empresa = '" & _Pedido.CodigoEmpresa & "' " & vbCrLf & _
                    "   AND T.EndEmpresa =  " & _Pedido.EnderecoEmpresa & vbCrLf

            If Not String.IsNullOrWhiteSpace(_Pedido.Codigo) Then
                strSQL &= "   AND T.Pedido IN (" & _Pedido.Codigo & ") " & vbCrLf
            End If

            Dim dsFinanceiro As DataSet = Banco.ConsultaDataSet(strSQL, "Financeiro")
            Return dsFinanceiro
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetResumoFinanceiro() As DataSet
        Try
            Dim sql As String = "    SELECT UnidadeDeNegocio,                                 " & vbCrLf & _
                    "       Empresa,                                                          " & vbCrLf & _
                    "       EndEmpresa,                                                       " & vbCrLf & _
                    "       Pedido,                                              " & vbCrLf & _
                    "       Cliente,                                                          " & vbCrLf & _
                    "       EndCliente,                                                       " & vbCrLf & _
                    "       Cifrao,                                                           " & vbCrLf & _
                    "       Troca,                                                            " & vbCrLf & _
                    "       ContaContabilProduto,                                             " & vbCrLf & _
                    "       isnull(ContaContabilAdiantamento,'') as ContaContabilAdiantamento," & vbCrLf & _
                    "       ValorPedido,                                                      " & vbCrLf & _
                    "       ValorTitulosEmPrevisao,                                           " & vbCrLf & _
                    "       ValorTitulosEmProvisao,                                           " & vbCrLf & _
                    "       ValorTitulosBaixado,                                              " & vbCrLf & _
                    "       ValorTitulosCompensado,                                           " & vbCrLf & _
                    "       ValorAdiantamentoOriginal,                                        " & vbCrLf & _
                    "       ValorAdiantamento,                                                " & vbCrLf & _
                    "       ValorAdiantamentoCompensado,                                      " & vbCrLf & _
                    "       ValorAdiantamentoPagoDireto,                                      " & vbCrLf & _
                    "       SaldoAdiantamento,                                                " & vbCrLf & _
                    "       Case When Troca = 0                                               " & vbCrLf & _
                    "			Then 0                                                        " & vbCrLf & _
                    "			else	Case when ISNULL(ValorAdiantamento,0) > 0             " & vbCrLf & _
                    "						Then ISNULL(ValorPedido,0) - ISNULL(ValorTitulosEmProvisao,0) - ISNULL(ValorTitulosBaixado,0) - ISNULL(SaldoAdiantamento,0)  " & vbCrLf & _
                    "						Else 0                                            " & vbCrLf & _
                    "                End                                                      " & vbCrLf & _
                    "	   End as ValorAdiantamentoAmortizado,                                " & vbCrLf & _
                    "	   ISNULL(ValorTitulosEmPrevisao,0) + ISNULL(ValorTitulosEmProvisao,0) + ISNULL(SaldoAdiantamento,0) as SaldoBaixaPedido" & vbCrLf & _
                    "                FROM ResumoFinanceiro                                    " & vbCrLf & _
                    " Where Empresa    ='" & _Pedido.CodigoEmpresa & "'      " & vbCrLf & _
                    "   And EndEmpresa = '" & _Pedido.EnderecoEmpresa & "'     " & vbCrLf

            If Not String.IsNullOrWhiteSpace(_Pedido.Codigo) Then
                sql &= "   And Pedido IN (" & _Pedido.Codigo & ") " & vbCrLf
            End If

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ResumoFinanceiro")

            Return ds

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetRazao() As DataSet
        Try
            Dim strSQL As String = ""

            strSQL = "   SELECT Razao.Pedido,                                         " & vbCrLf & _
                     "          Razao.Movimento_Id as Movimento,                      " & vbCrLf & _
                     "          Razao.Lote_Id as Lote, 	                             " & vbCrLf & _
                     "   	   isnull(Razao.Produto,'') as Produto,                  " & vbCrLf & _
                     "   	   Case isnull(PC.Conta_ID,'')                           " & vbCrLf & _
                     "   	     WHEN ''                                             " & vbCrLf & _
                     "   		   THEN 'CONTA NAO CADASTRADA, Verifique erro.'      " & vbCrLf & _
                     "   		   ELSE Razao.Conta_ID                               " & vbCrLf & _
                     "          END AS Conta,                                         " & vbCrLf & _
                     "   	   PC.Titulo AS NomeConta,                               " & vbCrLf & _
                     "   	   Razao.Historico,                                      " & vbCrLf & _
                     "   	   Case isnull(PC.Conta_ID,'')                           " & vbCrLf & _
                     "   	     WHEN ''                                             " & vbCrLf & _
                     "   		   THEN 0                                            " & vbCrLf & _
                     "   		   ELSE Razao.DebitoOficial                          " & vbCrLf & _
                     "   	   END AS DebitoOficial,                                 " & vbCrLf & _
                     "   	   Case isnull(PC.Conta_ID,'')                           " & vbCrLf & _
                     "   	     WHEN ''                                             " & vbCrLf & _
                     "   		   THEN 0                                            " & vbCrLf & _
                     "   		   ELSE Razao.CreditoOficial                         " & vbCrLf & _
                     "   	   END AS CreditoOficial, 	                             " & vbCrLf & _
                     "   	   Case isnull(PC.Conta_ID,'')                           " & vbCrLf & _
                     "   	     WHEN ''                                             " & vbCrLf & _
                     "   		   THEN 0                                            " & vbCrLf & _
                     "   		   ELSE Razao.DebitoMoeda                            " & vbCrLf & _
                     "   	   END AS DebitoMoeda,                                   " & vbCrLf & _
                     "   	   Case isnull(PC.Conta_ID,'')                           " & vbCrLf & _
                     "   	     WHEN ''                                             " & vbCrLf & _
                     "   		   THEN 0                                            " & vbCrLf & _
                     "   		   ELSE Razao.CreditoMoeda                           " & vbCrLf & _
                     "   	   END AS CreditoMoeda,                                  " & vbCrLf & _
                     "   	   Razao.Cliente_Id as Cliente,                          " & vbCrLf & _
                     "   	   Razao.EndCliente_Id as EndCliente                     " & vbCrLf & _
                     "                       FROM Razao                              " & vbCrLf & _
                     "     LEFT JOIN PlanoDeContas PC                                 " & vbCrLf & _
                     "       ON Razao.Conta_Id = PC.Conta_Id                          " & vbCrLf & _
                     "  Where   Razao.Empresa_Id = '" & _Pedido.CodigoEmpresa & "'" & vbCrLf & _
                     "      And Razao.EndEmpresa_Id = " & _Pedido.EnderecoEmpresa & vbCrLf

            If Not String.IsNullOrWhiteSpace(_Pedido.Codigo) Then
                strSQL &= "And Razao.Pedido IN (" & _Pedido.Codigo & ") " & vbCrLf
            End If

            strSQL &= "     AND Razao.Lote_Id not in (9, 10, 11, 21, 70) " & vbCrLf & _
                      " order by razao.Lote_Id, razao.Movimento_Id " & vbCrLf

            Dim dsRazao As DataSet = Banco.ConsultaDataSet(strSQL, "Razao")

            Return dsRazao
            'Dim Saldo As Decimal = 0
            'Dim SaldoMoeda As Decimal = 0
            'For Each dr As DataRow In dsRazao.Tables(0).Rows
            '    Saldo += dr("Debito") - dr("Credito")
            '    SaldoMoeda += dr("DebitoMoeda") - dr("CreditoMoeda")

            '    If dr("Saldo") = "A" Then
            '        dr("Debito") = 0
            '        dr("Credito") = 0
            '    End If
            '    If dr("SaldoMoeda") = "A" Then
            '        dr("DebitoMoeda") = 0
            '        dr("CreditoMoeda") = 0
            '    End If

            '    If Saldo > 0 Then
            '        dr("Saldo") = Saldo.ToString("N2") & "-D"
            '    End If
            '    If SaldoMoeda > 0 Then
            '        dr("SaldoMoeda") = SaldoMoeda.ToString("N2") & "-D"
            '    End If

            '    If Saldo < 0 Then
            '        dr("Saldo") = (Saldo * -1).ToString("N2") & "-C"
            '    End If
            '    If SaldoMoeda < 0 Then
            '        dr("SaldoMoeda") = (SaldoMoeda * -1).ToString("N2") & "-C"
            '    End If

            '    If Saldo = 0 Then
            '        dr("Saldo") = Saldo.ToString("N2") & "DC"
            '    End If
            '    If SaldoMoeda = 0 Then
            '        dr("SaldoMoeda") = SaldoMoeda.ToString("N2") & "DC"
            '    End If
            'Next

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetAdiantamentos() As DataSet
        Try
            Dim strSQL As String = ""



            strSQL = "SELECT T.Pedido," & vbCrLf & _
                     "       T.Movimento," & vbCrLf & _
                     "       T.Vencimento," & vbCrLf & _
                     "       T.Titulo_id as Titulo," & vbCrLf & _
                     "       m.Descricao as Moeda," & vbCrLf & _
                     "       T.Indexador," & vbCrLf & _
                     "       T.indice," & vbCrLf & _
                     "       TxC.ValorOficial," & vbCrLf & _
                     "       TxC.ValorMoeda," & vbCrLf & _
                     "       isnull(AB.BaixaOficial,0) as Baixas," & vbCrLf & _
                     "       isnull(AB.BaixaMoeda,0)   as BaixasMoeda," & vbCrLf & _
                     "       A.Taxa," & vbCrLf & _
                     "       T.DataBaixa," & vbCrLf & _
                     "       Case" & vbCrLf & _
                     "         When T.DataBaixa > T.Vencimento" & vbCrLf & _
                     "			 Then  txc.ValorOficial * 8 / 100 * DATEDIFF(day,T.Vencimento, T.DataBaixa)" & vbCrLf & _
                     "			 Else 0" & vbCrLf & _
                     "	     End as Juros," & vbCrLf & _
                     "	     Case" & vbCrLf & _
                     "         When T.DataBaixa > T.Vencimento" & vbCrLf & _
                     "			 Then  txc.ValorMoeda * 8 / 100 * DATEDIFF(day,T.Vencimento, T.DataBaixa)" & vbCrLf & _
                     "			 Else 0" & vbCrLf & _
                     "	     End as JurosMoeda," & vbCrLf & _
                     "       T.CliFor as Cliente," & vbCrLf & _
                     "       T.EnderecoCliFor as EndCliente" & vbCrLf & _
                     "  FROM Titulos T" & vbCrLf & _
                     " inner join vw_Adiantamento A" & vbCrLf & _
                     "    on T.Titulo_id     = A.Titulo_id" & vbCrLf & _
                     " inner join TitulosxContacontabil TxC" & vbCrLf & _
                     "    on T.Titulo_id           = TxC.Titulo_id" & vbCrLf & _
                     "   and T.ContaContabilRecPag = TxC.Conta_id" & vbCrLf & _
                     " inner join Moedas M" & vbCrLf & _
                     "	on m.Moeda_Id = t.Moeda" & vbCrLf & _
                     "  LEFT JOIN (SELECT Titulo_Id," & vbCrLf & _
                     "                    SUM(ValorOficial) AS BaixaOficial," & vbCrLf & _
                     "                    SUM(ValorMoeda)   AS BaixaMoeda" & vbCrLf & _
                     "               FROM vw_AdiantamentosXBaixas" & vbCrLf & _
                     "              Where Lancamento = 'B'" & vbCrLf & _
                     "              GROUP BY Titulo_Id" & vbCrLf & _
                     "             ) AB" & vbCrLf & _
                     "    ON T.Titulo_Id = AB.Titulo_Id" & vbCrLf & _
                     " WHERE T.Empresa     = '" & _Pedido.CodigoEmpresa & "'    " & vbCrLf & _
                     "   AND T.EndEmpresa  = " & _Pedido.EnderecoEmpresa & " " & vbCrLf

            If Not String.IsNullOrWhiteSpace(_Pedido.Codigo) Then
                strSQL &= "   AND T.Pedido IN (" & _Pedido.Codigo & ") " & vbCrLf
            End If

            strSQL &= "   ORDER BY T.Movimento" & vbCrLf

            Return Banco.ConsultaDataSet(strSQL, "Adiantamentos")

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getDataSetProcuracoes(ByVal Cedente As Boolean) As DataSet
        Dim strTipo As String = IIf(Cedente, "Cedente", "Cessionario")
        Dim strOutro As String = IIf(Not Cedente, "Cedente", "Cessionario")

        Dim strSQL As String
        strSQL = "SELECT P.Pedido" & strTipo & " as Pedido, P.Procuracao_ID AS Procuracao, P.Documento, P." & strOutro & " + ' - '  + C.Nome AS Cliente, P.Quantidade, " & vbCrLf & _
                 "       ISNULL(sb_REALIZADO.Quantidade, 0) AS QuantidadeFixado, P.Quantidade - ISNULL(sb_REALIZADO.Quantidade, 0) as QtdeFixar " & vbCrLf

        If Not Cedente Then
            strSQL &= ", COALESCE((SELECT COALESCE(SUM(CP.ValorLiquido), 0) " & vbCrLf & _
                                  "FROM ContasAPagar CP " & vbCrLf & _
                                  "WHERE CP.Empresa = P.Empresa_Id " & vbCrLf & _
                                  "AND CP.EndEmpresa = P.EndEmpresa_Id " & vbCrLf & _
                                  "AND CP.Pedido = P.Pedido" & strOutro & " " & vbCrLf & _
                                  "AND CP.Procuracao = P.Procuracao_Id " & vbCrLf & _
                                  "AND CP.Provisao = 1), 0) AS ValorFixado, " & vbCrLf & _
                      "COALESCE((SELECT SUM(PED.ValorOficial) " & vbCrLf & _
                                "FROM VW_PedidosXItensXFixacoesXEncargos PED " & vbCrLf & _
                                "WHERE PED.Empresa_Id = P.Empresa_Id " & vbCrLf & _
                                "AND PED.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
                                "AND PED.Pedido_Id = P.Pedido" & strOutro & " " & vbCrLf & _
                                "AND PED.Encargo_Id = 'LIQUIDO'), 0) AS ValorLiquido, " & vbCrLf &
                      "COALESCE((SELECT COALESCE(SUM(CP.ValorLiquido), 0) " & vbCrLf & _
                                  "FROM ContasAPagar CP " & vbCrLf & _
                                  "WHERE CP.Empresa = P.Empresa_Id " & vbCrLf & _
                                  "AND CP.EndEmpresa = P.EndEmpresa_Id " & vbCrLf & _
                                  "AND CP.Pedido = P.Pedido" & strOutro & " " & vbCrLf & _
                                  "AND CP.Procuracao = P.Procuracao_Id " & vbCrLf & _
                                  "AND CP.Provisao = 1), 0) - " & vbCrLf & _
                      "COALESCE((SELECT SUM(PED.ValorOficial) " & vbCrLf & _
                                "FROM VW_PedidosXItensXFixacoesXEncargos PED " & vbCrLf & _
                                "WHERE PED.Empresa_Id = P.Empresa_Id " & vbCrLf & _
                                "AND PED.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
                                "AND PED.Pedido_Id = P.Pedido" & strOutro & " " & vbCrLf & _
                                "AND PED.Encargo_Id = 'LIQUIDO'), 0) AS Saldo, " & vbCrLf & vbCrLf & _
                                "'00/00/0000' as Vencto" & vbCrLf
        End If

        strSQL &= "FROM Procuracoes P " & _
                  "   INNER JOIN Clientes C " & _
                  "           ON C.Cliente_Id = P." & strOutro & " " & _
                  "          AND C.Endereco_Id = P.End" & strOutro & " " & _
                  "   LEFT OUTER JOIN (SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Pedido, NotasFiscais.Procuracao, SUM(NotasFiscaisXItens.QuantidadeFisica) AS Quantidade " & vbCrLf & _
                  "                      FROM NotasFiscais " & vbCrLf & _
                  "                     INNER JOIN NotasFiscaisXItens " & vbCrLf & _
                  "                        ON NotasFiscais.Empresa_Id    = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                  "                       AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                  "                       AND NotasFiscais.Cliente_Id    = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                  "                       AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                  "                       AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                  "                       AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                  "                       AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                  "                     GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Pedido, NotasFiscais.Procuracao " & vbCrLf & _
                  "                    ) AS sb_REALIZADO " & vbCrLf & _
                  "                ON sb_REALIZADO.Empresa_Id    = P.Empresa_Id " & vbCrLf & _
                  "               AND sb_REALIZADO.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
                  "               AND sb_REALIZADO.Pedido        = P.PedidoCedente " & vbCrLf & _
                  "               AND sb_REALIZADO.Procuracao    = P.Procuracao_ID " & vbCrLf & _
                  "WHERE P.Empresa_Id = '" & _Pedido.CodigoEmpresa & "' " & vbCrLf & _
                  "AND P.EndEmpresa_Id = " & _Pedido.EnderecoEmpresa & " " & vbCrLf

        If Not String.IsNullOrWhiteSpace(_Pedido.Codigo) Then
            strSQL &= "AND P.Pedido" & strTipo & " IN (" & _Pedido.Codigo & ") " & vbCrLf
        End If

        strSQL &= "AND P.Situacao = 1"

        Dim dsProcuracoes As DataSet = Banco.ConsultaDataSet(strSQL, IIf(Cedente, "ProcuracoesCedente", "Procuracoes"))
        Return dsProcuracoes
    End Function

    Private Sub GetParamCabecalho(ByRef param As Dictionary(Of String, Object))
        Try
            Dim strSQL As String = "SELECT Nome, Cidade, Estado, Endereco, Inscricao " & _
                                   "FROM Clientes " & _
                                   "WHERE Cliente_Id = '" & _Pedido.CodigoEmpresa & "' " & _
                                   "AND Endereco_Id = " & _Pedido.EnderecoEmpresa

            Dim dsEmpresa As DataSet = Banco.ConsultaDataSet(strSQL, "Clientes")
            Dim drEmpresa As DataRow = dsEmpresa.Tables(0).Rows(0)

            Dim strEmpresa As String() = New String() {_Pedido.CodigoEmpresa, _Pedido.EnderecoEmpresa, drEmpresa("Nome"), drEmpresa("Cidade"), _
                                                       drEmpresa("Estado"), drEmpresa("Endereco"), drEmpresa("Inscricao")}
            param.Add("CabNomeEmpresa", strEmpresa(2))
            param.Add("CabEmpresaCidade", "CNPJ: " & String.Format("{0:00\.000\.000\./0000\-00}", strEmpresa(0)) & " - Cidade: " & strEmpresa(3) & "-" & strEmpresa(4))
            param.Add("CabCliente", "")
            param.Add("CabInscricao", "")

            'If Not String.IsNullOrWhiteSpace(txtClientes.Text) Then
            '    strSQL = "SELECT Nome, Cidade, Estado, Endereco, Complemento, Inscricao " & _
            '            "FROM Clientes " & _
            '            "WHERE Cliente_Id = '" & txtCodigoCliente.Value.Split("-")(0) & "' " & _
            '            "AND Endereco_Id = " & txtCodigoCliente.Value.Split("-")(1)

            '    Dim dsClientes As DataSet = Banco.ConsultaDataSet(strSQL, "Clientes")
            '    Dim drCliente As DataRow = dsClientes.Tables(0).Rows(0)
            '    Dim strCliente As String() = New String() {txtCodigoCliente.Value.Split("-")(0), txtCodigoCliente.Value.Split("-")(1), drCliente("Nome"), drCliente("Cidade"), _
            '                                               drCliente("Estado"), drCliente("Endereco"), drCliente("Complemento"), drCliente("Inscricao")}

            '    param("CabCliente") = Funcoes.FormatarCpfCnpj(strCliente(0)) & " - " & strCliente(2) & " - " & strCliente(3) & "/" & strCliente(4)

            '    If strCliente(6).ToString.Length > 0 Or strCliente(7).ToString.Length > 0 Then
            '        If strCliente(7).ToString.Length > 0 Then
            '            param("CabInscricao") = "INSCR: " & strCliente(7) & " - " & strCliente(6)
            '        End If
            '    End If
            'End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub getSumNota(ByVal ds As DataSet, ByRef param As Dictionary(Of String, Object))
        Dim pesoBalanca As Decimal
        Dim umidade As Decimal
        Dim impur As Decimal
        Dim avar As Decimal
        Dim ph As Decimal
        Dim gmo As Decimal
        Dim pesoLiquido As Decimal
        Dim pesoFiscal As Decimal
        Dim valorNota As Decimal
        Dim LiquidoNota As Decimal

        For Each row As DataRow In ds.Tables("NotasFiscais").Rows
            If row("EntradaSaida") = "E" Then
                pesoBalanca += IIf(String.IsNullOrWhiteSpace(row("PesoBalanca")), 0, row("PesoBalanca"))
                umidade += IIf(String.IsNullOrWhiteSpace(row("Umidade")), 0, row("Umidade"))
                impur += IIf(String.IsNullOrWhiteSpace(row("Impureza")), 0, row("Impureza"))
                avar += IIf(String.IsNullOrWhiteSpace(row("Avariados")), 0, row("Avariados"))
                ph += IIf(String.IsNullOrWhiteSpace(row("PH")), 0, row("PH"))
                gmo += IIf(String.IsNullOrWhiteSpace(row("GMO")), 0, row("GMO"))

                If row("Classe").Equals("GLOBAL") Then
                    pesoFiscal = row("PesoFiscal") 'Operação Global
                    valorNota = row("Valor")
                    LiquidoNota = row("Liquido")
                ElseIf row("Classe").Equals("REMESSAS") Then
                    pesoFiscal -= row("PesoFiscal")  'Operação Remessa
                    valorNota -= row("Valor")
                    LiquidoNota -= row("Liquido")
                Else
                    pesoFiscal += IIf(String.IsNullOrWhiteSpace(row("PesoFiscal")), 0, row("PesoFiscal"))
                    valorNota += IIf(String.IsNullOrWhiteSpace(row("Valor")), 0, row("Valor"))
                    LiquidoNota += IIf(String.IsNullOrWhiteSpace(row("Liquido")), 0, row("Liquido"))
                End If
                pesoLiquido += IIf(String.IsNullOrWhiteSpace(row("PesoLiquido")), 0, row("PesoLiquido"))
            Else
                pesoBalanca -= IIf(String.IsNullOrWhiteSpace(row("PesoBalanca")), 0, row("PesoBalanca"))
                umidade -= IIf(String.IsNullOrWhiteSpace(row("Umidade")), 0, row("Umidade"))
                impur -= IIf(String.IsNullOrWhiteSpace(row("Impureza")), 0, row("Impureza"))
                avar -= IIf(String.IsNullOrWhiteSpace(row("Avariados")), 0, row("Avariados"))
                ph -= IIf(String.IsNullOrWhiteSpace(row("PH")), 0, row("PH"))
                gmo -= IIf(String.IsNullOrWhiteSpace(row("GMO")), 0, row("GMO"))

                If row("Classe").Equals("GLOBAL") Then
                    pesoFiscal = row("PesoFiscal") 'Operação Global
                    valorNota = row("Valor")
                    LiquidoNota = row("Liquido")
                ElseIf row("Classe").Equals("REMESSAS") Then
                    pesoFiscal -= row("PesoFiscal")  'Operação Remessa
                    valorNota -= row("Valor")
                    LiquidoNota -= row("Liquido")
                Else
                    pesoFiscal -= IIf(String.IsNullOrWhiteSpace(row("PesoFiscal")), 0, row("PesoFiscal"))
                    valorNota -= IIf(String.IsNullOrWhiteSpace(row("Valor")), 0, row("Valor"))
                    LiquidoNota -= IIf(String.IsNullOrWhiteSpace(row("Liquido")), 0, row("Liquido"))
                End If
                pesoLiquido -= IIf(String.IsNullOrWhiteSpace(row("PesoLiquido")), 0, row("PesoLiquido"))
            End If
        Next

        param.Add("SumPesoBalanca", IIf(pesoBalanca < 0, pesoBalanca * -1, pesoBalanca))
        param.Add("SumUmidade", IIf(umidade < 0, umidade * -1, umidade))
        param.Add("SumImpureza", IIf(impur < 0, impur * -1, impur))
        param.Add("SumAvar", IIf(avar < 0, avar * -1, avar))
        param.Add("SumPH", IIf(ph < 0, ph * -1, ph))
        param.Add("SumGMO", IIf(gmo < 0, gmo * -1, gmo))
        param.Add("SumPesoLiquido", IIf(pesoLiquido < 0, pesoLiquido * -1, pesoLiquido))
        param.Add("SumPesoLiquidoPrd", IIf(pesoLiquido < 0, pesoLiquido * -1, pesoLiquido))
        param.Add("SumPesoNota", IIf(pesoFiscal < 0, pesoFiscal * -1, pesoFiscal))
        param.Add("SumPesoNotaPrd", IIf(pesoFiscal < 0, pesoFiscal * -1, pesoFiscal))
        param.Add("SumValorNota", IIf(valorNota < 0, valorNota * -1, valorNota))
        param.Add("SumValorNotaPrd", IIf(valorNota < 0, valorNota * -1, valorNota))
        param.Add("SumLiquidoNota", IIf(LiquidoNota < 0, LiquidoNota * -1, LiquidoNota))
        param.Add("SumLiquidoNotaPrd", IIf(LiquidoNota < 0, LiquidoNota * -1, LiquidoNota))
    End Sub

End Class
