Imports System.Web
Imports System.Web.UI

Public Class Extrato

    Public Shared Sub Emitir(ByVal page As Page, ByVal FinanceiroNovo As Boolean, ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal SituacaoPedido As String, Optional ByVal CodigoPedido As String = "",
                                       Optional ByVal CodigoCliente As String = "", Optional ByVal EndCliente As Integer = 0, Optional ByVal Safra As String = "",
                                       Optional ByVal DataLimite As Date? = Nothing, Optional ByVal CodigoProduto As String = "", Optional ByVal Entrada As Boolean = False,
                                       Optional ByVal Saida As Boolean = False, Optional ByVal MostraDescProd As Boolean = False, Optional ByVal OcultarFinanceiro As Boolean = False,
                                       Optional ByVal ocultarLancamentoContabil As Boolean = False, Optional ByVal Sintetico As Boolean = False, Optional ByVal OcultarFrete As Boolean = False, Optional ByVal Pesagem As Boolean = False)
        Try

            If String.IsNullOrWhiteSpace(CodigoPedido) AndAlso String.IsNullOrWhiteSpace(CodigoCliente) Then
                Throw New Exception("Informe o Cliente ou o número do pedido.")
            End If

            Dim ds As DataSet = getDataSetPedido(CodigoEmpresa, EndEmpresa, CodigoPedido, CodigoCliente, EndCliente, Safra, DataLimite, CodigoProduto, SituacaoPedido)

            If ds Is Nothing OrElse ds.Tables("Pedidos").Rows.Count = 0 Then
                Throw New Exception("Pedido não encontrado ou com situação diferente de Normal.")
            End If

            Dim param As New Dictionary(Of String, Object)
            param.Add("ParametersConsulta", getParametersConsulta(CodigoEmpresa, EndEmpresa, CodigoPedido, CodigoCliente, EndCliente, Safra, DataLimite, CodigoProduto, Entrada, Saida, MostraDescProd))
            param.Add("EsconderFinanceiro", OcultarFinanceiro)
            param.Add("EsconderFrete", OcultarFrete)
            param.Add("EsconderLancamentoContabil", ocultarLancamentoContabil)
            param.Add("Sintetico", Sintetico)
            param.Add("EsconderPesagem", Pesagem)

            For Each rowPedido As DataRow In ds.Tables("Pedidos").Rows
                ds.Merge(getDataSetContratos(CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente, SituacaoPedido))
                ds.Merge(getDataSetFixacoes(CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente, SituacaoPedido))
                ds.Merge(getDataSetFinanceiro(CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente, Sintetico, SituacaoPedido))
                ds.Merge(getDataSetAdiantamentos(CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente, SituacaoPedido))
                ds.Merge(getDataSetRazao(CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente, SituacaoPedido))
                ds.Merge(getDataSetResumoFinanceiro(CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente)) 'Vazio
                ds.Merge(getDataSetProcuracoes(True, CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente))
                ds.Merge(getDataSetProcuracoes(False, CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente))
                ds.Merge(getDataSetNotasFiscais(CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente, Entrada, Saida, DataLimite, SituacaoPedido))
                ds.Merge(getDataSetPesagem(CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente))
                ds.Merge(getDataSetResumoNotasFiscais(CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente, Entrada, Saida, DataLimite, SituacaoPedido))
                ds.Merge(getDataSetResumoNotasFiscaisPrd(CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente, Entrada, Saida, DataLimite, SituacaoPedido))
                ds.Merge(getDataSetFretes(CodigoEmpresa, EndEmpresa, rowPedido("Pedido"), CodigoCliente, EndCliente, Entrada, Saida, DataLimite))
                ds.Merge(getDataSetTotalizadorImpostos(CodigoEmpresa, EndEmpresa, rowPedido("Pedido")))
                getDataSetResumoPedido(ds, rowPedido("Pedido"))
            Next

            Funcoes.BindReport(page, ds, "Cr_ExtratoDePedido_Novo", eExportType.PDF, param, False, "", CodigoEmpresa, EndEmpresa)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Shared Function getParametersConsulta(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String,
                                                  ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal Safra As String,
                                                  ByVal DataLimite As Date?, ByVal CodigoProduto As String,
                                                  ByVal Entrada As Boolean, ByVal Saida As Boolean, ByVal MostraDescProd As Boolean) As String
        Dim obj As Cliente = New Cliente(CodigoEmpresa, EndEmpresa)

        Dim param As String = "Empresa: " & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & " / " & obj.Estado.Codigo

        If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
            obj = New Cliente(CodigoCliente, EndCliente)
            param &= vbCrLf & "Cliente: " & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & " / " & obj.Estado.Codigo
        End If

        If MostraDescProd Then
            param &= " - Parâmetros: Mostrar descrição do produto"
        End If

        If DataLimite IsNot Nothing Then
            param &= " - Data Limite: " & CDate(DataLimite).ToString("dd/MM/yyyy")
        End If

        If Not String.IsNullOrWhiteSpace(Safra) Then
            param &= " - Safra: " & Safra
        End If

        If Not String.IsNullOrWhiteSpace(CodigoPedido) Then
            param &= " - Pedido(s): " & CodigoPedido
        End If

        If Entrada OrElse Saida Then
            If Entrada AndAlso Saida Then
                param &= " - Entrada/Saída: Entrada e Saída"
            Else
                param &= " - Entrada/Saída: " & IIf(Entrada, "Entrada", "Saída")
            End If
        End If

        Return param
    End Function

    Private Shared Function getDataSetPedido(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer,
                                             ByVal Safra As String, ByVal DataLimite As Date?, ByVal CodigoProduto As String, ByVal SituacaoPedido As String) As DataSet
        Try
            Dim sql As String = "SELECT P.Empresa_Id as Empresa , P.EndEmpresa_Id as EndEmpresa, P.Pedido_Id as Pedido, so.EntradaSaida, so.Descricao as DescricaoOperacao," & vbCrLf & _
                                "       P.UnidadeDeNegocio, P.Cliente, P.EndCliente, P.Praca, P.EndPraca, P.PedidoEfetivo, P.Safra, " & vbCrLf & _
                                "       P.Moeda, P.Indexador, P.Operacao, P.SubOperacao, P.Situacao, P.DataPedido, P.DataEntrega, "

            If Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "24450490" Then
                sql &= "P.PedidoEfetivo AS PedidoOrigem, " & vbCrLf
            Else
                sql &= "P.PedidoOrigem, " & vbCrLf
            End If

            sql &= "       P.FreteCIFFOB, isnull(P.OrigemDestino,'') AS OrigemDestino, P.Solicitacao, P.UsuarioInclusao, P.UsuarioInclusaoData, isnull(P.UsuarioAlteracao,'') AS UsuarioAlteracao, " & vbCrLf & _
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
                                "       dbo.FormatarCpfCnpj(Cli.Cliente_Id) + ' - ' + RTRIM(CONVERT(CHAR,Cli.Endereco_Id)) + '   ' +  Cli.Nome + " & vbCrLf & _
                                "       CASE WHEN Cli.Complemento<>'' THEN ' - FAZ. ' ELSE '' END + Cli.Complemento +' - ' + Cli.Cidade  + '/' + Cli.Estado + ' - INSCRIÇÃO: ' + Cli.Inscricao as NomeCliente, " & vbCrLf & _
                                "       SO.PrecoFixo, SO.Classe, M.Descricao AS NomeMoeda" & vbCrLf & _
                                "  FROM Pedidos P " & vbCrLf & _
                                " INNER JOIN Moedas M " & vbCrLf & _
                                "    ON M.Moeda_Id = P.Moeda  " & vbCrLf & _
                                " INNER JOIN SubOperacoes SO " & vbCrLf & _
                                "    ON SO.Operacao_Id = P.Operacao " & vbCrLf & _
                                "   AND SO.SubOperacoes_Id = P.SubOperacao " & vbCrLf & _
                                " Inner Join Clientes cli" & vbCrLf & _
                                "    ON cli.Cliente_Id = P.Cliente" & vbCrLf & _
                                "   And cli.Endereco_Id = P.EndCliente" & vbCrLf & _
                                " WHERE P.UnidadeDeNegocio IS NOT NULL " & vbCrLf & _
                                "       AND P.PedidoOrigem = 0 " & vbCrLf & _
                                "       AND P.Situacao = 1" & vbCrLf

            '"       AND EXISTS(SELECT 1 FROM NotasFiscais WHERE isnull(NFG,0) = 0 and Empresa_Id = P.Empresa_Id AND Endempresa_id = P.EndEmpresa_Id AND Pedido = P.Pedido_Id)" & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then sql &= "    AND P.Pedido_Id in (" & CodigoPedido.Trim & ")" & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoEmpresa) Then
                sql &= "    AND P.Empresa_Id    = '" & CodigoEmpresa & "' " & vbCrLf & _
                       "    AND P.EndEmpresa_Id = " & EndEmpresa & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                sql &= "    AND P.Cliente    = '" & CodigoCliente & "' " & vbCrLf & _
                       "    AND P.EndCliente = " & EndCliente.ToString() & vbCrLf
            End If

            If SituacaoPedido = "A" Then
                sql &= "   AND (P.FiscalAberto = 'true' OR P.FinanceiroAberto = 'true') " & vbCrLf
            ElseIf SituacaoPedido = "F" Then
                sql &= "   AND P.FiscalAberto = 'false' AND P.FinanceiroAberto = 'false' " & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(Safra) Then sql &= "AND P.Safra = '" & Safra & "'" & vbCrLf

            If DataLimite IsNot Nothing Then sql &= "AND P.DataPedido <= '" & CDate(DataLimite).ToString("yyyy-MM-dd") & "'" & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoProduto) Then
                sql &= "AND EXISTS (SELECT NULL " & vbCrLf & _
                       "             FROM PedidoXItem PIT " & vbCrLf & _
                       "             WHERE P.Empresa_Id = PIT.Empresa_Id " & vbCrLf & _
                       "                 AND P.EndEmpresa_Id = PIT.EndEmpresa_Id " & vbCrLf & _
                       "                 AND P.Pedido_Id = PIT.Pedido_Id " & vbCrLf & _
                       "                 AND PIT.Produto_Id = '" & CodigoProduto & "' )" & vbCrLf
            End If
            sql &= "ORDER BY DataPedido" & vbCrLf

            Return New AcessaBanco().ConsultaDataSet(sql, "Pedidos")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getDataSetContratos_BKP(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal SituacaoPedido As String) As DataSet
        Try
            Dim strSQL As String
            strSQL = "  SELECT case when TipoDeLancamento = 'N' then 1" & vbCrLf &
                     "              when TipoDeLancamento = 'C' then 2" & vbCrLf &
                     "              else 3" & vbCrLf &
                     "         end as Ordem, " & vbCrLf &
                     "         PIT.Movimento as DataPedido, PIT.TipoDeLancamento, PIT.Pedido_Id as Pedido, sum(PIT.Quantidade) as Quantidade, M.Descricao AS NomeMoeda," & vbCrLf &
                     "         PIT.UnitarioOficial, PIT.UnitarioMoeda, sum(PIT.TotalOficial) as TotalOficial, sum(PIT.TotalMoeda) as TotalMoeda,                       " & vbCrLf


            Dim tEmpresas As ClienteXEmpresa = New ClienteXEmpresa(CodigoEmpresa, 0)

            If tEmpresas.UsarRegistroMinAgr Then
                strSQL &= "		 case" & vbCrLf &
                          "			when len(isnull(Produtos.RegMinAgr,'')) > 0" & vbCrLf &
                          "				then PIT.Produto_Id + '-' + Produtos.Nome + '(' + Produtos.RegMinAgr  + ')'" & vbCrLf &
                          "				else PIT.Produto_Id + '-' + Produtos.Nome" & vbCrLf &
                          "			end AS NomeProduto," & vbCrLf

            ElseIf tEmpresas.UsarDescricaoProduto Then
                strSQL &= "		 case" & vbCrLf &
                          "			when Produtos.Nome = Produtos.Descricao" & vbCrLf &
                          "				then PIT.Produto_Id + '-' + Produtos.Nome" & vbCrLf &
                          "				else PIT.Produto_Id + '-' + Produtos.Nome + '-' + Produtos.Descricao" & vbCrLf &
                          "			end AS NomeProduto," & vbCrLf
            Else
                strSQL &= "         PIT.Produto_Id + '-' + Produtos.Nome AS NomeProduto, " & vbCrLf
            End If

            strSQL &= "         PIT.Produto_Id as Produto, isnull(P.IndiceFixado,0) AS IndiceFixado,               " & vbCrLf &
                     "         P.Cliente, P.EndCliente, sum(PxE.ValorOficial) AS TotalOficialLiquido, sum(PxE.ValorMoeda) AS TotalMoedaLiquido                         " & vbCrLf &
                     "    FROM PedidoXItemXLancamento PIT               " & vbCrLf &
                     "   INNER JOIN Pedidos P                           " & vbCrLf &
                     "      ON P.Empresa_Id    = PIT.Empresa_Id         " & vbCrLf &
                     "     AND P.EndEmpresa_Id = PIT.EndEmpresa_Id      " & vbCrLf &
                     "     AND P.Pedido_Id     = PIT.Pedido_Id          " & vbCrLf &
                     "   INNER JOIN PedidosXEncargos AS PxE             " & vbCrLf &
                     "      on PxE.Empresa_Id    = P.Empresa_Id         " & vbCrLf &
                     "     and PxE.EndEmpresa_Id = P.EndEmpresa_Id      " & vbCrLf &
                     "     and PxE.Pedido_Id     = P.Pedido_Id          " & vbCrLf &
                     "     and PxE.Encargo_Id    = 'LIQUIDO'            " & vbCrLf &
                     "     and PxE.Produto_Id    = PIT.Produto_Id       " & vbCrLf &
                     "   INNER JOIN Moedas M                            " & vbCrLf &
                     "      ON M.Moeda_Id = P.Moeda                     " & vbCrLf &
                     "   INNER JOIN Produtos                            " & vbCrLf &
                     "      ON PIT.Produto_Id = Produtos.Produto_Id     " & vbCrLf &
                     "   WHERE P.Empresa_Id    = '" & CodigoEmpresa & "'" & vbCrLf &
                     "     AND P.EndEmpresa_Id = " & EndEmpresa & "     " & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= "     AND P.Pedido_Id IN (" & CodigoPedido.Trim & ") " & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                strSQL &= "     AND P.Cliente    = '" & CodigoCliente & "'" & vbCrLf &
                          "     AND P.EndCliente =  " & EndCliente & vbCrLf
            End If

            If SituacaoPedido = "A" Then
                strSQL &= "   AND (P.FiscalAberto = 'true' OR P.FinanceiroAberto = 'true') " & vbCrLf
            ElseIf SituacaoPedido = "F" Then
                strSQL &= "   AND P.FiscalAberto = 'false' AND P.FinanceiroAberto = 'false' " & vbCrLf
            End If

            strSQL &= "group by PIT.Movimento, PIT.TipoDeLancamento, PIT.Pedido_Id, M.Descricao, " & vbCrLf &
                      "         PIT.UnitarioOficial, PIT.UnitarioMoeda," & vbCrLf

            If tEmpresas.UsarRegistroMinAgr Then
                strSQL &= "PIT.Produto_Id, Produtos.Nome, Produtos.RegMinAgr, " & vbCrLf
            ElseIf tEmpresas.UsarDescricaoProduto Then
                strSQL &= "PIT.Produto_Id, Produtos.Nome, Produtos.Descricao, " & vbCrLf
            Else
                strSQL &= "PIT.Produto_Id, Produtos.Nome, " & vbCrLf
            End If

            strSQL &= "         isnull(P.IndiceFixado,0), P.Cliente, P.EndCliente" & vbCrLf &
                      "order by DataPedido, Ordem" & vbCrLf

            Return New AcessaBanco().ConsultaDataSet(strSQL, "Contratos")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getDataSetContratos(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal SituacaoPedido As String) As DataSet
        Try
            Dim strSQL As String
            strSQL = "SELECT CASE WHEN TipoDeLancamento = 'N' THEN 1" & vbCrLf &
                 "            WHEN TipoDeLancamento = 'C' THEN 2" & vbCrLf &
                 "            ELSE 3 END AS Ordem," & vbCrLf &
                 "       PIT.Movimento AS DataPedido, PIT.TipoDeLancamento, PIT.Pedido_Id AS Pedido, SUM(PIT.Quantidade) AS Quantidade, M.Descricao AS NomeMoeda," & vbCrLf &
                 "       PIT.UnitarioOficial, PIT.UnitarioMoeda, SUM(PIT.TotalOficial) AS TotalOficial, SUM(PIT.TotalMoeda) AS TotalMoeda," & vbCrLf

            Dim tEmpresas As ClienteXEmpresa = New ClienteXEmpresa(CodigoEmpresa, 0)

            If tEmpresas.UsarRegistroMinAgr Then
                strSQL &= "       CASE WHEN LEN(ISNULL(Produtos.RegMinAgr, '')) > 0" & vbCrLf &
                      "            THEN PIT.Produto_Id + '-' + Produtos.Nome + '(' + Produtos.RegMinAgr + ')'" & vbCrLf &
                      "            ELSE PIT.Produto_Id + '-' + Produtos.Nome END AS NomeProduto," & vbCrLf
            ElseIf tEmpresas.UsarDescricaoProduto Then
                strSQL &= "       CASE WHEN Produtos.Nome = Produtos.Descricao" & vbCrLf &
                      "            THEN PIT.Produto_Id + '-' + Produtos.Nome" & vbCrLf &
                      "            ELSE PIT.Produto_Id + '-' + Produtos.Nome + '-' + Produtos.Descricao END AS NomeProduto," & vbCrLf
            Else
                strSQL &= "       PIT.Produto_Id + '-' + Produtos.Nome AS NomeProduto," & vbCrLf
            End If

            strSQL &= "       PIT.Produto_Id AS Produto, ISNULL(P.IndiceFixado, 0) AS IndiceFixado," & vbCrLf &
                  "       P.Cliente, P.EndCliente," & vbCrLf &
                  "       ISNULL(LIQ.ValorOficial, 0) AS TotalOficialLiquido, ISNULL(LIQ.ValorMoeda, 0) AS TotalMoedaLiquido" & vbCrLf &
                  "FROM PedidoXItemXLancamento PIT" & vbCrLf &
                  "INNER JOIN Pedidos P ON P.Empresa_Id = PIT.Empresa_Id AND P.EndEmpresa_Id = PIT.EndEmpresa_Id AND P.Pedido_Id = PIT.Pedido_Id" & vbCrLf &
                  "LEFT JOIN (" & vbCrLf &
                  "    SELECT Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id," & vbCrLf &
                  "           SUM(ValorOficial) AS ValorOficial, SUM(ValorMoeda) AS ValorMoeda" & vbCrLf &
                  "      FROM PedidosXEncargos" & vbCrLf &
                  "     WHERE Encargo_Id = 'LIQUIDO'" & vbCrLf &
                  "     GROUP BY Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id" & vbCrLf &
                  ") LIQ ON LIQ.Empresa_Id = P.Empresa_Id AND LIQ.EndEmpresa_Id = P.EndEmpresa_Id AND LIQ.Pedido_Id = P.Pedido_Id AND LIQ.Produto_Id = PIT.Produto_Id" & vbCrLf &
                  "INNER JOIN Moedas M ON M.Moeda_Id = P.Moeda" & vbCrLf &
                  "INNER JOIN Produtos ON PIT.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                  "WHERE P.Empresa_Id = '" & CodigoEmpresa & "'" & vbCrLf &
                  "  AND P.EndEmpresa_Id = " & EndEmpresa & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then
                strSQL &= "  AND P.Pedido_Id IN (" & CodigoPedido.Trim & ")" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                strSQL &= "  AND P.Cliente = '" & CodigoCliente & "' AND P.EndCliente = " & EndCliente & vbCrLf
            End If

            If SituacaoPedido = "A" Then
                strSQL &= "  AND (P.FiscalAberto = 'true' OR P.FinanceiroAberto = 'true')" & vbCrLf
            ElseIf SituacaoPedido = "F" Then
                strSQL &= "  AND P.FiscalAberto = 'false' AND P.FinanceiroAberto = 'false'" & vbCrLf
            End If

            strSQL &= "GROUP BY PIT.Movimento, PIT.TipoDeLancamento, PIT.Pedido_Id, M.Descricao," & vbCrLf &
                  "         PIT.UnitarioOficial, PIT.UnitarioMoeda," & vbCrLf

            If tEmpresas.UsarRegistroMinAgr Then
                strSQL &= "         PIT.Produto_Id, Produtos.Nome, Produtos.RegMinAgr," & vbCrLf
            ElseIf tEmpresas.UsarDescricaoProduto Then
                strSQL &= "         PIT.Produto_Id, Produtos.Nome, Produtos.Descricao," & vbCrLf
            Else
                strSQL &= "         PIT.Produto_Id, Produtos.Nome," & vbCrLf
            End If

            strSQL &= "         ISNULL(P.IndiceFixado, 0), P.Cliente, P.EndCliente, LIQ.ValorOficial, LIQ.ValorMoeda" & vbCrLf &
                  "ORDER BY DataPedido, Ordem" & vbCrLf

            Return New AcessaBanco().ConsultaDataSet(strSQL, "Contratos")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function


    Private Shared Function getDataSetFixacoes(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal SituacaoPedido As String) As DataSet
        Try
            Dim strSQL As String
            strSQL = "SELECT f.Pedido_id as Pedido, f.Fixacao_Id as Fixacao, f.Procuracao, f.Operacao, f.SubOperacao, f.Documento, f.Movimento, f.Quantidade," & vbCrLf & _
                     "       f.UnitarioOficial, f.UnitarioMoeda, f.TotalOficial, f.TotalMoeda, f.CondicaoPagamento, f.IndiceFixado, f.Produto_Id + '-' + p.Descricao as Produto, " & vbCrLf & _
                     "       m.Descricao as Moeda, FE.ValorOficial AS Liquido, FE.ValorMoeda AS LiquidoMoeda, Pd.Cliente, Pd.EndCliente" & vbCrLf & _
                     "  FROM VW_PedidosXItensXFixacoes f" & vbCrLf & _
                     " Inner Join Produtos p" & vbCrLf & _
                     "    On p.Produto_Id = f.Produto_Id" & vbCrLf & _
                     " Inner Join Pedidos Pd" & vbCrLf & _
                     "    on Pd.empresa_Id    = f.Empresa_Id" & vbCrLf & _
                     " 	 and pd.EndEmpresa_Id = f.EndEmpresa_Id" & vbCrLf & _
                     " 	 and pd.Pedido_Id     = f.Pedido_Id" & vbCrLf
            '" Inner JOin PedidosXItens pxi" & vbCrLf & _
            '" 	  On pxi.Empresa_Id    = f.Empresa_Id" & vbCrLf & _
            '" 	 and pxi.EndEmpresa_Id = f.EndEmpresa_Id" & vbCrLf & _
            '" 	 and pxi.Pedido_Id     = f.Pedido_Id" & vbCrLf & _
            '" 	 and pxi.Sequencia_Id  = f.Sequencia_Id" & vbCrLf & _
            '"   and pxi.Produto_Id    = f.Produto_Id" & vbCrLf & _
            strSQL &= " Inner Join VW_PedidosXItensXFixacoesXEncargos fe" & vbCrLf & _
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
                     "   AND f.Empresa_Id    = '" & CodigoEmpresa & "'" & vbCrLf & _
                     "   AND f.EndEmpresa_Id = " & EndEmpresa & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= "AND f.Pedido_Id IN (" & CodigoPedido.Trim & ") " & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                strSQL &= "AND Pd.Cliente = '" & CodigoCliente & "' " & vbCrLf & _
                          "AND Pd.EndCliente = " & EndCliente.ToString() & vbCrLf
            End If

            If SituacaoPedido = "A" Then
                strSQL &= "   AND (Pd.FiscalAberto = 'true' OR Pd.FinanceiroAberto = 'true') " & vbCrLf
            ElseIf SituacaoPedido = "F" Then
                strSQL &= "   AND Pd.FiscalAberto = 'false' AND Pd.FinanceiroAberto = 'false' " & vbCrLf
            End If

            strSQL &= "    ORDER BY f.Pedido_id, f.Fixacao_Id" & vbCrLf

            Return New AcessaBanco().ConsultaDataSet(strSQL, "Fixacoes")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getDataSetFinanceiro(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal sintetico As Boolean, ByVal SituacaoPedido As String) As DataSet
        Try
            Dim strSQL As String = ""

            strSQL = "SELECT T.Movimento, T.Prorrogacao AS Vencimento, T.Baixa," & vbCrLf & _
                     "       isnull(T.ValorDoDocumento, 0) AS ValorDoDocumento," & vbCrLf & _
                     "       isnull(T.Descontos, 0) AS Descontos," & vbCrLf & _
                     "       isnull(T.Deducoes, 0) AS Deducoes," & vbCrLf & _
                     "       isnull(T.Juros, 0) AS Juros," & vbCrLf & _
                     "       isnull(T.Acrescimos, 0) AS Acrescimos," & vbCrLf & _
                     "       isnull(Case" & vbCrLf & _
                     "                  when T.Provisao = 1" & vbCrLf & _
                     "                    then T.ValorLiquido" & vbCrLf & _
                     "                    else Case" & vbCrLf & _
                     "                           when T.Moeda <> 1" & vbCrLf & _
                     "                             then convert(numeric(18,2), T.MoedaValorDoDocumento * Cotacoes.indice)" & vbCrLf & _
                     "                             else T.ValorDoDocumento" & vbCrLf & _
                     "                         end" & vbCrLf & _
                     "                end, 0) AS ValorLiquido," & vbCrLf & _
                     "       isnull(T.MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento," & vbCrLf & _
                     "       isnull(T.MoedaDescontos, 0) AS MoedaDescontos," & vbCrLf & _
                     "       isnull(T.MoedaDeducoes, 0) AS MoedaDeducoes," & vbCrLf & _
                     "       isnull(T.MoedaJuros, 0) AS MoedaJuros," & vbCrLf & _
                     "       isnull(T.MoedaAcrescimos, 0) AS MoedaAcrescimos," & vbCrLf & _
                     "       isnull(Case" & vbCrLf & _
                     "                  when T.Provisao = 1" & vbCrLf & _
                     "                    then T.MoedaValorLiquido" & vbCrLf & _
                     "                    else Case" & vbCrLf & _
                     "                           when T.Moeda = 1" & vbCrLf & _
                     "                             then convert(numeric(18,2),T.ValorDoDocumento / Cotacoes.indice)" & vbCrLf & _
                     "                             else T.MoedaValorDoDocumento" & vbCrLf & _
                     "                         end" & vbCrLf & _
                     "                end, 0) AS MoedaValorLiquido," & vbCrLf & _
                     "       T.Provisao," & vbCrLf & _
                     "	   case" & vbCrLf & _
                     "			when isnull(nXt.Nota_Id,0) > 0" & vbCrLf & _
                     "				then convert(nvarchar,T.Registro_Id) + ' (' + convert(nvarchar,isnull(nXt.Nota_Id,0)) + ')'" & vbCrLf & _
                     "				else convert(nvarchar,T.Registro_Id)" & vbCrLf & _
                     "			end AS Registro," & vbCrLf & _
                     "        case" & vbCrLf & _
                     "		  when exists(Select 1 from AdiantamentosXBaixas where titulo = T.registro_id) and cXp.Adiantamento = 'S' then Tipo + '|BA'" & vbCrLf & _
                     "		  else Tipo" & vbCrLf & _
                     "		end  AS Tipo," & vbCrLf & _
                     "	   T.Moeda, T.Cliente, T.EndCliente, T.Pedido" & vbCrLf & _
                     "  INTO #Financeiro" & vbCrLf & _
                     "  FROM (Select 'P' as Tipo, Empresa, EndEmpresa, Registro_Id, Situacao, Provisao, Grupado, Cliente, EndCliente, Pedido, Moeda, Indexador, Movimento, Prorrogacao, baixa, ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos , MoedaValorLiquido, Carteira, EmpresaPedido, EndEmpresaPedido" & vbCrLf & _
                     "          from ContasAPagar" & vbCrLf & _
                     "		 Union All" & vbCrLf & _
                     "		Select 'R',Empresa, EndEmpresa, Registro_Id, Situacao, Provisao, Grupado, Cliente, EndCliente, Pedido, Moeda, Indexador, Movimento, Prorrogacao, baixa, ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos , MoedaValorLiquido, Carteira, EmpresaPedido, EndEmpresaPedido" & vbCrLf & _
                     "          from ContasAReceber) T " & vbCrLf & _
                     "  inner Join ComprasXProdutos cXp" & vbCrLf & _
                     "    on cXp.Produto_id = T.Carteira" & vbCrLf & _
                     "  left Join NotaFiscalXTitulo nXt" & vbCrLf & _
                     "    on  nXt.Empresa_Id      = T.EmpresaPedido" & vbCrLf & _
                     "   and nXt.EndEmpresa_Id    = T.EndEmpresaPedido" & vbCrLf & _
                     "   and nXt.Titulo_Id        = T.Registro_Id" & vbCrLf & _
                     "  left Join Cotacoes " & vbCrLf & _
                     "    on Cotacoes.Data_id = CASE" & vbCrLf & _
                     "                            WHEN T.Provisao = 1" & vbCrLf & _
                     "       		                THEN T.baixa" & vbCrLf & _
                     "       			            ELSE T.Prorrogacao" & vbCrLf & _
                     "       			        END" & vbCrLf & _
                     "   and Cotacoes.Indexador_Id = T.Indexador" & vbCrLf & _
                     " WHERE T.Empresa    ='" & CodigoEmpresa & "' " & vbCrLf & _
                     "   AND T.EndEmpresa = " & EndEmpresa & vbCrLf & _
                     "   AND T.Situacao   IN(1,101,102)" & vbCrLf & _
                     "   AND isnull(T.Grupado,'N') <> 'M'" & vbCrLf & _
                     "   AND not exists(Select 1 from Adiantamentos where titulo = T.registro_id)" & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                strSQL &= "   AND T.Cliente in ('" & CodigoCliente & "') " & vbCrLf
                strSQL &= "   AND T.EndCliente in (" & EndCliente & ") " & vbCrLf
            End If
            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= "   AND T.Pedido in (" & CodigoPedido & ") " & vbCrLf

            strSQL &= "Update #Financeiro Set " & vbCrLf & _
                      "   Acrescimos      = CASE" & vbCrLf & _
                      "                       WHEN Moeda <> 1 AND ValorLiquido > ValorDoDocumento " & vbCrLf & _
                      "       		     	    THEN ValorLiquido - ValorDoDocumento " & vbCrLf & _
                      "                         ELSE CASE" & vbCrLf & _
                      "                                WHEN Moeda <> 1 AND Acrescimos > 0 " & vbCrLf & _
                      "                                  THEN 0  " & vbCrLf & _
                      "                                  ELSE Acrescimos " & vbCrLf & _
                      "                              END  " & vbCrLf & _
                      "       		        END" & vbCrLf & _
                      "   ,Deducoes       = CASE" & vbCrLf & _
                      "                       WHEN Moeda <> 1 AND ValorLiquido < ValorDoDocumento " & vbCrLf & _
                      "       		     	    THEN ValorDoDocumento - ValorLiquido " & vbCrLf & _
                      "                         ELSE CASE" & vbCrLf & _
                      "                                WHEN Moeda <> 1 AND Deducoes > 0 " & vbCrLf & _
                      "                                  THEN 0 " & vbCrLf & _
                      "                                  ELSE Deducoes " & vbCrLf & _
                      "                              END " & vbCrLf & _
                      "       		        END" & vbCrLf & _
                      "  ,MoedaAcrescimos = CASE" & vbCrLf & _
                      "                       WHEN Moeda <> 1 " & vbCrLf & _
                      "       			        THEN 0 " & vbCrLf & _
                      "                         ELSE MoedaAcrescimos " & vbCrLf & _
                      "       			    END, " & vbCrLf & _
                      "  MoedaDeducoes    = CASE" & vbCrLf & _
                      "                       WHEN Moeda <> 1 " & vbCrLf & _
                      "       			        THEN 0 " & vbCrLf & _
                      "                         ELSE MoedaDeducoes " & vbCrLf & _
                      "       			    END " & vbCrLf & _
                      " Where Provisao <> 1;" & vbCrLf & vbCrLf

            strSQL &= "INSERT INTO #Financeiro" & vbCrLf & _
                      "Select Vencimento," & vbCrLf & _
                      "       Vencimento," & vbCrLf & _
                      "       Vencimento," & vbCrLf & _
                      "	      ParcelaOficial," & vbCrLf & _
                      "	      0,0,0,0," & vbCrLf & _
                      "	      ParcelaOficial," & vbCrLf & _
                      "	      ParcelaMoeda," & vbCrLf & _
                      "	      0,0,0,0," & vbCrLf & _
                      "	      ParcelaMoeda," & vbCrLf & _
                      "	      3," & vbCrLf & _
                      "	      'Parc.-' + convert(nvarchar,Parcela_id)," & vbCrLf & _
                      "	      cr," & vbCrLf & _
                      "	      moeda," & vbCrLf & _
                      "	      cliente," & vbCrLf & _
                      "	      endcliente," & vbCrLf & _
                      "	      pedido_id" & vbCrLf & _
                      "  from VW_TituloVirtual" & vbCrLf & _
                      " WHERE Empresa_id    ='" & CodigoEmpresa & "' " & vbCrLf & _
                      "   AND EndEmpresa_id = " & EndEmpresa & " " & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                strSQL &= "   AND cliente in ('" & CodigoCliente & "') " & vbCrLf
                strSQL &= "   AND endcliente in (" & EndCliente & ") " & vbCrLf
            End If
            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= "    AND Pedido_id in (" & CodigoPedido & ") " & vbCrLf

            strSQL &= "SELECT f.Movimento, f.Vencimento, f.Baixa, Sum(f.ValorDoDocumento) as ValorDoDocumento, Sum(f.Descontos) as Descontos, Sum(f.Deducoes) as Deducoes, Sum(f.Juros) as Juros, " & vbCrLf & _
                      "       Sum(f.Acrescimos) as Acrescimos, Sum(f.ValorLiquido) as ValorLiquido, Sum(f.MoedaValorDoDocumento) as MoedaValorDoDocumento, Sum(f.MoedaDescontos) as MoedaDescontos," & vbCrLf & _
                      "       Sum(f.MoedaDeducoes) as MoedaDeducoes, Sum(f.MoedaJuros) as MoedaJuros, Sum(f.MoedaAcrescimos) as MoedaAcrescimos, Sum(f.MoedaValorLiquido) as MoedaValorLiquido," & vbCrLf & _
                      "       f.Provisao, " & IIf(sintetico, "0 as Registro", "f.Registro") & ", f.Tipo, f.Moeda, f.Cliente, f.EndCliente, f.Pedido" & vbCrLf & _
                      "FROM #Financeiro f " & vbCrLf & _
                      "	inner join Pedidos ped " & vbCrLf & _
                      "			on ped.Empresa_Id     ='" & CodigoEmpresa & "'" & vbCrLf & _
                      "			and ped.EndEmpresa_id = " & EndEmpresa & vbCrLf & _
                      "			and ped.Pedido_id     = f.Pedido " & vbCrLf & _
                      "Where f.Tipo <> 'RA' and f.Tipo <> 'P|BA' " & vbCrLf

            If SituacaoPedido = "A" Then
                strSQL &= "   AND (ped.FiscalAberto = 'true' OR ped.FinanceiroAberto = 'true') " & vbCrLf
            ElseIf SituacaoPedido = "F" Then
                strSQL &= "   AND ped.FiscalAberto = 'false' AND ped.FinanceiroAberto = 'false' " & vbCrLf
            End If

            strSQL &= " group by f.Movimento, f.Vencimento, f.Provisao, " & IIf(sintetico, "", "f.Registro") & ", f.Tipo, f.Moeda, f.Cliente, f.EndCliente, f.Pedido, f.Baixa" & vbCrLf & _
                      " order BY f.Movimento " & vbCrLf

            Return New AcessaBanco().ConsultaDataSet(strSQL, "Financeiro")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getDataSetAdiantamentos(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal SituacaoPedido As String) As DataSet
        Try
            Dim strSQL As String = ""

            strSQL = "SELECT 1 AS Ordem, A.Baixa, A.Movimento, A.Vencimento, A.Adiantamento_Id as Adiantamento, A.Indexador, a.Cliente_ID as Cliente, a.EndCliente_ID as EndCliente, A.Taxa, A.ValorOficial, A.ValorMoeda, 0.00 AS Correcao, " & vbCrLf & _
                "	         'A' as Tipo, ('A-' + convert(varchar,A.Titulo)) AS Titulo, cast(0.00 as decimal(18,2)) AS Juros, cast(0.00 as decimal(18,2)) AS JurosMoeda, cast(0.00 as decimal(18,2)) AS Baixas, cast(0.00 as decimal(18,2)) AS BaixasMoeda, A.RegistroPedido as Pedido " & vbCrLf & _
                "       into #Adto " & vbCrLf & _
                "       FROM vw_Adiantamento A " & vbCrLf & _
                "      WHERE A.Empresa_ID    = '" & CodigoEmpresa & "' " & vbCrLf & _
                "        AND A.EndEmpresa_ID = " & EndEmpresa & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= " AND A.RegistroPedido IN (" & CodigoPedido.Trim & ") " & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                strSQL &= "AND A.Cliente_Id = '" & CodigoCliente & "' " & vbCrLf & _
                          "AND A.EndCliente_Id = " & EndCliente & vbCrLf
            End If
            strSQL &= " ORDER BY A.Movimento " & vbCrLf

            '************************************************************
            '****   Olha o Pedido do titulo de Adiantamento   ***********
            '************************************************************
            strSQL &= "insert into #Adto " & vbCrLf & _
                      "SELECT 2 AS Ordem, " & vbCrLf & _
                      "       case " & vbCrLf & _
                      "     	when isnull(CR.Registro_Id,0) = 0 " & vbCrLf & _
                      "     	  then CP.Movimento " & vbCrLf & _
                      "     	  else CR.Movimento " & vbCrLf & _
                      "       end AS Movimento, " & vbCrLf & _
                      "       case " & vbCrLf & _
                      "     	when isnull(CR.Registro_Id,0) = 0 " & vbCrLf & _
                      "     	  then CP.Baixa " & vbCrLf & _
                      "     	  else CR.Baixa " & vbCrLf & _
                      "       end AS Baixa, " & vbCrLf & _
                      "       case " & vbCrLf & _
                      "     	when isnull(CR.Registro_Id,0) = 0 " & vbCrLf & _
                      "     	  then CP.Prorrogacao " & vbCrLf & _
                      "     	  else CR.Prorrogacao " & vbCrLf & _
                      "       end AS Vencimento, " & vbCrLf & _
                      "       A.Adiantamento_Id, " & vbCrLf & _
                      "       Case " & vbCrLf & _
                      "     	when isnull(CR.Registro_Id,0) = 0 " & vbCrLf & _
                      "     	  then CP.Indexador " & vbCrLf & _
                      "     	  else CR.Indexador " & vbCrLf & _
                      "       end AS Indexador,  " & vbCrLf & _
                      "       case" & vbCrLf & _
                      "         when isnull(CR.Registro_Id,0) = 0" & vbCrLf & _
                      "           then CP.Cliente" & vbCrLf & _
                      "           else CR.Cliente" & vbCrLf & _
                      "       end AS Cliente," & vbCrLf & _
                      "       case" & vbCrLf & _
                      "         when isnull(CR.Registro_Id,0) = 0" & vbCrLf & _
                      "           then CP.EndCliente" & vbCrLf & _
                      "           else CR.EndCliente" & vbCrLf & _
                      "       end AS EndCliente," & vbCrLf & _
                      "       case " & vbCrLf & _
                      "     	when isnull(CR.Registro_Id,0) = 0 " & vbCrLf & _
                      "           then isnull(CP.TaxaAdto,0) " & vbCrLf & _
                      "     	  else isnull(CR.TaxaAdto,0) " & vbCrLf & _
                      "       end AS Taxa,  " & vbCrLf & _
                      "       0 AS ValorOficial, 0 AS ValorMoeda, 0.00 AS Correcao, 'B' as Tipo, ('B-' + convert(varchar,A.Titulo)) AS Titulo, " & vbCrLf & _
                      "       0 AS Juros, 0 AS JurosMoeda, ValorOficial AS Baixas, " & vbCrLf & _
                      "       ValorMoeda AS BaixasMoeda," & vbCrLf & _
                      "       isnull(A.RegistroPedido,0) AS Pedido" & vbCrLf & _
                      "  FROM vw_AdiantamentosXBaixas A " & vbCrLf & _
                      "  LEFT JOIN ContasAPagar CP " & vbCrLf & _
                      "	   on CP.Registro_Id = A.Titulo " & vbCrLf & _
                      "   and CP.Situacao    = 1" & vbCrLf & _
                      "  LEFT JOIN ContasAReceber CR " & vbCrLf & _
                      "	   on CR.Registro_Id = A.Titulo " & vbCrLf & _
                      "   and CR.Situacao    = 1 " & vbCrLf & _
                      " WHERE A.Empresa_Id       ='" & CodigoEmpresa & "'" & vbCrLf & _
                      "   AND A.EndEmpresa_Id    = " & EndEmpresa & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= "   AND isnull(A.RegistroPedido,0) IN (" & CodigoPedido.Trim & ") " & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= "   AND isnull(A.registropedido,isnull(A.RegistroPedidobaixa,0)) = isnull(A.RegistroPedidobaixa,isnull(A.registropedido,0))"

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                strSQL &= "AND A.Cliente_Id    ='" & CodigoCliente & "' " & vbCrLf & _
                          "AND A.EndCliente_Id = " & EndCliente & vbCrLf
            End If
            strSQL &= " ORDER BY A.Adiantamento_Id " & vbCrLf

            '*********************************************************************
            '****   Olha o Pedido do titulo de Baixa do Adiantamento   ***********
            '*********************************************************************
            strSQL &= "insert into #Adto " & vbCrLf & _
                      "SELECT 2 AS Ordem, " & vbCrLf & _
                      "       case " & vbCrLf & _
                      "     	when isnull(CR.Registro_Id,0) = 0 " & vbCrLf & _
                      "     	  then CP.Movimento " & vbCrLf & _
                      "     	  else CR.Movimento " & vbCrLf & _
                      "       end AS Movimento, " & vbCrLf & _
                      "       case " & vbCrLf & _
                      "     	when isnull(CR.Registro_Id,0) = 0 " & vbCrLf & _
                      "     	  then CP.Baixa " & vbCrLf & _
                      "     	  else CR.Baixa " & vbCrLf & _
                      "       end AS Baixa, " & vbCrLf & _
                      "       case " & vbCrLf & _
                      "     	when isnull(CR.Registro_Id,0) = 0 " & vbCrLf & _
                      "     	  then CP.Prorrogacao " & vbCrLf & _
                      "     	  else CR.Prorrogacao " & vbCrLf & _
                      "       end AS Vencimento, " & vbCrLf & _
                      "       A.Adiantamento_Id, " & vbCrLf & _
                      "       Case " & vbCrLf & _
                      "     	when isnull(CR.Registro_Id,0) = 0 " & vbCrLf & _
                      "     	  then CP.Indexador " & vbCrLf & _
                      "     	  else CR.Indexador " & vbCrLf & _
                      "       end AS Indexador,  " & vbCrLf & _
                      "       case" & vbCrLf & _
                      "         when isnull(CR.Registro_Id,0) = 0" & vbCrLf & _
                      "           then CP.Cliente" & vbCrLf & _
                      "           else CR.Cliente" & vbCrLf & _
                      "       end AS Cliente," & vbCrLf & _
                      "       case" & vbCrLf & _
                      "         when isnull(CR.Registro_Id,0) = 0" & vbCrLf & _
                      "           then CP.EndCliente" & vbCrLf & _
                      "           else CR.EndCliente" & vbCrLf & _
                      "       end AS EndCliente," & vbCrLf & _
                      "       case " & vbCrLf & _
                      "     	when isnull(CR.Registro_Id,0) = 0 " & vbCrLf & _
                      "           then isnull(CP.TaxaAdto,0) " & vbCrLf & _
                      "     	  else isnull(CR.TaxaAdto,0) " & vbCrLf & _
                      "       end AS Taxa,  " & vbCrLf & _
                      "       0 AS ValorOficial, 0 AS ValorMoeda, 0.00 AS Correcao, 'B' as Tipo, ('B-' + convert(varchar,A.Titulo)) AS Titulo, " & vbCrLf & _
                      "       0 AS Juros, 0 AS JurosMoeda, ValorOficial AS Baixas, " & vbCrLf & _
                      "       ValorMoeda AS BaixasMoeda," & vbCrLf & _
                      "       isnull(A.RegistroPedido,0) AS Pedido" & vbCrLf & _
                      "  FROM vw_AdiantamentosXBaixas A " & vbCrLf & _
                      "  LEFT JOIN ContasAPagar CP " & vbCrLf & _
                      "	   on CP.Registro_Id = A.Titulo " & vbCrLf & _
                      "   and CP.Situacao    = 1" & vbCrLf & _
                      "  LEFT JOIN ContasAReceber CR " & vbCrLf & _
                      "	   on CR.Registro_Id = A.Titulo " & vbCrLf & _
                      "   and CR.Situacao    = 1 " & vbCrLf & _
                      " WHERE A.Empresa_Id       ='" & CodigoEmpresa & "'" & vbCrLf & _
                      "   AND A.EndEmpresa_Id    = " & EndEmpresa & vbCrLf
            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then
                strSQL &= "   AND (isnull(A.RegistroPedido,0) IN (" & CodigoPedido.Trim & ") or isnull(A.RegistroPedidobaixa,0) IN (" & CodigoPedido.Trim & "))" & vbCrLf & _
                          "   And isnull(A.RegistroPedidobaixa,0) not IN (" & CodigoPedido.Trim & ") " & vbCrLf & _
                          "   AND isnull(A.registropedido,isnull(A.RegistroPedidobaixa,0)) <> isnull(A.RegistroPedidobaixa,isnull(A.registropedido,0))"
            End If

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                strSQL &= "AND A.Cliente_Id    ='" & CodigoCliente & "' " & vbCrLf & _
                          "AND A.EndCliente_Id = " & EndCliente & vbCrLf
            End If
            strSQL &= " ORDER BY A.Adiantamento_Id " & vbCrLf


            '*********************************************************************
            '*********************************************************************
            '*********************************************************************
            strSQL &= "select ROW_NUMBER() OVER(ORDER BY ordem) AS linha, Movimento, Vencimento, Baixa, Adiantamento, Indexador, Cliente, EndCliente, Taxa, ValorOficial, ValorMoeda, Correcao, " & vbCrLf & _
                      "		  Tipo, Titulo, isnull(Juros, 0.00) as Juros, isnull(JurosMoeda, 0.00) as JurosMoeda, Baixas, BaixasMoeda, Pedido" & vbCrLf & _
                      "  into #Adiantamento" & vbCrLf & _
                      "  from #Adto " & vbCrLf & _
                      " order by Ordem; " & vbCrLf

            strSQL &= " select adto.Movimento, adto.Vencimento, adto.Baixa, adto.Adiantamento, adto.Indexador, adto.Cliente, adto.EndCliente, adto.Taxa, adto.ValorOficial, adto.ValorMoeda, adto.Correcao," & vbCrLf & _
                      "        adto.Tipo, adto.Titulo, adto.Juros, adto.JurosMoeda, adto.Baixas, adto.BaixasMoeda , adto.Pedido," & vbCrLf & _
                      "        (select SUM(ValorOficial) + SUM(Juros) - SUM(Baixas) from #Adiantamento where linha <= adto.linha) as SaldoOficial," & vbCrLf & _
                      "        (select SUM(ValorMoeda) + SUM(JurosMoeda) - SUM(BaixasMoeda) from #Adiantamento where linha <= adto.linha) as SaldoMoeda" & vbCrLf & _
                      "   from #Adiantamento as adto" & vbCrLf & _
                      "	inner join Pedidos ped " & vbCrLf & _
                      "			on ped.Empresa_Id     ='" & CodigoEmpresa & "'" & vbCrLf & _
                      "			and ped.EndEmpresa_id = " & EndEmpresa & vbCrLf & _
                      "			and ped.Pedido_id     = adto.Pedido" & vbCrLf

            If SituacaoPedido = "A" Then
                strSQL &= "   WHERE (ped.FiscalAberto = 'true' OR ped.FinanceiroAberto = 'true') " & vbCrLf
            ElseIf SituacaoPedido = "F" Then
                strSQL &= "   WHERE ped.FiscalAberto = 'false' AND ped.FinanceiroAberto = 'false' " & vbCrLf
            End If

            Return New AcessaBanco().ConsultaDataSet(strSQL, "Adiantamentos")

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getDataSetRazao(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal SituacaoPedido As String) As DataSet
        Try
            Dim strSQL As String = ""

            strSQL = "SELECT Razao.Pedido," & vbCrLf & _
                     "       Razao.Movimento_Id as Movimento," & vbCrLf & _
                     "       Razao.Lote_Id as Lote," & vbCrLf & _
                     "	     isnull(Razao.Produto,'') as Produto," & vbCrLf & _
                     "	     Case isnull(PC.Conta_ID,'')" & vbCrLf & _
                     "	       WHEN ''" & vbCrLf & _
                     "		     THEN 'CONTA NAO CADASTRADA, Verifique erro.'" & vbCrLf & _
                     "		     ELSE Razao.Conta_ID" & vbCrLf & _
                     "       END AS Conta," & vbCrLf & _
                     "	     PC.Titulo AS NomeConta," & vbCrLf & _
                     "	     Razao.Historico," & vbCrLf & _
                     "	     Case isnull(PC.Conta_ID,'')" & vbCrLf & _
                     "	       WHEN ''" & vbCrLf & _
                     "		     THEN 0" & vbCrLf & _
                     "		     ELSE Razao.DebitoOficial" & vbCrLf & _
                     "	     END AS DebitoOficial," & vbCrLf & _
                     "	     Case isnull(PC.Conta_ID,'')" & vbCrLf & _
                     "	       WHEN ''" & vbCrLf & _
                     "		      THEN 0" & vbCrLf & _
                     "		      ELSE Razao.CreditoOficial" & vbCrLf & _
                     "	     END AS CreditoOficial," & vbCrLf & _
                     "	     Case isnull(PC.Conta_ID,'')" & vbCrLf & _
                     "	       WHEN ''" & vbCrLf & _
                     "		      THEN 0" & vbCrLf & _
                     "		      ELSE Razao.DebitoMoeda" & vbCrLf & _
                     "	     END AS DebitoMoeda," & vbCrLf & _
                     "	     Case isnull(PC.Conta_ID,'')" & vbCrLf & _
                     "	       WHEN ''" & vbCrLf & _
                     "		      THEN 0" & vbCrLf & _
                     "		      ELSE Razao.CreditoMoeda" & vbCrLf & _
                     "	     END AS CreditoMoeda," & vbCrLf & _
                     "	     Razao.Cliente_Id as Cliente," & vbCrLf & _
                     "	     Razao.EndCliente_Id as EndCliente" & vbCrLf & _
                     "  FROM Razao" & vbCrLf & _
                     "  LEFT JOIN PlanoDeContas PC" & vbCrLf & _
                     "    ON Razao.Conta_Id = PC.Conta_Id" & vbCrLf & _
                     "	LEFT JOIN Pedidos ped " & vbCrLf & _
                     "	  ON ped.Empresa_Id     = Razao.Empresa_Id" & vbCrLf & _
                     "	  AND ped.EndEmpresa_id = Razao.EndEmpresa_Id" & vbCrLf & _
                     "	  AND ped.Pedido_id     = Razao.Pedido" & vbCrLf & _
                     " Where Razao.Empresa_Id    ='" & CodigoEmpresa & "'" & vbCrLf & _
                     "   And Razao.EndEmpresa_Id = " & EndEmpresa & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= "And Razao.Pedido IN (" & CodigoPedido.Trim & ") " & vbCrLf

            If SituacaoPedido = "A" Then
                strSQL &= "   AND (ped.FiscalAberto = 'true' OR ped.FinanceiroAberto = 'true') " & vbCrLf
            ElseIf SituacaoPedido = "F" Then
                strSQL &= "   AND ped.FiscalAberto = 'false' AND ped.FinanceiroAberto = 'false' " & vbCrLf
            End If

            strSQL &= "   AND Razao.Lote_Id not in (9, 10, 11, 21, 70) " & vbCrLf & _
                      " order by razao.Lote_Id, razao.Movimento_Id " & vbCrLf

            Dim dsRazao As DataSet = New AcessaBanco().ConsultaDataSet(strSQL, "Razao")

            Return dsRazao
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Shared Function getDataSetResumoFinanceiro(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer) As DataSet
        Try
            Dim sql As String = ""

            sql = " SELECT '' as UnidadeDeNegocio,                        " & vbCrLf & _
                "          '' as Empresa,                                 " & vbCrLf & _
                "          0 as EndEmpresa,                               " & vbCrLf & _
                "          0 as Pedido,                                   " & vbCrLf & _
                "          '' as Cliente,                                 " & vbCrLf & _
                "          0 as EndCliente,                               " & vbCrLf & _
                "          '' as Cifrao,                                  " & vbCrLf & _
                "          0 as Troca,                                    " & vbCrLf & _
                "          '' as ContaContabilProduto,                    " & vbCrLf & _
                "          '' as ContaContabilAdiantamento,               " & vbCrLf & _
                "          0 as ValorPedido,                              " & vbCrLf & _
                "          0 as ValorTitulosEmPrevisao,                   " & vbCrLf & _
                "          0 as ValorTitulosEmProvisao,                   " & vbCrLf & _
                "          0 as ValorTitulosBaixado,                      " & vbCrLf & _
                "          0 as ValorTitulosCompensado,                   " & vbCrLf & _
                "          0 as ValorAdiantamentoOriginal,                " & vbCrLf & _
                "          0 as ValorAdiantamento,                        " & vbCrLf & _
                "          0 as ValorAdiantamentoCompensado,              " & vbCrLf & _
                "          0 as ValorAdiantamentoPagoDireto,              " & vbCrLf & _
                "          0 as SaldoAdiantamento,                        " & vbCrLf & _
                "          0 as ValorAdiantamentoAmortizado,              " & vbCrLf & _
                "          0 AS ValorPago,                                " & vbCrLf & _
                "   	   0 as SaldoBaixaPedido into #ResumoFinanceiro   " & vbCrLf & _
                "      delete #ResumoFinanceiro                           " & vbCrLf & _
                "      select * from #ResumoFinanceiro                    " & vbCrLf


            Dim ds As DataSet = New AcessaBanco().ConsultaDataSet(sql, "ResumoFinanceiro")

            Return ds
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getDataSetProcuracoes(ByVal Cedente As Boolean, ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer) As DataSet
        Dim strTipo As String = IIf(Cedente, "Cedente", "Cessionario")
        Dim strOutro As String = IIf(Not Cedente, "Cedente", "Cessionario")

        Dim strSQL As String

        If Cedente Then
            strSQL = "SELECT P.PedidoCedente as Pedido," & vbCrLf
        Else
            strSQL = "SELECT " & IIf(CodigoPedido.Length = 0, "0", CodigoPedido) & " as Pedido," & vbCrLf & _
                     "       P.PedidoCedente as PedidoOrigem," & vbCrLf
        End If

        strSQL &= "       P.Procuracao_ID AS Procuracao," & vbCrLf & _
                  "       P.Documento," & vbCrLf & _
                  "       P." & strOutro & " + ' - '  + C.Nome AS Cliente," & vbCrLf & _
                  "       P.Quantidade, " & vbCrLf & _
                  "       ISNULL(sb_REALIZADO.QtdeDev, 0) AS QtdeDev," & vbCrLf & _
                  "       ISNULL(sb_REALIZADO.Quantidade, 0) AS QuantidadeFixado," & vbCrLf & _
                  "       P.Quantidade - ISNULL(sb_REALIZADO.Quantidade, 0) as QtdeFixar" & vbCrLf & _
                  "  FROM Procuracoes P " & _
                  " INNER JOIN Clientes C " & _
                  "    ON C.Cliente_Id  = P." & strOutro & " " & _
                  "   AND C.Endereco_Id = P.End" & strOutro & " " & _
                  "  LEFT OUTER JOIN (SELECT NF.Empresa_Id," & vbCrLf & _
                  "                          NF.EndEmpresa_Id," & vbCrLf & _
                  "                          --NF.Pedido," & vbCrLf & _
                  "                          NF.Procuracao," & vbCrLf & _
                  "						     SUM(case" & vbCrLf & _
                  "						           when so.Devolucao = 'S'" & vbCrLf & _
                  "						            then NFI.QuantidadeFisica" & vbCrLf & _
                  "								    else 0" & vbCrLf & _
                  " 							  end) AS QtdeDev," & vbCrLf & _
                  "             			 SUM(case" & vbCrLf & _
                  "						           when so.Devolucao = 'N'" & vbCrLf & _
                  "						             then NFI.QuantidadeFisica" & vbCrLf & _
                  "								     else 0" & vbCrLf & _
                  "							     end) Quantidade" & vbCrLf & _
                  "                      FROM NotasFiscais NF" & vbCrLf & _
                  "                     INNER JOIN NotasFiscaisXItens NFI " & vbCrLf & _
                  "                        ON NF.Empresa_Id      = NFI.Empresa_Id " & vbCrLf & _
                  "                       AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf & _
                  "                       AND NF.Cliente_Id      = NFI.Cliente_Id " & vbCrLf & _
                  "                       AND NF.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf & _
                  "                       AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
                  "                       AND NF.Serie_Id        = NFI.Serie_Id " & vbCrLf & _
                  "                       AND NF.Nota_Id         = NFI.Nota_Id" & vbCrLf & _
                  "                     Inner Join SubOperacoes SO" & vbCrLf & _
                  "                        ON NFI.Operacao    = so.Operacao_Id" & vbCrLf & _
                  "                       AND NFI.SubOperacao = so.SubOperacoes_Id" & vbCrLf & _
                  "                     GROUP BY NF.Empresa_Id," & vbCrLf & _
                  "                              NF.EndEmpresa_Id," & vbCrLf & _
                  "                              --NF.Pedido," & vbCrLf & _
                  "                              NF.Procuracao " & vbCrLf & _
                  "                    ) AS sb_REALIZADO " & vbCrLf & _
                  "                ON sb_REALIZADO.Empresa_Id    = P.Empresa_Id " & vbCrLf & _
                  "               AND sb_REALIZADO.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
                  "               --AND sb_REALIZADO.Pedido        = P.PedidoCedente " & vbCrLf & _
                  "               AND sb_REALIZADO.Procuracao    = P.Procuracao_ID " & vbCrLf & _
                  "WHERE P.Empresa_Id = '" & CodigoEmpresa & "' " & vbCrLf & _
                  "AND P.EndEmpresa_Id = " & EndEmpresa & " " & vbCrLf

        If Not String.IsNullOrWhiteSpace(CodigoPedido) Then
            If Cedente Then
                strSQL &= " AND P.PedidoCedente IN (" & CodigoPedido.Trim & ") " & vbCrLf
            Else
                strSQL &= " AND P.Procuracao_ID IN (Select procuracao from notasfiscais where pedido = " & CodigoPedido.Trim & vbCrLf
                strSQL &= " AND ISNULL(Situacao, 1) = 1 AND LEN(procuracao) > 0) " & vbCrLf
                strSQL &= " AND p.PedidoCedente <> " & CodigoPedido.Trim
            End If
        End If

        If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
            strSQL &= "AND P." & strOutro & " = '" & CodigoCliente & "' " & vbCrLf & _
                      "AND P.End" & strOutro & " = " & EndCliente.ToString() & vbCrLf
        End If
        strSQL &= "AND P.Situacao = 1"

        Dim dsProcuracoes As DataSet = New AcessaBanco().ConsultaDataSet(strSQL, IIf(Cedente, "ProcuracoesCedente", "Procuracoes"))
        Return dsProcuracoes
    End Function

    Public Shared Function getDataSetNotasFiscais(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal Entrada As Boolean, ByVal Saida As Boolean, ByVal DataLimite As Date?, ByVal SituacaoPedido As String) As DataSet
        Try
            Dim strSQL As String = "SELECT NF.Empresa_Id as Empresa, NF.EndEmpresa_Id as EndEmpresa, NF.Cliente_Id as Cliente, NF.EndCliente_Id as EndCliente, NF.Movimento, NFI.Produto_Id as Produto, " & vbCrLf

            Dim tEmpresas As ClienteXEmpresa = New ClienteXEmpresa(CodigoEmpresa, 0)

            If tEmpresas.UsarRegistroMinAgr Then
                strSQL &= "		 case" & vbCrLf & _
                          "			when len(isnull(Produtos.RegMinAgr,'')) > 0" & vbCrLf & _
                          "				then Produtos.Nome + '(' + Produtos.RegMinAgr  + ')'" & vbCrLf & _
                          "				else Produtos.Nome" & vbCrLf & _
                          "			end AS NomeProduto," & vbCrLf
            ElseIf tEmpresas.UsarDescricaoProduto Then
                strSQL &= "		 case" & vbCrLf & _
                          "			when Produtos.Nome = Produtos.Descricao" & vbCrLf & _
                          "				then Produtos.Nome" & vbCrLf & _
                          "				else Produtos.Nome + '-' + Produtos.Descricao" & vbCrLf & _
                          "			end AS NomeProduto," & vbCrLf
            Else
                strSQL &= "Produtos.Nome AS NomeProduto, " & vbCrLf
            End If

            strSQL &= "       NF.Operacao, NF.SubOperacao, SubOperacoes.Classe, SubOperacoes.Devolucao, NFI.CFOP_Id AS CFOP, NULL AS Reduzido, NF.EntradaSaida_Id as EntradaSaida, " & vbCrLf & _
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
                     "  into #temp" & vbCrLf & _
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
                     "	INNER JOIN Pedidos ped " & vbCrLf & _
                     "			ON ped.Empresa_Id     = NF.Empresa_Id " & vbCrLf & _
                     "			AND ped.EndEmpresa_id = NF.EndEmpresa_Id " & vbCrLf & _
                     "			AND ped.Pedido_id     = NF.Pedido " & vbCrLf & _
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
                     "					  SUM(CASE WHEN ANALISE_ID = 12 THEN Percentual ELSE 0 END) AS GMO " & vbCrLf & _
                     "				 FROM RomaneiosXDescontos " & vbCrLf & _
                     "				GROUP BY Empresa_Id, EndEmpresa_Id, Romaneio_ID" & vbCrLf & _
                     "             ) sb_Descontos " & vbCrLf & _
                     "    ON sb_Descontos.Empresa_Id    = Romaneios.Empresa_Id " & vbCrLf & _
                     "   AND sb_Descontos.EndEmpresa_Id = Romaneios.EndEmpresa_Id " & vbCrLf & _
                     "   AND sb_Descontos.Romaneio_Id   = Romaneios.Romaneio_Id " & vbCrLf & _
                     " WHERE NF.Empresa_Id    ='" & CodigoEmpresa & "'" & vbCrLf & _
                     "   AND NF.EndEmpresa_Id = " & EndEmpresa & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= "AND NF.Pedido  in (" & CodigoPedido.Trim & ")" & vbCrLf

            If SituacaoPedido = "A" Then
                strSQL &= "   AND (ped.FiscalAberto = 'true' OR ped.FinanceiroAberto = 'true') " & vbCrLf
            ElseIf SituacaoPedido = "F" Then
                strSQL &= "   AND ped.FiscalAberto = 'false' AND ped.FinanceiroAberto = 'false' " & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                strSQL &= "AND NF.Cliente_Id = '" & CodigoCliente & "' " & vbCrLf & _
                          "AND NF.EndCliente_Id = " & EndCliente & vbCrLf
            End If

            strSQL &= "   AND (NF.TipoDeDocumento = 1  or isnull(NF.NFG,0) = 1) " & vbCrLf
            strSQL &= "   AND NF.Situacao in (1,4,7) " & vbCrLf

            If Entrada And Not Saida Then
                strSQL &= "   AND (NF.EntradaSaida_id = 'E') " & vbCrLf
            ElseIf Saida And Not Entrada Then
                strSQL &= "   AND (NF.EntradaSaida_id = 'S') " & vbCrLf
            End If

            If DataLimite IsNot Nothing Then strSQL &= "   AND (NF.Movimento <= '" & CDate(DataLimite).ToString("yyyy-MM-dd") & "')" & vbCrLf

            strSQL &= " ORDER BY NF.Movimento " & vbCrLf & _
                      "" & vbCrLf & _
                      " select Empresa, EndEmpresa, Cliente, EndCliente, Pedido," & vbCrLf & _
                      "         isnull(sum(case when Classe = '" & eClassesOperacoes.GLOBAL.ToString & "' then case when Devolucao = 'S' then PesoFiscal * -1 else PesoFiscal end end), 0) as PesoBalancaGlobal," & vbCrLf & _
                      "         isnull(sum(case when Classe = '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then PesoBalanca * -1 else PesoBalanca end end), 0) as PesoBalancaRemessas," & vbCrLf & _
                      "         isnull(sum(case when Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "' and Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then PesoBalanca * -1 else PesoBalanca end end), 0) as PesoBalanca, " & vbCrLf & _
                      "" & vbCrLf & _
                      " 		isnull(sum(case when Classe = '" & eClassesOperacoes.GLOBAL.ToString & "' then case when Devolucao = 'S' then PesoChegada * -1 else PesoChegada end end), 0) as PesoChegadaGlobal," & vbCrLf & _
                      "         isnull(sum(case when Classe = '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then PesoChegada * -1 else PesoChegada end end), 0) as PesoChegadaRemessas," & vbCrLf & _
                      "         isnull(sum(case when Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "' and Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then PesoChegada * -1 else PesoChegada end end), 0) as PesoChegada, " & vbCrLf & _
                      "" & vbCrLf & _
                      " 		isnull(sum(case when Classe = '" & eClassesOperacoes.GLOBAL.ToString & "' then case when Devolucao = 'S' then PesoFiscal * -1 else PesoFiscal end end), 0) as PesoLiquidoGlobal," & vbCrLf & _
                      "         isnull(sum(case when Classe = '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then PesoLiquido * -1 else PesoLiquido end end), 0) as PesoLiquidoRemessas," & vbCrLf & _
                      "         isnull(sum(case when Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "' and Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then PesoLiquido * -1 else PesoLiquido end end), 0) as PesoLiquido, " & vbCrLf & _
                      "" & vbCrLf & _
                      " 		isnull(sum(case when Classe = '" & eClassesOperacoes.GLOBAL.ToString & "' then case when Devolucao = 'S' then PesoFiscal * -1 else PesoFiscal end end), 0) as PesoFiscalGlobal," & vbCrLf & _
                      "         isnull(sum(case when Classe = '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then PesoFiscal * -1 else PesoFiscal end end), 0) as PesoFiscalRemessas," & vbCrLf & _
                      "         isnull(sum(case when Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "' and Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then PesoFiscal * -1 else PesoFiscal end end), 0) as PesoFiscal," & vbCrLf & _
                      "" & vbCrLf & _
                      " 		isnull(sum(case when Classe = '" & eClassesOperacoes.GLOBAL.ToString & "' then case when Devolucao = 'S' then Valor * -1 else Valor end end), 0) as ValorGlobal," & vbCrLf & _
                      "         isnull(sum(case when Classe = '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then Valor * -1 else Valor end end), 0) as ValorRemessas," & vbCrLf & _
                      "         isnull(sum(case when Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "' and Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then Valor * -1 else Valor end end), 0) as Valor," & vbCrLf & _
                      "" & vbCrLf & _
                      " 		isnull(sum(case when Classe = '" & eClassesOperacoes.GLOBAL.ToString & "' then case when Devolucao = 'S' then Liquido * -1 else Liquido end end), 0) as LiquidoGlobal," & vbCrLf & _
                      "         isnull(sum(case when Classe = '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then Liquido * -1 else Liquido end end), 0) as LiquidoRemessas," & vbCrLf & _
                      "         isnull(sum(case when Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "' and Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "' then case when Devolucao = 'S' then Liquido * -1 else Liquido end end), 0) as Liquido" & vbCrLf & _
                      "" & vbCrLf & _
                      "    into #SomaNf" & vbCrLf & _
                      "    from #temp" & vbCrLf & _
                      "   group by Empresa, EndEmpresa, Cliente, EndCliente, Pedido" & vbCrLf & _
                      "" & vbCrLf & _
                      "  select Empresa, EndEmpresa, Cliente, EndCliente, Movimento, Produto, NomeProduto," & vbCrLf & _
                      "         Operacao, SubOperacao, Classe, Devolucao, CFOP, Reduzido, EntradaSaida," & vbCrLf & _
                      "         Serie, Nota, Agrupar, NumeroTicket, NumeroContrato, PesoBalanca, PesoChegada," & vbCrLf & _
                      "         Umidade, Impureza, Avariados, Verde, PH, GMO, PesoLiquido, PesoFiscal, Unitario," & vbCrLf & _
                      "         Valor, Pedido, Finalidade, Liquido, Placa" & vbCrLf & _
                      "    from #temp" & vbCrLf & _
                      "" & vbCrLf & _
                      "  select Empresa, EndEmpresa, Pedido," & vbCrLf & _
                      "         sum(case when PesoBalancaRemessas > 0 then PesoBalancaRemessas" & vbCrLf & _
                      " 	         when PesoBalancaGlobal > 0 then PesoBalancaGlobal" & vbCrLf & _
                      " 		     else PesoBalanca" & vbCrLf & _
                      " 	    end) as PesoBalanca," & vbCrLf & _
                      " 		sum(case when PesoChegadaRemessas > 0 then PesoChegadaRemessas" & vbCrLf & _
                      " 		     when PesoChegadaGlobal > 0 then PesoChegadaGlobal" & vbCrLf & _
                      " 		     else PesoChegada" & vbCrLf & _
                      " 		end) as PesoChegada," & vbCrLf & _
                      " 		sum(case when PesoLiquidoRemessas > 0 then PesoLiquidoRemessas" & vbCrLf & _
                      " 		     when PesoLiquidoGlobal > 0 then PesoLiquidoGlobal" & vbCrLf & _
                      " 			  else PesoLiquido" & vbCrLf & _
                      " 		end) as PesoLiquido," & vbCrLf & _
                      " 		sum(case when PesoFiscalRemessas > 0 then PesoFiscalRemessas" & vbCrLf & _
                      " 		     when PesoFiscalGlobal > 0 then PesoFiscalGlobal" & vbCrLf & _
                      " 		     else PesoFiscal" & vbCrLf & _
                      " 		end) as PesoFiscal," & vbCrLf & _
                      " 		sum(case when ValorRemessas > 0 then ValorRemessas" & vbCrLf & _
                      " 		     when ValorGlobal > 0 then ValorGlobal" & vbCrLf & _
                      " 		     else Valor" & vbCrLf & _
                      " 		end) as Valor," & vbCrLf & _
                      " 		sum(case when LiquidoRemessas > 0 then LiquidoRemessas" & vbCrLf & _
                      " 		     when LiquidoGlobal > 0 then LiquidoGlobal" & vbCrLf & _
                      " 		     else Liquido" & vbCrLf & _
                      " 		end) as Liquido" & vbCrLf & _
                      "     from #SomaNF  group by Empresa, EndEmpresa, Pedido" & vbCrLf

            Dim dsNotasFiscais As DataSet = New AcessaBanco().ConsultaDataSet(strSQL, "NotasFiscais")
            dsNotasFiscais.Tables("NotasFiscais1").TableName = "SomaNotasFiscais"
            Return dsNotasFiscais '#NFD 06
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getDataSetPesagem(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer) As DataSet
        Try
            Dim strSQL As String = " 
              SELECT p.Pesagem_Id AS Pesagem, ISNULL(r.Romaneio_Id,0) AS Romaneio, p.Movimento, p.Placa, " & vbCrLf &
            "   p.PrimeiraPesagem AS Primeira, p.SegundaPesagem AS Segunda, p.BrutoBalanca AS Bruto, " & vbCrLf &
            "   sb_Descontos.Umidade AS Umidade, sb_Descontos.Impureza AS Impureza, sb_Descontos.Avariados AS Avariados," & vbCrLf &
            "   sb_Descontos.Verde AS Verde, sb_Descontos.Quebrados AS Quebrados, " & vbCrLf &
            "   case" & vbCrLf &
            "		when sb_Descontos.INTACTA = 0" & vbCrLf &
            "			then 'NÃO'" & vbCrLf &
            "		when sb_Descontos.INTACTA = 1" & vbCrLf &
            "			then 'NEGATIVO'" & vbCrLf &
            "		when sb_Descontos.INTACTA = 2" & vbCrLf &
            "			then 'POSITIVO'" & vbCrLf &
            "		when sb_Descontos.INTACTA = 3" & vbCrLf &
            "			then 'DECLARADO'" & vbCrLf &
            "		when sb_Descontos.INTACTA = 4" & vbCrLf &
            "			then 'PARTICIPANTE'" & vbCrLf &
            "		when sb_Descontos.INTACTA = 5" & vbCrLf &
            "			then 'ORIGEM PARTICIPANTE'" & vbCrLf &
            "	end AS Intacta," & vbCrLf &
            "   p.Liquido" & vbCrLf &
            " FROM Pesagem p " & vbCrLf &
            "   LEFT JOIN " & vbCrLf &
            "       (SELECT Empresa_Id, EndEmpresa_Id, Pesagem_Id, " & vbCrLf &
            "                   SUM(CASE WHEN ANALISE_ID = 1 THEN Desconto ELSE 0 END) AS Umidade, " & vbCrLf &
            "        	        SUM(CASE WHEN ANALISE_ID = 2 THEN Desconto ELSE 0 END) As Impureza, " & vbCrLf &
            "                   SUM(CASE WHEN ANALISE_ID = 3 THEN Desconto ELSE 0 END) AS Avariados, " & vbCrLf &
            "        	        SUM(CASE WHEN ANALISE_ID = 4 THEN Desconto ELSE 0 END) As Verde, " & vbCrLf &
            "                   SUM(CASE WHEN ANALISE_ID = 5 THEN Desconto ELSE 0 END) AS Quebrados, " & vbCrLf &
            "        	        SUM(CASE WHEN ANALISE_ID = 12 THEN Percentual ELSE 0 END) As Intacta " & vbCrLf &
            "               FROM PesagemXAnalises " & vbCrLf &
            "           GROUP BY Empresa_Id, EndEmpresa_Id, Pesagem_Id " & vbCrLf &
            "       ) sb_Descontos " & vbCrLf &
            "           ON sb_Descontos.Empresa_Id  = P.Empresa_Id " & vbCrLf &
            "           AND sb_Descontos.EndEmpresa_Id = p.EndEmpresa_Id " & vbCrLf &
            "           AND sb_Descontos.Pesagem_Id    = p.Pesagem_Id " & vbCrLf &
            "   LEFT JOIN RomaneiosXPesagens r " & vbCrLf &
            "       ON r.Empresa_Id = p.Empresa_Id " & vbCrLf &
            "       AND r.EndEmpresa_Id = p.EndEmpresa_Id " & vbCrLf &
            "       AND r.Pesagem_Id = p.Pesagem_Id " & vbCrLf &
            "   WHERE p.Empresa_Id    ='" & CodigoEmpresa & "'" & vbCrLf &
            "       AND p.EndEmpresa_Id = " & EndEmpresa & vbCrLf &
            "       AND p.Situacao = 1 "
            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= "And p.Pedido in (" & CodigoPedido.Trim & ")" & vbCrLf

            Dim dsPesagem As DataSet = New AcessaBanco().ConsultaDataSet(strSQL, "Pesagem")

            Return dsPesagem
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function getSumNotasFiscais(ByVal ds As DataSet) As DataSet
        ds.Tables.Add("SomaNotasFiscais")

        ds.Tables("SomaNotasFiscais").Columns.Add("Empresa", GetType(System.String))
        ds.Tables("SomaNotasFiscais").Columns.Add("EndEmpresa", GetType(System.Int32))
        ds.Tables("SomaNotasFiscais").Columns.Add("Cliente", GetType(System.String))
        ds.Tables("SomaNotasFiscais").Columns.Add("EndCliente", GetType(System.Int32))
        ds.Tables("SomaNotasFiscais").Columns.Add("Pedido", GetType(System.Int32))

        ds.Tables("SomaNotasFiscais").Columns.Add("PesoBalanca", GetType(System.Decimal))
        ds.Tables("SomaNotasFiscais").Columns.Add("PesoChegada", GetType(System.Decimal))

        ds.Tables("SomaNotasFiscais").Columns.Add("PesoLiquido", GetType(System.Decimal))
        ds.Tables("SomaNotasFiscais").Columns.Add("PesoFiscal", GetType(System.Decimal))

        ds.Tables("SomaNotasFiscais").Columns.Add("Valor", GetType(System.Decimal))
        ds.Tables("SomaNotasFiscais").Columns.Add("Liquido", GetType(System.Decimal))

        For Each rowPedido As DataRow In ds.Tables("Pedidos").Rows
            Dim rowSumNota As DataRow = ds.Tables("SomaNotasFiscais").NewRow
            rowSumNota("Empresa") = rowPedido("Empresa")
            rowSumNota("EndEmpresa") = rowPedido("EndEmpresa")
            rowSumNota("Cliente") = rowPedido("Cliente")
            rowSumNota("EndCliente") = rowPedido("EndCliente")
            rowSumNota("Pedido") = rowPedido("Pedido")

            Dim classeGlobal As Decimal = IIf(IsNumeric(ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Devolucao = 'N' ")), _
                                                        ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Devolucao = 'N' "), 0) _
                                        - IIf(IsNumeric(ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Devolucao = 'S' ")), _
                                                        ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Devolucao = 'S' "), 0)

        Next

        Return ds
    End Function

    Public Shared Function getDataSetResumoNotasFiscais(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal Entrada As Boolean, ByVal Saida As Boolean, ByVal DataLimite As Date?, ByVal SituacaoPedido As String) As DataSet
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
                                   "	INNER JOIN Pedidos ped " & vbCrLf & _
                                   "			ON ped.Empresa_Id     = NF.Empresa_Id " & vbCrLf & _
                                   "			AND ped.EndEmpresa_id = NF.EndEmpresa_Id " & vbCrLf & _
                                   "			AND ped.Pedido_id     = NF.Pedido " & vbCrLf & _
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
                                   " WHERE (NF.Empresa_Id = '" & CodigoEmpresa & "') " & vbCrLf & _
                                   "   AND (NF.EndEmpresa_Id = " & EndEmpresa & ") " & vbCrLf & _
                                   "    AND (NF.Situacao in (1,4,7)) " & vbCrLf & _
                                   "   AND (NF.TipoDeDocumento = 1  or isnull(NF.NFG,0) = 1 ) " & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= "   AND NF.Pedido IN (" & CodigoPedido.Trim & ") " & vbCrLf

            If SituacaoPedido = "A" Then
                strSQL &= "   AND (ped.FiscalAberto = 'true' OR ped.FinanceiroAberto = 'true') " & vbCrLf
            ElseIf SituacaoPedido = "F" Then
                strSQL &= "   AND ped.FiscalAberto = 'false' AND ped.FinanceiroAberto = 'false' " & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                strSQL &= "AND NF.Cliente_Id = '" & CodigoCliente & "' " & vbCrLf & _
                          "AND NF.EndCliente_Id = " & EndCliente.ToString() & vbCrLf
            End If

            If Entrada And Not Saida Then
                strSQL &= "   AND (NF.EntradaSaida_id = 'E') " & vbCrLf
            ElseIf Saida And Not Entrada Then
                strSQL &= "   AND (NF.EntradaSaida_id = 'S') " & vbCrLf
            End If

            If DataLimite IsNot Nothing Then strSQL &= "   AND (NF.Movimento <= '" & CDate(DataLimite).ToString("yyyy-MM-dd") & "')" & vbCrLf

            strSQL &= "   group by prd.Agrupar, NF.Operacao, NF.SubOperacao, NF.Cliente_Id, NF.EndCliente_Id, NF.EntradaSaida_Id, SO.Descricao, NF.Pedido " & vbCrLf & _
                      "   ORDER BY 1"

            Dim dsResumoNotasFiscais As DataSet = New AcessaBanco().ConsultaDataSet(strSQL, "ResumoNotasFiscais")
            Return dsResumoNotasFiscais
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getDataSetResumoNotasFiscaisPrd(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal Entrada As Boolean, ByVal Saida As Boolean, ByVal DataLimite As Date?, ByVal SituacaoPedido As String) As DataSet
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
                     "	INNER JOIN Pedidos ped " & vbCrLf & _
                     "			ON ped.Empresa_Id     = NF.Empresa_Id " & vbCrLf & _
                     "			AND ped.EndEmpresa_id = NF.EndEmpresa_Id " & vbCrLf & _
                     "			AND ped.Pedido_id     = NF.Pedido " & vbCrLf & _
                     "  LEFT JOIN Romaneios " & vbCrLf & _
                     "    ON Romaneios.Empresa_Id    = NFR.Empresa_Id " & vbCrLf & _
                     "   AND Romaneios.EndEmpresa_Id = NFR.EndEmpresa_Id " & vbCrLf & _
                     "   AND Romaneios.Romaneio_Id   = NFR.Romaneio_Id " & vbCrLf & _
                     " WHERE (NF.Empresa_Id    = '" & CodigoEmpresa & "') " & vbCrLf & _
                     "   AND (NF.EndEmpresa_Id = " & EndEmpresa & ") " & vbCrLf & _
                     "   AND (NF.Situacao in (1,4,7)) " & vbCrLf & _
                     "   AND (NF.TipoDeDocumento = 1  or isnull(NF.NFG,0) = 1 ) " & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then strSQL &= "   AND NF.Pedido IN (" & CodigoPedido.Trim & ") " & vbCrLf

            If SituacaoPedido = "A" Then
                strSQL &= "   AND (ped.FiscalAberto = 'true' OR ped.FinanceiroAberto = 'true') " & vbCrLf
            ElseIf SituacaoPedido = "F" Then
                strSQL &= "   AND ped.FiscalAberto = 'false' AND ped.FinanceiroAberto = 'false' " & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                strSQL &= "AND NF.Cliente_Id    = '" & CodigoCliente & "'" & vbCrLf & _
                          "AND NF.EndCliente_Id = " & EndCliente.ToString() & vbCrLf
            End If

            If DataLimite IsNot Nothing Then
                strSQL &= "   AND (NF.Movimento <= '" & CDate(DataLimite).ToString("yyyy-MM-dd") & "')" & vbCrLf
            End If

            strSQL &= "   group by NF.Cliente_Id, NF.EndCliente_Id, NFI.Produto_Id, PRD.Nome, NF.Pedido  " & vbCrLf & _
                      "   ORDER BY 1"

            Dim dsResumoNotasFiscaisPrd As DataSet = New AcessaBanco().ConsultaDataSet(strSQL, "ResumoNotasFiscaisPrd")
            Return dsResumoNotasFiscaisPrd
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Shared Function getDataSetFretes(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal Entrada As Boolean, ByVal Saida As Boolean, ByVal DataLimite As Date?)
        Try
            Dim sql As String = " SELECT F.NotaEmpresa, F.NotaEndEmpresa, F.Pedido," & vbCrLf & _
                                "        F.ClienteCodigo, f.CteMovimento," & vbCrLf & _
                                "        f.CteCliente, f.CteEndCliente, NotaTransportador Transportador," & vbCrLf & _
                                "        CASE " & vbCrLf & _
                                "            WHEN (CteValor + CtePedagioValor) > 0 " & vbCrLf & _
                                "                THEN CAST(CteNota as varchar) + '-' + CAST(CteSerie as varchar)" & vbCrLf & _
                                "            WHEN RPAValor > 0 " & vbCrLf & _
                                "                THEN CAST(RPANota AS VARCHAR) + '-' + RPASerie " & vbCrLf & _
                                "            WHEN NotaDeTrocaVendaPedido > 0 " & vbCrLf & _
                                "                THEN 'Troca: ' + CAST(NotaDeTrocaVendaPedido AS VARCHAR)" & vbCrLf & _
                                "                ELSE ' '" & vbCrLf & _
                                "        END AS CTRC," & vbCrLf & _
                                "       CONVERT(VARCHAR,F.Nota)+'-'+F.NotaSerie  AS NotaSerie," & vbCrLf & _
                                "		F.PedidoFreteOrcadoUnitario UnitFreteOrcado," & vbCrLf & _
                                "		NotaDeTrocaCompraValor AS UnitFreteDif, --FRETE COMPRA" & vbCrLf & _
                                "        CASE " & vbCrLf & _
                                "            WHEN (CteValor + CtePedagioValor+ RpaValor) > 0 " & vbCrLf & _
                                "                THEN(CteValor + CtePedagioValor + RpaValor) * (CteUnitario + RPAUnitario) / (CteValor + CtePedagioValor+ RpaValor) " & vbCrLf & _
                                "                ELSE 0" & vbCrLf & _
                                "        END AS UnitarioFrete, --Unitario do Frete Realizado " & vbCrLf & _
                                "        CASE " & vbCrLf & _
                                "            WHEN (CteValor + CtePedagioValor + RPAValor)> 0" & vbCrLf & _
                                "               THEN PedidoFreteOrcadoUnitario  + NotaDeTrocaCompraValor - ((CteValor + CtePedagioValor + RPAValor) * (CteUnitario+RPAUnitario))  /  (CteValor + CtePedagioValor + RPAValor)" & vbCrLf & _
                                "               ELSE 0" & vbCrLf & _
                                "        END AS DiferencaUnitFrete, --DIF. UNIT." & vbCrLf & _
                                "        CtePedagioValor AS TarifaPedagio," & vbCrLf & _
                                "        F.LaudoQuebraSobraValor AS QuebraSobraValor," & vbCrLf & _
                                "        CteValor + CtePedagioValor + RPAValor as ValorFrete," & vbCrLf & _
                                "       CASE" & vbCrLf & _
                                "             WHEN (CteValor + CtePedagioValor + RPAValor)> 0" & vbCrLf & _
                                "                THEN (F.PedidoFreteOrcadoUnitario  " & vbCrLf & _
                                "                       + F.NotaDeTrocaCompraValor " & vbCrLf & _
                                "                       - ((CteValor + CtePedagioValor + RPAValor)) * (CteUnitario+RPAUnitario)  /  (CteValor + CtePedagioValor + RPAValor))" & vbCrLf & _
                                "                     * LaudoPesoChegada /1000" & vbCrLf & _
                                "                ELSE 0" & vbCrLf & _
                                "        END VlrDiferenca," & vbCrLf & _
                                "        NotaLocalEmbarque LocalEmbarque" & vbCrLf & _
                                "   FROM vw_RelatorioDeFrete AS F" & vbCrLf & _
                                "  WHERE CteNota > 0 " & vbCrLf & _
                                "    AND F.NotaEmpresa = '" & CodigoEmpresa & "'" & vbCrLf & _
                                "    AND F.NotaEndEmpresa = " & EndEmpresa & vbCrLf & _
                                "    AND ((F.NotaEntradasaida = 'S' and F.NotaCIFFOB  = 'CIF') OR (F.NotaEntradasaida = 'E' and F.NotaCIFFOB = 'FOB')) " & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoPedido) Then sql &= "   AND F.Pedido IN (" & CodigoPedido.Trim & ") " & vbCrLf

            If Not String.IsNullOrWhiteSpace(CodigoCliente) Then
                sql &= "AND F.ClienteCodigo = '" & CodigoCliente & "'" & vbCrLf & _
                       "AND F.ClienteEndereco = " & EndCliente.ToString() & vbCrLf
            End If

            If DataLimite IsNot Nothing Then sql &= "   AND (F.NotaMovimento <= '" & CDate(DataLimite).ToString("yyyy-MM-dd") & "')" & vbCrLf

            sql &= "  ORDER BY CASE WHEN F.NotaTransportador = 'NENHUM' THEN 'ZZ' ELSE F.NotaLocalEmbarque END, F.NotaMovimento                          " & vbCrLf


            Dim dsFretes As DataSet = New AcessaBanco().ConsultaDataSet(sql, "Fretes")

            Return dsFretes
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Shared Function getDataSetResumoPedido(ByRef ds As DataSet, ByVal CodigoPedido As String) As DataSet
        Try
            ds.Merge(createTableResumo())

            For Each rowPedido As DataRow In ds.Tables("Pedidos").Rows

                If rowPedido("Pedido") = CodigoPedido Then
                    Dim SaldoContratadoNotas As Decimal
                    Dim SaldoContratado As Decimal
                    Dim qtdeEntregue As Integer
                    Dim qtdeDevolvida As Integer

                    Dim rowResumo As DataRow = ds.Tables("ResumoPedido").NewRow
                    rowResumo("Pedido") = rowPedido("Pedido")

                    If rowPedido("Classe").Equals("AFIXAR") OrElse rowPedido("Classe").Equals("DEPOSITOS") OrElse rowPedido("Classe").Equals("MUTUO") Then
                        SaldoContratadoNotas = IIf(IsNumeric(ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Devolucao = 'N' ")),
                                                            ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Devolucao = 'N' "), 0) _
                                            - IIf(IsNumeric(ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Devolucao = 'S' ")),
                                                            ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Devolucao = 'S' "), 0)
                    End If

                    SaldoContratado = IIf(IsNumeric(ds.Tables("Contratos").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido") & " And TipoDeLancamento <> 'E'")),
                                                    ds.Tables("Contratos").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido") & " And TipoDeLancamento <> 'E'"), 0) _
                                    - IIf(IsNumeric(ds.Tables("Contratos").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido") & " And TipoDeLancamento = 'E'")),
                                                    ds.Tables("Contratos").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido") & " And TipoDeLancamento = 'E'"), 0)

                    rowResumo("QtdeContratada") = IIf(SaldoContratado > SaldoContratadoNotas, SaldoContratado, SaldoContratadoNotas)

                    If rowPedido("Classe").Equals("AFIXAR") Then
                        rowResumo("QtdeFixada") = IIf(IsNumeric(ds.Tables("Fixacoes").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido"))),
                                                                ds.Tables("Fixacoes").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido")), 0)
                    Else
                        rowResumo("QtdeFixada") = rowResumo("QtdeContratada")
                    End If

                    rowResumo("QtdeCessionario") = IIf(IsNumeric(ds.Tables("Procuracoes").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido"))),
                                                       ds.Tables("Procuracoes").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido")), 0)

                    rowResumo("QtdeCedente") = IIf(IsNumeric(ds.Tables("ProcuracoesCedente").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido"))),
                                                   ds.Tables("ProcuracoesCedente").Compute("Sum(Quantidade)", "Pedido = " & rowPedido("Pedido")), 0)

                    rowResumo("SaldoProcuracao") = rowResumo("QtdeCessionario") - rowResumo("QtdeCedente")



                    qtdeEntregue = IIf(IsNumeric(ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "' And Devolucao = 'N' ")),
                                      ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "' And Devolucao = 'N' "), 0)

                    qtdeDevolvida = IIf(IsNumeric(ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "' And Devolucao = 'S' ")),
                                    ds.Tables("NotasFiscais").Compute("Sum(PesoFiscal)", "Pedido = " & rowPedido("Pedido") & " And Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "' And Devolucao = 'S' "), 0)



                    rowResumo("QtdeEntregue") = qtdeEntregue
                    rowResumo("QtdeDevolvida") = qtdeDevolvida

                    If rowPedido("Classe").Equals("AFIXAR") Then
                        rowResumo("SaldoAFixar") = qtdeEntregue - qtdeDevolvida - rowResumo("QtdeFixada")
                    Else
                        rowResumo("SaldoAFixar") = rowResumo("QtdeContratada") - rowResumo("QtdeFixada")
                    End If

                    If rowPedido("Classe") = "DEPOSITOS" OrElse rowPedido("Classe") = "MUTUO" Then
                        rowResumo("SaldoAEntregar") = rowResumo("QtdeEntregue") - rowResumo("QtdeDevolvida")
                    Else
                        rowResumo("SaldoAEntregar") = rowResumo("QtdeContratada") - rowResumo("QtdeEntregue") + rowResumo("QtdeDevolvida")
                    End If

                    Dim valorLiquidoNota As Decimal = 0

                    If rowPedido("Classe").Equals("AFIXAR") OrElse rowPedido("Classe").Equals("DEPOSITOS") OrElse rowPedido("Classe").Equals("MUTUO") Then
                        valorLiquidoNota = IIf(IsNumeric(ds.Tables("Fixacoes").Compute("Sum(Liquido)", "Pedido = " & rowPedido("Pedido"))),
                                                         ds.Tables("Fixacoes").Compute("Sum(Liquido)", "Pedido = " & rowPedido("Pedido")), 0)
                    Else
                        Dim vl = ds.Tables("NotasFiscais").Compute("Sum(Liquido)", "Pedido = " & rowPedido("Pedido") & " And Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "' And Devolucao = 'N'")
                        Dim vld = ds.Tables("NotasFiscais").Compute("Sum(Liquido)", "Pedido = " & rowPedido("Pedido") & " And Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "' And Devolucao = 'S'")

                        valorLiquidoNota = CDec(IIf(IsNumeric(vl), vl, 0)) - CDec(IIf(IsNumeric(vld), vld, 0))
                    End If

                    If rowPedido("PrecoFixo") = "S" Then
                        rowResumo("ValorFixado") = IIf(IsNumeric(ds.Tables("Contratos").Compute("Sum(TotalOficial)", "Pedido = " & rowPedido("Pedido") & " And TipoDeLancamento = 'N' ")),
                                                       ds.Tables("Contratos").Compute("Sum(TotalOficial)", "Pedido = " & rowPedido("Pedido") & " And TipoDeLancamento = 'N' "), 0) _
                                        + IIf(IsNumeric(ds.Tables("Contratos").Compute("Sum(TotalOficial)", "Pedido = " & rowPedido("Pedido") & " And TipoDeLancamento = 'C' ")),
                                                       ds.Tables("Contratos").Compute("Sum(TotalOficial)", "Pedido = " & rowPedido("Pedido") & " And TipoDeLancamento = 'C' "), 0) _
                                        - IIf(IsNumeric(ds.Tables("Contratos").Compute("Sum(TotalOficial)", "Pedido = " & rowPedido("Pedido") & " And TipoDeLancamento = 'E' ")),
                                                       ds.Tables("Contratos").Compute("Sum(TotalOficial)", "Pedido = " & rowPedido("Pedido") & " And TipoDeLancamento = 'E' "), 0)
                    Else
                        rowResumo("ValorFixado") = IIf(IsNumeric(ds.Tables("Financeiro").Compute("SUM(ValorDoDocumento)", "Pedido = " & rowPedido("Pedido") & " And Provisao <> 4 And Provisao <> 5 and (Tipo = 'P' OR Tipo = 'P|BA')")),
                                                           ds.Tables("Financeiro").Compute("SUM(ValorDoDocumento)", "Pedido = " & rowPedido("Pedido") & " And Provisao <> 4 And Provisao <> 5 and (Tipo = 'P' OR Tipo = 'P|BA')"), 0) _
                                           - IIf(IsNumeric(ds.Tables("Financeiro").Compute("SUM(ValorDoDocumento)", "Pedido = " & rowPedido("Pedido") & " And Provisao <> 4 And Provisao <> 5 and Tipo = 'R'")),
                                                           ds.Tables("Financeiro").Compute("SUM(ValorDoDocumento)", "Pedido = " & rowPedido("Pedido") & " And Provisao <> 4 And Provisao <> 5 and Tipo = 'R'"), 0) _
                                           - IIf(IsNumeric(ds.Tables("Financeiro").Compute("SUM(ValorDoDocumento)", "Pedido = " & rowPedido("Pedido") & " And Provisao <> 4 And Provisao <> 5 and Tipo = 'RA'")),
                                                           ds.Tables("Financeiro").Compute("SUM(ValorDoDocumento)", "Pedido = " & rowPedido("Pedido") & " And Provisao <> 4 And Provisao <> 5 and Tipo = 'RA'"), 0)
                    End If

                    If rowResumo("ValorFixado") < 0 Then rowResumo("ValorFixado") = rowResumo("ValorFixado") * (-1)

                    '**********************************************************************************************************************************************
                    '**********************************************************************************************************************************************
                    '**********************************************************************************************************************************************

                    Dim vlrpago = ds.Tables("Financeiro").Compute("SUM(ValorDoDocumento)", "Pedido = " & rowPedido("Pedido") & " And Provisao = 1 and (Tipo = 'P' OR Tipo = 'P|BA')")
                    Dim vlrrecebido = ds.Tables("Financeiro").Compute("SUM(ValorDoDocumento)", "Pedido = " & rowPedido("Pedido") & " And Provisao = 1 and (Tipo = 'R' OR Tipo = 'R|BA')")
                    Dim vlrrecadiant = ds.Tables("Financeiro").Compute("SUM(ValorDoDocumento)", "Pedido = " & rowPedido("Pedido") & " And Provisao = 1 and Tipo = 'RA'")

                    If rowPedido("EntradaSaida") = "E" Then
                        rowResumo("ValorPago") = CDec(IIf(IsNumeric(vlrpago), vlrpago, 0)) - CDec(IIf(IsNumeric(vlrrecebido), vlrrecebido, 0)) - CDec(IIf(IsNumeric(vlrrecadiant), vlrrecadiant, 0))
                    Else
                        rowResumo("ValorPago") = CDec(IIf(IsNumeric(vlrrecebido), vlrrecebido, 0)) - CDec(IIf(IsNumeric(vlrpago), vlrpago, 0)) - CDec(IIf(IsNumeric(vlrrecadiant), vlrrecadiant, 0))
                    End If

                    rowResumo("SaldoFinanceiro") = valorLiquidoNota - rowResumo("ValorPago")

                    ds.Tables("ResumoPedido").Rows.Add(rowResumo)
                End If
            Next

            Return ds
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Shared Function getDataSetTotalizadorImpostos(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoPedido As String) As DataSet
        Try
            Dim sql As String = "
                    DECLARE @Pedido INT = " & CodigoPedido & ";
                    DECLARE @Empresa VARCHAR(50) = '" & CodigoEmpresa & "';
                    DECLARE @EndEmpresa INT = " & EndEmpresa & ";

                    WITH TotalNotasTb AS (
                     SELECT
                      NFxI.Pedido AS Pedido,
                      NFxE.Encargo_Id,
                      SUM(NFxE.Valor) AS TotalNotas
                     FROM NotasFiscaisXItens NFxI
                      INNER JOIN NotasFiscaisXEncargos NFxE
                       ON NFxI.Empresa_Id = NFxE.Empresa_Id
                        AND NFxI.EndEmpresa_Id = NFxE.EndEmpresa_Id
                        AND NFxI.Cliente_Id = NFxE.Cliente_Id
                        AND NFxI.EndCliente_Id = NFxE.EndCliente_Id
                        AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id
                        AND NFxI.Serie_Id = NFxE.Serie_Id
                        AND NFxI.Nota_Id = NFxE.Nota_Id
                        AND NFxI.Sequencia_id = NFxE.Sequencia_id
                      INNER JOIN SubOperacoes SubOp
                       ON SubOp.Operacao_Id = NFxI.Operacao
                        AND SubOp.SubOperacoes_Id = NFxI.SubOperacao
                     WHERE NFxI.Pedido = @Pedido
                      AND NFxI.Empresa_Id = @Empresa
                      AND NFxI.EndEmpresa_Id = @EndEmpresa
                      AND SubOp.Devolucao != 'S'
                     GROUP BY NFxE.Encargo_Id, NFxI.Pedido
                    ), TotalDevolucoesTb AS (
                     SELECT
                      NFxE.Encargo_Id,
                      SUM(NFxE.Valor) AS TotalDevolucoes
                     FROM NotasFiscaisXItens NFxI
                      INNER JOIN NotasFiscaisXEncargos NFxE
                       ON NFxI.Empresa_Id = NFxE.Empresa_Id
                        AND NFxI.EndEmpresa_Id = NFxE.EndEmpresa_Id
                        AND NFxI.Cliente_Id = NFxE.Cliente_Id
                        AND NFxI.EndCliente_Id = NFxE.EndCliente_Id
                        AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id
                        AND NFxI.Serie_Id = NFxE.Serie_Id
                        AND NFxI.Nota_Id = NFxE.Nota_Id
                        AND NFxI.Sequencia_id = NFxE.Sequencia_id
                      INNER JOIN SubOperacoes SubOp
                       ON SubOp.Operacao_Id = NFxI.Operacao
                        AND SubOp.SubOperacoes_Id = NFxI.SubOperacao
                     WHERE NFxI.Pedido = @Pedido
                      AND NFxI.Empresa_Id = @Empresa
                      AND NFxI.EndEmpresa_Id = @EndEmpresa
                      AND SubOp.Devolucao != 'N'
                     GROUP BY NFxE.Encargo_Id
                    ), Sinal AS (
					 SELECT
                      Oee.Encargo_Id,
                      Oee.Sinal
                     FROM OperacaoXEstadoXEncargo Oee
					  INNER JOIN NotasFiscaisXItens NFxI
					   ON Oee.Codigo_Id = NFxI.OperacaoXEstado               
                     WHERE NFxI.Pedido = @Pedido
                      AND NFxI.Empresa_Id = @Empresa
                      AND NFxI.EndEmpresa_Id = @EndEmpresa
					 GROUP BY Oee.Encargo_Id, Oee.Sinal) 

                    SELECT
                     tn.Pedido AS Pedido,
                     COALESCE(tn.Encargo_Id, td.Encargo_Id) AS Encargo_Id,
					 COALESCE(sn.Sinal, 'NOT') AS Sinal,
                     COALESCE(tn.TotalNotas, 0) AS TotalNotas,
                     COALESCE(td.TotalDevolucoes, 0) AS TotalDevolucoes,
                     ISNULL((TotalNotas - TotalDevolucoes), TotalNotas) AS Saldo
                    FROM TotalNotasTb tn
                     FULL OUTER JOIN TotalDevolucoesTb td
                      ON tn.Encargo_Id = td.Encargo_Id
					 FULL OUTER JOIN Sinal sn
					  ON sn.Encargo_Id = tn.Encargo_Id
                    ORDER BY
                     CASE
					  WHEN COALESCE(tn.Encargo_Id, td.Encargo_Id) = 'PRODUTO' THEN 1
					  WHEN COALESCE(tn.Encargo_Id, td.Encargo_Id) = 'LIQUIDO' THEN 3
					  ELSE 2
					 END,
					  COALESCE(tn.Encargo_Id, td.Encargo_Id); "


            Dim dsTotalizador As DataSet = New AcessaBanco().ConsultaDataSet(sql, "Totalizador")
            Return dsTotalizador

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Shared Function createTableResumo() As DataTable
        Dim dt As New DataTable
        dt.TableName = "ResumoPedido"
        dt.Columns.Add("Pedido", GetType(System.Int32))
        dt.Columns.Add("QtdeContratada", GetType(System.Int32))
        dt.Columns.Add("QtdeFixada", GetType(System.Int32))
        dt.Columns.Add("SaldoAFixar", GetType(System.Int32))

        dt.Columns.Add("QtdeEntregue", GetType(System.Int32))
        dt.Columns.Add("QtdeDevolvida", GetType(System.Int32))
        dt.Columns.Add("SaldoAEntregar", GetType(System.Int32))

        dt.Columns.Add("QtdeCessionario", GetType(System.Int32))
        dt.Columns.Add("QtdeCedente", GetType(System.Int32))
        dt.Columns.Add("SaldoProcuracao", GetType(System.Int32))

        dt.Columns.Add("ValorFixado", GetType(System.Decimal))
        dt.Columns.Add("ValorPago", GetType(System.Decimal))
        dt.Columns.Add("SaldoFinanceiro", GetType(System.Decimal))

        Return dt
    End Function

End Class
