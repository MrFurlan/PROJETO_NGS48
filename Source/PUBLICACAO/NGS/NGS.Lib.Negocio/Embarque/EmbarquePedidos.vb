Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListEmbarquePedido
    Inherits List(Of EmbarquePedido)

    Public Sub New(ByVal pedido As String, ByVal safra As String, Optional ByVal empresa As String = "", Optional ByVal enderecoEmpresa As String = "", Optional ByVal cliente As String = "", Optional ByVal enderecoCliente As String = "", Optional ByVal entradaSaida As String = "", Optional ByVal RetEntrega As Boolean = False)
        Dim sql As String = ""
        sql &= "select P.Empresa_Id as Empresa_Id,                                                                                                 " & vbCrLf & _
               "       P.EndEmpresa_Id as EndEmpresa_Id,                                                                                           " & vbCrLf & _
               "       P.Cliente as Cliente,                                                                                                       " & vbCrLf & _
               "       P.EndCliente as EndCliente,                                                                                                 " & vbCrLf & _
               "       C.Nome as NomeCliente,                                                                                                      " & vbCrLf & _
               "       P.Safra as Safra,                                                                                                           " & vbCrLf & _
               "       P.Pedido_Id as Pedido_Id,                                                                                                   " & vbCrLf & _
               "       convert(varchar,P.DataEntrega, 103) as DataEntrega,                                                                         " & vbCrLf & _
               "       P.EmbarqueAtivo as EmbarqueAtivo,                                                                                           " & vbCrLf & _
               "       convert(nvarchar(2),SO.Operacao_id) + '-' + convert(nvarchar(2),SO.suboperacoes_id) + ' : ' + SO.Descricao as DescOperacao, " & vbCrLf

        If RetEntrega Then
            sql &= "       isnull(Cast((AEE.Entrega_Id + '-' + Cast(AEE.EndEntrega_Id as nvarchar)) as nvarchar),'') as Entrega," & vbCrLf & _
                   "       isnull(AEE.Entrega_Id, '') as Entrega_Id,                                                            " & vbCrLf & _
                   "       isnull(AEE.EndEntrega_Id,0) as EndEntrega_Id,                                                        " & vbCrLf
        Else
            sql &= "      '' as Entrega,                                            " & vbCrLf & _
                   "      '' as Entrega_Id,                                         " & vbCrLf & _
                   "      0 as EndEntrega_Id,                                       " & vbCrLf
        End If

        sql &= "       SO.EntradaSaida,                                             " & vbCrLf & _
               "       isnull(P.CidadeEntrega,'') as CidadeEntrega,                 " & vbCrLf & _
               "       ISNULL(P.EstadoEntrega,'') as EstadoEntrega,                 " & vbCrLf & _
               "       Pxi.Produto_Id,                                              " & vbCrLf & _
               "       Prd.Nome as NomeProduto,                                     " & vbCrLf & _
               "       PxI.Qtde as QtdePedido,                                      " & vbCrLf & _
               "       isnull(AutEmb.QtdeNormal,0) as QtdeNormal,                   " & vbCrLf & _
               "       isnull(AutEmb.QtdeComplemento,0) as QtdeComplemento,         " & vbCrLf & _
               "       isnull(AutEmb.QtdeEstorno,0) as QtdeEstorno,                 " & vbCrLf & _
               "       isnull(AutEmb.QtdeAutorizado,0) as QtdeAutorizadoEmbarque,   " & vbCrLf & _
               "       isnull(Case                                                  " & vbCrLf & _
               "				 when SO.EntradaSaida = 'S'                         " & vbCrLf & _
               "				   then Rom.PesoSaida                               " & vbCrLf & _
               "				   else Rom.PesoEntrada                             " & vbCrLf & _
               "			   end,0) Embarcado,                                    " & vbCrLf & _
               "       isnull(Case                                                  " & vbCrLf & _
               "				 when SO.EntradaSaida = 'S'                         " & vbCrLf & _
               "				   then Rom.PesoEntrada                             " & vbCrLf & _
               "				   else Rom.PesoSaida                               " & vbCrLf & _
               "			   end,0) Devolvido                                     " & vbCrLf & _
               "  FROM Pedidos P                                                    " & vbCrLf & _
               " INNER JOIN (Select Empresa_Id,                                     " & vbCrLf & _
               "                    EndEmpresa_Id,                                  " & vbCrLf & _
               "                    Pedido_Id,                                      " & vbCrLf & _
               "                    Produto_Id,                                     " & vbCrLf & _
               "                    SUM(Case                                        " & vbCrLf & _
               "                          when TipoDeLancamento = 'E'               " & vbCrLf & _
               "                            Then Quantidade * - 1                   " & vbCrLf & _
               "                            else Quantidade                         " & vbCrLf & _
               "                        end) Qtde,                                  " & vbCrLf & _
               "                    SUM(Case                                        " & vbCrLf & _
               "                          when TipoDeLancamento = 'E'               " & vbCrLf & _
               "                            Then isnull(QuantidadeEstimada,0) * - 1 " & vbCrLf & _
               "                            else isnull(QuantidadeEstimada,0)       " & vbCrLf & _
               "                        end) QtdeEstimada                           " & vbCrLf & _
               "               from PedidoXItemXLancamento                          " & vbCrLf & _
               "              Group by Empresa_Id,                                  " & vbCrLf & _
               "                       EndEmpresa_Id,                               " & vbCrLf & _
               "                       Pedido_Id,                                   " & vbCrLf & _
               "                       Produto_Id                                   " & vbCrLf & _
               "             ) PxI                                                  " & vbCrLf & _
               "    on P.Empresa_Id    = PxI.Empresa_Id                             " & vbCrLf & _
               "   and P.EndEmpresa_Id = PxI.EndEmpresa_Id                          " & vbCrLf & _
               "   and P.Pedido_Id     = PxI.Pedido_Id                              " & vbCrLf & _
               " INNER JOIN SubOperacoes  SO                                        " & vbCrLf & _
               "    on P.Operacao    = SO.Operacao_Id                               " & vbCrLf & _
               "   and p.SubOperacao = SO.SubOperacoes_Id                           " & vbCrLf

        If RetEntrega Then
            sql &= "  INNER JOIN AutEmbarqueXEntrega AEE on (AEE.Empresa_Id = P.Empresa_Id AND AEE.EndEmpresa_Id = P.EndEmpresa_Id AND AEE.Pedido_Id = P.Pedido_Id)  " & vbCrLf
        End If

        sql &= "LEFT JOIN (Select ae.Empresa_Id,                                    " & vbCrLf & _
               "                    ae.EndEmpresa_Id,                               " & vbCrLf & _
               "                    ae.Pedido_Id,                                   " & vbCrLf & _
               "                    ae.Produto_Id,                                  " & vbCrLf

        If RetEntrega Then
            sql &= "                    ae.Entrega_Id,                              " & vbCrLf & _
                   "                    ae.EndEntrega_Id,                           " & vbCrLf
        End If

        sql &= "                    SUM(Case                                        " & vbCrLf & _
               "                          when ae.TipoDeLancamento = 'N'            " & vbCrLf & _
               "                            Then ae.Quantidade                      " & vbCrLf & _
               "                            else 0                                  " & vbCrLf & _
               "                        end) QtdeNormal,                            " & vbCrLf & _
               "                    SUM(Case                                        " & vbCrLf & _
               "                          when ae.TipoDeLancamento = 'C'            " & vbCrLf & _
               "                            Then ae.Quantidade                      " & vbCrLf & _
               "                            else 0                                  " & vbCrLf & _
               "                        end) QtdeComplemento,                       " & vbCrLf & _
               "                    SUM(Case                                        " & vbCrLf & _
               "                          when ae.TipoDeLancamento = 'E'            " & vbCrLf & _
               "                            Then ae.Quantidade                      " & vbCrLf & _
               "                            else 0                                  " & vbCrLf & _
               "                        end) QtdeEstorno,                           " & vbCrLf & _
               "                    SUM(Case                                        " & vbCrLf & _
               "                          when ae.TipoDeLancamento = 'E'            " & vbCrLf & _
               "                            Then ae.Quantidade * - 1                " & vbCrLf & _
               "                            else ae.Quantidade                      " & vbCrLf & _
               "                        end) QtdeAutorizado                         " & vbCrLf & _
               "              from autembarqueXentrega axe                          " & vbCrLf & _
               "             Inner Join AutEmbarque ae                              " & vbCrLf & _
               "                ON axe.Entrega_Id = ae.Entrega_Id                   " & vbCrLf & _
               "               And axe.EndEntrega_Id = ae.EndEntrega_Id             " & vbCrLf & _
               "             Group by ae.Empresa_Id,                                " & vbCrLf & _
               "                      ae.EndEmpresa_Id,                             " & vbCrLf & _
               "                      ae.Pedido_Id,                                 " & vbCrLf & _
               "                      ae.Produto_Id                                 " & vbCrLf

        If RetEntrega Then
            sql &= "                     , ae.Entrega_Id,                           " & vbCrLf & _
                   "                       ae.EndEntrega_Id                         " & vbCrLf
        End If

        sql &= "             ) AutEmb                                               " & vbCrLf & _
               "    on PxI.Empresa_Id    = AutEmb.Empresa_Id                        " & vbCrLf & _
               "   and PxI.EndEmpresa_Id = AutEmb.EndEmpresa_Id                     " & vbCrLf & _
               "   and PxI.Pedido_Id     = AutEmb.Pedido_Id                         " & vbCrLf & _
               "   and PxI.Produto_Id    = AutEmb.Produto_Id                        " & vbCrLf

        If RetEntrega Then
            sql &= "   and AEE.Entrega_Id    = AutEmb.Entrega_Id                    " & vbCrLf & _
                   "   and AEE.EndEntrega_Id = AutEmb.EndEntrega_Id                 " & vbCrLf
        End If

        sql &= "  LEFT JOIN (                                                       " & vbCrLf & _
               "            Select Empresa_Id,                                      " & vbCrLf & _
               "                   EndEmpresa_Id,                                   " & vbCrLf & _
               "                   Pedido,                                          " & vbCrLf & _
               "                   Produto,                                         " & vbCrLf & _
               "                   SUM(case                                         " & vbCrLf & _
               "                         when EntradaSaida = 'E'                    " & vbCrLf & _
               "                           then PesoLiquido                         " & vbCrLf & _
               "                           else 0                                   " & vbCrLf & _
               "                       end) as PesoEntrada,                         " & vbCrLf & _
               "                   SUM(case                                         " & vbCrLf & _
               "                         when EntradaSaida = 'S'                    " & vbCrLf & _
               "                           then PesoLiquido                         " & vbCrLf & _
               "                           else 0                                   " & vbCrLf & _
               "                       end) as PesoSaida                            " & vbCrLf & _
               "              from Romaneios                                        " & vbCrLf & _
               "             group by Empresa_Id,                                   " & vbCrLf & _
               "                      EndEmpresa_Id,                                " & vbCrLf & _
               "                      Pedido,                                       " & vbCrLf & _
               "                      Produto                                       " & vbCrLf & _
               "           ) Rom                                                    " & vbCrLf & _
               "    on PxI.Empresa_Id    = Rom.Empresa_Id                           " & vbCrLf & _
               "   and PxI.EndEmpresa_Id = Rom.EndEmpresa_Id                        " & vbCrLf & _
               "   and PxI.Pedido_Id     = Rom.Pedido                               " & vbCrLf & _
               "   and PxI.Produto_Id    = Rom.Produto                              " & vbCrLf & _
               "  LEFT JOIN (                                                       " & vbCrLf & _
               "			Select AC.Empresa,                                      " & vbCrLf & _
               "				   AC.EndEmpresa,                                   " & vbCrLf & _
               "				   AC.Pedido,                                       " & vbCrLf & _
               "				   AxI.Produto_Id                                   " & vbCrLf & _
               "			  from AutCarregamento AC                               " & vbCrLf & _
               "			 inner join AutCarregamentoXItens AxI                   " & vbCrLf & _
               "				on AC.Carregamento_Id = AxI.Carregamento_Id         " & vbCrLf & _
               "			 Where AC.Situacao = 1                                  " & vbCrLf & _
               "			 group by AC.Empresa,                                   " & vbCrLf & _
               "					  AC.EndEmpresa,                                " & vbCrLf & _
               "					  AC.Pedido,                                    " & vbCrLf & _
               "					  AxI.Produto_id                                " & vbCrLf & _
               "            ) Carr                                                  " & vbCrLf & _
               "    on PxI.Empresa_Id    = Carr.Empresa                             " & vbCrLf & _
               "   and PxI.EndEmpresa_Id = Carr.EndEmpresa                          " & vbCrLf & _
               "   and PxI.Pedido_Id     = Carr.Pedido                              " & vbCrLf & _
               "   and PxI.Produto_Id    = carr.Produto_Id                          " & vbCrLf & _
               " INNER JOIN Clientes C                                              " & vbCrLf & _
               "    on C.Cliente_Id  = P.Cliente                                    " & vbCrLf & _
               "   and C.Endereco_Id = P.EndCliente                                 " & vbCrLf & _
               " INNER JOIN Produtos Prd                                            " & vbCrLf & _
               "    on Prd.Produto_Id = PxI.Produto_Id                              " & vbCrLf & _
               " WHERE (                                                            " & vbCrLf & _
               "            so.EntradaSaida = 'S'                                   " & vbCrLf & _
               "            or                                                      " & vbCrLf & _
               "            (so.EntradaSaida = 'E' and P.FreteCIFFOB = 'FOB')       " & vbCrLf & _
               "        )                                                           " & vbCrLf & _
               "    and P.Situacao = 1                                              " & vbCrLf & _
               "    and P.FiscalAberto = 1                                          " & vbCrLf

        If Not String.IsNullOrWhiteSpace(pedido) Then
            sql &= "    and P.Pedido_Id = '" & pedido & "'" & vbCrLf
        Else
            sql &= "    and P.Safra = '" & safra & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(empresa) AndAlso Not String.IsNullOrWhiteSpace(enderecoEmpresa) Then
            sql &= "    and P.Empresa_Id = '" & empresa & "'" & vbCrLf & _
                   "    and P.EndEmpresa_Id = '" & enderecoEmpresa & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(cliente) AndAlso Not String.IsNullOrWhiteSpace(enderecoCliente) Then
            sql &= "    and P.Cliente    = '" & cliente & "'" & vbCrLf & _
                   "    and P.EndCliente = " & enderecoCliente & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(entradaSaida) Then
            sql &= "    and so.EntradaSaida = '" & entradaSaida & "'" & vbCrLf
        End If

        sql &= "    and Prd.Agrupar = 'N'" & vbCrLf
        '"    and SO.QuantidadePedido = SO.Precofixo                      " & vbCrLf

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "Pedidos")

        For Each row In ds.Tables(0).Rows
            Dim p As New EmbarquePedido
            p.CodigoEmpresa = row("Empresa_Id")
            p.EndEmpresa = row("EndEmpresa_Id")
            p.CodigoCliente = row("Cliente")
            p.EndCliente = row("EndCliente")
            p.NomeCliente = row("NomeCliente")
            p.CodigoSafra = row("Safra")
            p.CodigoPedido = row("Pedido_Id")
            p.DataEntrega = row("DataEntrega")
            p.EmbarqueAtivo = row("EmbarqueAtivo")
            p.DescOperacao = row("DescOperacao")
            p.ES = row("EntradaSaida")
            p.CidadeEntrega = row("CidadeEntrega")
            p.EstadoEntrega = row("EstadoEntrega")
            p.CodigoProduto = row("Produto_Id")
            p.NomeProduto = row("NomeProduto")
            p.QtdePedido = row("QtdePedido")
            p.QtdeNormal = row("QtdeNormal")
            p.QtdeComplemento = row("QtdeComplemento")
            p.QtdeEstorno = row("QtdeEstorno")
            p.QtdeAutorizadoEmbarque = row("QtdeAutorizadoEmbarque")
            p.Embarcado = row("Embarcado")
            p.Devolvido = row("Devolvido")
            Me.Add(p)
        Next
    End Sub

    'Public Sub New(ByVal objPedido As Pedido, ByVal Entrega_Id As String, ByVal EndEntrega_Id As Integer, ByVal Roteiro_Id As Integer)
    '    Dim sql As String = ""
    '    sql &= "select P.Empresa_Id," & vbCrLf & _
    '           "       P.EndEmpresa_Id," & vbCrLf & _
    '           "       P.Cliente," & vbCrLf & _
    '           "       P.EndCliente," & vbCrLf & _
    '           "       C.Nome as NomeCliente," & vbCrLf & _
    '           "       P.Safra," & vbCrLf & _
    '           "       P.Pedido_Id," & vbCrLf & _
    '           "        aer.Roteiro_Id," & vbCrLf & _
    '           "       convert(varchar,P.DataEntrega, 103) as DataEntrega," & vbCrLf & _
    '           "       P.EmbarqueAtivo," & vbCrLf & _
    '           "       isnull(Cast((AEE.Entrega_Id + '-' + Cast(AEE.EndEntrega_Id as nvarchar)) as nvarchar),'') as Entrega," & vbCrLf & _
    '           "       Case" & vbCrLf & _
    '           "         When P.EmbarqueAtivo = 1" & vbCrLf & _
    '           "           Then 'Ativo'" & vbCrLf & _
    '           "           else 'Bloqueado'" & vbCrLf & _
    '           "       End as Ativo," & vbCrLf & _
    '           "       SO.EntradaSaida," & vbCrLf & _
    '           "       isnull(P.CidadeEntrega,'') as CidadeEntrega," & vbCrLf & _
    '           "       ISNULL(P.EstadoEntrega,'') as EstadoEntrega," & vbCrLf & _
    '           "       Pxi.Produto_Id," & vbCrLf & _
    '           "       Prd.Nome as NomeProduto," & vbCrLf & _
    '           "       PxI.Qtde as QtdePedido," & vbCrLf & _
    '           "       isnull(AutEmb.QtdeNormal,0) as QtdeNormal," & vbCrLf & _
    '           "       isnull(AutEmb.QtdeComplemento,0) as QtdeComplemento," & vbCrLf & _
    '           "       isnull(AutEmb.QtdeEstorno,0) as QtdeEstorno," & vbCrLf & _
    '           "       isnull(AutEmb.QtdeAutorizado,0) as QtdeAutorizadoEmbarque," & vbCrLf & _
    '           "       isnull(AutEmb.QtdeCarregamento,0) as QtdeCarregamento," & vbCrLf & _
    '           "       isnull(Case" & vbCrLf & _
    '           "				 when SO.EntradaSaida = 'S'" & vbCrLf & _
    '           "				   then Rom.PesoSaida" & vbCrLf & _
    '           "				   else Rom.PesoEntrada" & vbCrLf & _
    '           "			   end,0) Embarcado," & vbCrLf & _
    '           "       isnull(Case" & vbCrLf & _
    '           "				 when SO.EntradaSaida = 'S'" & vbCrLf & _
    '           "				   then Rom.PesoEntrada" & vbCrLf & _
    '           "				   else Rom.PesoSaida" & vbCrLf & _
    '           "			   end,0) Devolvido" & vbCrLf & _
    '           "  from Pedidos P" & vbCrLf & _
    '           " Inner Join (Select Empresa_Id," & vbCrLf & _
    '           "                    EndEmpresa_Id," & vbCrLf & _
    '           "                    Pedido_Id," & vbCrLf & _
    '           "                    Produto_Id," & vbCrLf & _
    '           "                    SUM(Case" & vbCrLf & _
    '           "                          when TipoDeLancamento = 'E'" & vbCrLf & _
    '           "                            Then Quantidade * - 1" & vbCrLf & _
    '           "                            else Quantidade" & vbCrLf & _
    '           "                        end) Qtde," & vbCrLf & _
    '           "                    SUM(Case" & vbCrLf & _
    '           "                          when TipoDeLancamento = 'E'" & vbCrLf & _
    '           "                            Then isnull(QuantidadeEstimada,0) * - 1" & vbCrLf & _
    '           "                            else isnull(QuantidadeEstimada,0)" & vbCrLf & _
    '           "                        end) QtdeEstimada" & vbCrLf & _
    '           "               from PedidoXItemXLancamento" & vbCrLf & _
    '           "              Group by Empresa_Id," & vbCrLf & _
    '           "                       EndEmpresa_Id," & vbCrLf & _
    '           "                       Pedido_Id," & vbCrLf & _
    '           "                       Produto_Id" & vbCrLf & _
    '           "             ) PxI" & vbCrLf & _
    '           "    on P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
    '           "   and P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf & _
    '           "   and P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf & _
    '           " Inner Join SubOperacoes  SO" & vbCrLf & _
    '           "    on P.Operacao    = SO.Operacao_Id" & vbCrLf & _
    '           "   and p.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
    '           "  Left join AutEmbarqueXEntrega AEE" & vbCrLf & _
    '           "    on AEE.Empresa_Id    = P.Empresa_Id" & vbCrLf & _
    '           "   AND AEE.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
    '           "   AND AEE.Pedido_Id     = P.Pedido_Id" & vbCrLf & _
    '           "  Left join (Select ae.Empresa_Id," & vbCrLf & _
    '           "                    ae.EndEmpresa_Id," & vbCrLf & _
    '           "                    ae.Pedido_Id," & vbCrLf & _
    '           "                    ae.Produto_Id," & vbCrLf & _
    '           "                    ae.Entrega_Id," & vbCrLf & _
    '           "                    ae.EndEntrega_Id," & vbCrLf & _
    '           "                    SUM(Case" & vbCrLf & _
    '           "                          when ae.TipoDeLancamento = 'N'" & vbCrLf & _
    '           "                            Then ae.Quantidade" & vbCrLf & _
    '           "                            else 0" & vbCrLf & _
    '           "                        end) QtdeNormal," & vbCrLf & _
    '           "                    SUM(Case" & vbCrLf & _
    '           "                          when ae.TipoDeLancamento = 'C'" & vbCrLf & _
    '           "                            Then ae.Quantidade" & vbCrLf & _
    '           "                            else 0" & vbCrLf & _
    '           "                        end) QtdeComplemento," & vbCrLf & _
    '           "                    SUM(Case" & vbCrLf & _
    '           "                          when ae.TipoDeLancamento = 'E'" & vbCrLf & _
    '           "                            Then ae.Quantidade" & vbCrLf & _
    '           "                            else 0" & vbCrLf & _
    '           "                        end) QtdeEstorno," & vbCrLf & _
    '           "                    SUM(Case" & vbCrLf & _
    '           "                          when ae.TipoDeLancamento = 'E'" & vbCrLf & _
    '           "                            Then ae.Quantidade * - 1" & vbCrLf & _
    '           "                            else ae.Quantidade" & vbCrLf & _
    '           "                        end) QtdeAutorizado," & vbCrLf & _
    '           "                    SUM(Case" & vbCrLf & _
    '           "                               when ae.TipoDeLancamento = 'E'" & vbCrLf & _
    '           "                                 Then ae.Quantidade * - 1" & vbCrLf & _
    '           "                                 else ae.Quantidade" & vbCrLf & _
    '           "                             end) - sum(isnull(acxi.pesoprogramado,0)) QtdeCarregamento" & vbCrLf & _
    '           "               from autembarqueXentrega axe" & vbCrLf & _
    '           "              Inner Join AutEmbarque ae" & vbCrLf & _
    '           "                 ON axe.Entrega_Id = ae.Entrega_Id" & vbCrLf & _
    '           "                And axe.EndEntrega_Id = ae.EndEntrega_Id " & vbCrLf & _
    '           "               Left Join AutCarregamento ac" & vbCrLf & _
    '           "                 ON ae.empresa_Id    = ac.Empresa" & vbCrLf & _
    '           "                and ae.endEmpresa_Id = ac.endempresa" & vbCrLf & _
    '           "                and ae.pedido_id     = ac.pedido" & vbCrLf & _
    '           "               Left Join autcarregamentoxitens acxi" & vbCrLf & _
    '           "                 on acxi.carregamento_id = ac.carregamento_id" & vbCrLf & _
    '           "                and acxi.produto_id      = ae.produto_id" & vbCrLf & _
    '           "              Group by ae.Empresa_Id," & vbCrLf & _
    '           "                       ae.EndEmpresa_Id," & vbCrLf & _
    '           "                       ae.Pedido_Id," & vbCrLf & _
    '           "                       ae.Produto_Id," & vbCrLf & _
    '           "                       ae.Entrega_Id," & vbCrLf & _
    '           "                       ae.EndEntrega_Id" & vbCrLf & _
    '           "             ) AutEmb" & vbCrLf & _
    '           "    on PxI.Empresa_Id    = AutEmb.Empresa_Id" & vbCrLf & _
    '           "   and PxI.EndEmpresa_Id = AutEmb.EndEmpresa_Id" & vbCrLf & _
    '           "   and PxI.Pedido_Id     = AutEmb.Pedido_Id" & vbCrLf & _
    '           "   and PxI.Produto_Id    = AutEmb.Produto_Id" & vbCrLf & _
    '           "   and AEE.Entrega_Id    = AutEmb.Entrega_Id" & vbCrLf & _
    '           "   and AEE.EndEntrega_Id = AutEmb.EndEntrega_Id" & vbCrLf & _
    '           " inner join AutEmbarqueXRoteiro AER" & vbCrLf & _
    '           "    on AEE.Empresa_Id    = AER.Empresa_Id" & vbCrLf & _
    '           "   AND AEE.EndEmpresa_Id = AER.EndEmpresa_Id" & vbCrLf & _
    '           "   AND AEE.Pedido_Id     = AER.Pedido_Id " & vbCrLf & _
    '           "   AND AEE.Entrega_Id    = AER.Entrega_Id" & vbCrLf & _
    '           "   AND AEE.EndEntrega_Id = AER.EndEntrega_Id" & vbCrLf & _
    '           "  Left Join(" & vbCrLf & _
    '           "            Select Empresa_Id," & vbCrLf & _
    '           "                   EndEmpresa_Id," & vbCrLf & _
    '           "                   Pedido," & vbCrLf & _
    '           "                   Produto," & vbCrLf & _
    '           "                   SUM(case" & vbCrLf & _
    '           "                         when EntradaSaida = 'E'" & vbCrLf & _
    '           "                           then PesoLiquido" & vbCrLf & _
    '           "                           else 0" & vbCrLf & _
    '           "                       end) as PesoEntrada," & vbCrLf & _
    '           "                   SUM(case" & vbCrLf & _
    '           "                         when EntradaSaida = 'S'" & vbCrLf & _
    '           "                           then PesoLiquido" & vbCrLf & _
    '           "                           else 0" & vbCrLf & _
    '           "                       end) as PesoSaida" & vbCrLf & _
    '           "              from Romaneios" & vbCrLf & _
    '           "             group by Empresa_Id," & vbCrLf & _
    '           "                      EndEmpresa_Id," & vbCrLf & _
    '           "                      Pedido," & vbCrLf & _
    '           "                      Produto" & vbCrLf & _
    '           "           ) Rom" & vbCrLf & _
    '           "    on PxI.Empresa_Id    = Rom.Empresa_Id" & vbCrLf & _
    '           "   and PxI.EndEmpresa_Id = Rom.EndEmpresa_Id" & vbCrLf & _
    '           "   and PxI.Pedido_Id     = Rom.Pedido" & vbCrLf & _
    '           "   and PxI.Produto_Id    = Rom.Produto" & vbCrLf & _
    '           "  left Join(" & vbCrLf & _
    '           "			Select AC.Empresa," & vbCrLf & _
    '           "				   AC.EndEmpresa," & vbCrLf & _
    '           "				   AC.Pedido," & vbCrLf & _
    '           "				   AxI.Produto_Id" & vbCrLf & _
    '           "			  from AutCarregamento AC" & vbCrLf & _
    '           "			 inner join AutCarregamentoXItens AxI" & vbCrLf & _
    '           "				on AC.Carregamento_Id = AxI.Carregamento_Id" & vbCrLf & _
    '           "			 Where AC.Situacao = 1" & vbCrLf & _
    '           "			 group by AC.Empresa," & vbCrLf & _
    '           "					  AC.EndEmpresa," & vbCrLf & _
    '           "					  AC.Pedido," & vbCrLf & _
    '           "					  AxI.Produto_id" & vbCrLf & _
    '           "            ) Carr" & vbCrLf & _
    '           "    on PxI.Empresa_Id    = Carr.Empresa" & vbCrLf & _
    '           "   and PxI.EndEmpresa_Id = Carr.EndEmpresa" & vbCrLf & _
    '           "   and PxI.Pedido_Id     = Carr.Pedido" & vbCrLf & _
    '           "   and PxI.Produto_Id    = carr.Produto_Id" & vbCrLf & _
    '           " Inner Join Clientes C" & vbCrLf & _
    '           "    on C.Cliente_Id  = P.Cliente" & vbCrLf & _
    '           "   and C.Endereco_Id = P.EndCliente" & vbCrLf & _
    '           " Inner Join Produtos Prd" & vbCrLf & _
    '           "    on Prd.Produto_Id = PxI.Produto_Id" & vbCrLf & _
    '           " Where (" & vbCrLf & _
    '           "         so.EntradaSaida = 'S'" & vbCrLf & _
    '           "        or" & vbCrLf & _
    '           "        (so.EntradaSaida = 'E' and P.FreteCIFFOB = 'CIF')" & vbCrLf & _
    '           "        )" & vbCrLf & _
    '           "   and P.Safra                         ='" & objPedido.CodigoSafra & "'" & vbCrLf & _
    '           "   and P.Empresa_Id                    ='" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
    '           "   and P.EndEmpresa_Id                 = " & objPedido.EnderecoEmpresa & vbCrLf & _
    '           "   and P.Cliente                       ='" & objPedido.CodigoCliente & "'" & vbCrLf & _
    '           "   and P.EndCliente                    = " & objPedido.EnderecoCliente & vbCrLf & _
    '           "   and so.EntradaSaida                 ='" & IIf(objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
    '           "   and Prd.Agrupar                     ='N'" & vbCrLf & _
    '           "   and SO.QuantidadePedido             = SO.Precofixo" & vbCrLf & _
    '           "   and isnull(AutEmb.QtdeAutorizado,0) > 0" & vbCrLf & _
    '           "   and AEE.Entrega_Id                  ='" & Entrega_Id & "'" & vbCrLf & _
    '           "   and AEE.Entrega_Id                  = " & Entrega_Id & vbCrLf & _
    '           "   and aer.Roteiro_Id                  = " & Roteiro_Id & vbCrLf

    '    Dim db As New AcessaBanco
    '    Dim ds As DataSet = db.ConsultaDataSet(sql, "Pedidos")

    '    For Each row In ds.Tables(0).Rows
    '        Dim p As New EmbarquePedido
    '        p.ProdutoCodigoNome = row("Produto_Id") & "-" & row("NomeProduto")
    '        p.CodigoEmpresa = row("Empresa_Id")
    '        p.EndEmpresa = row("EndEmpresa_Id")
    '        p.CodigoCliente = row("Cliente")
    '        p.EndCliente = row("EndCliente")
    '        p.NomeCliente = row("NomeCliente")
    '        p.CodigoSafra = row("Safra")
    '        p.CodigoPedido = row("Pedido_Id")
    '        p.DataEntrega = row("DataEntrega")
    '        p.EmbarqueAtivo = row("EmbarqueAtivo")
    '        p.ES = row("EntradaSaida")
    '        p.CidadeEntrega = row("CidadeEntrega")
    '        p.EstadoEntrega = row("EstadoEntrega")
    '        p.CodigoProduto = row("Produto_Id")
    '        p.NomeProduto = row("NomeProduto")
    '        p.QtdePedido = row("QtdePedido")
    '        p.QtdeNormal = row("QtdeNormal")
    '        p.QtdeComplemento = row("QtdeComplemento")
    '        p.QtdeEstorno = row("QtdeEstorno")
    '        p.QtdeAutorizadoEmbarque = row("QtdeAutorizadoEmbarque")
    '        p.QtdeCarregamento = row("QtdeCarregamento")
    '        p.Embarcado = row("Embarcado")
    '        p.Devolvido = row("Devolvido")
    '        p.Ativo = row("Ativo")
    '        Me.Add(p)
    '    Next
    'End Sub

    Public Sub New(ByVal transportador() As String)
        Dim sql As String = "Select tr.empresa_id, tr.endempresa_id, tr.pedido_id, ApPedido.cliente, ApPedido.endcliente, c.nome   " & vbCrLf & _
                            "  from AutEmbarqueXPrecoFrete Pf" & vbCrLf & _
                            " inner join AutEmbarqueXTransportador Tr " & vbCrLf & _
                            "    on	tr.Empresa_Id = pf.Empresa_Id" & vbCrLf & _
                            "   and tr.EndEmpresa_Id = pf.EndEmpresa_Id" & vbCrLf & _
                            "   and tr.Pedido_Id = pf.Pedido_Id" & vbCrLf & _
                            "   and tr.Transportador_Id = pf.Transportador_Id" & vbCrLf & _
                            "   and tr.EndTransportar_Id = pf.EndTransportador_Id" & vbCrLf & _
                            "   and tr.Roteiro_Id = pf.Roteiro_Id" & vbCrLf & _
                            " inner Join(" & vbCrLf & _
                            "            SELECT ae.Empresa_Id," & vbCrLf & _
                            "                   ae.EndEmpresa_Id," & vbCrLf & _
                            "                   ae.Pedido_Id," & vbCrLf & _
                            "                   P.Cliente," & vbCrLf & _
                            "                   P.EndCliente," & vbCrLf & _
                            "                   ae.Produto_Id," & vbCrLf & _
                            "                   sum(case" & vbCrLf & _
                            "                         when ae.TipoDeLancamento = 'E'" & vbCrLf & _
                            "                           then ae.Quantidade * -1" & vbCrLf & _
                            "                           else ae.Quantidade" & vbCrLf & _
                            "                       end) Autorizado," & vbCrLf & _
                            " 				    isnull(Case" & vbCrLf & _
                            "  							 when SO.EntradaSaida = 'S'" & vbCrLf & _
                            "  							   then Rom.PesoSaida" & vbCrLf & _
                            "  							   else Rom.PesoEntrada" & vbCrLf & _
                            "  						   end,0) Embarcado," & vbCrLf & _
                            "  				    isnull(Case" & vbCrLf & _
                            "  					 		 when SO.EntradaSaida = 'S'" & vbCrLf & _
                            "  							   then Rom.PesoEntrada" & vbCrLf & _
                            "  							   else Rom.PesoSaida" & vbCrLf & _
                            "  						   end,0) Devolvido" & vbCrLf & _
                            "              FROM AutEmbarque ae" & vbCrLf & _
                            "             inner join Pedidos P" & vbCrLf & _
                            "                on ae.Empresa_Id    = P.Empresa_Id" & vbCrLf & _
                            "               and ae.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
                            "               and ae.Pedido_Id     = P.Pedido_Id" & vbCrLf & _
                            "             inner Join Suboperacoes SO" & vbCrLf & _
                            "                on P.Operacao    = SO.Operacao_id" & vbCrLf & _
                            "               and P.suboperacao = SO.SubOperacoes_id" & vbCrLf & _
                            "              LEFT JOIN(" & vbCrLf & _
                            "  						Select Empresa_Id," & vbCrLf & _
                            "  							   EndEmpresa_Id," & vbCrLf & _
                            "  							   Pedido," & vbCrLf & _
                            "  							   Produto," & vbCrLf & _
                            "  							   SUM(case" & vbCrLf & _
                            "  									 when EntradaSaida = 'E'" & vbCrLf & _
                            "  									   then PesoLiquido" & vbCrLf & _
                            "  									   else 0" & vbCrLf & _
                            "  								   end) as PesoEntrada," & vbCrLf & _
                            "  							   SUM(case" & vbCrLf & _
                            "  									 when EntradaSaida = 'S'" & vbCrLf & _
                            "  									   then PesoLiquido" & vbCrLf & _
                            "  									   else 0" & vbCrLf & _
                            "  								   end) as PesoSaida" & vbCrLf & _
                            "  						  from Romaneios" & vbCrLf & _
                            "  						 group by Empresa_Id," & vbCrLf & _
                            "  								  EndEmpresa_Id," & vbCrLf & _
                            "  								  Pedido," & vbCrLf & _
                            "  								  Produto" & vbCrLf & _
                            "  					     ) Rom" & vbCrLf & _
                            "  				  on ae.Empresa_Id    = Rom.Empresa_Id" & vbCrLf & _
                            "                and ae.EndEmpresa_Id = Rom.EndEmpresa_Id" & vbCrLf & _
                            "                and ae.Pedido_Id     = Rom.Pedido" & vbCrLf & _
                            "                and ae.Produto_Id    = Rom.Produto" & vbCrLf & _
                            "            --where P.PedidoBloqueado = 1" & vbCrLf & _
                            "              Group By ae.Empresa_Id," & vbCrLf & _
                            "                       ae.EndEmpresa_Id," & vbCrLf & _
                            "                       ae.Pedido_Id," & vbCrLf & _
                            "                       ae.Produto_Id," & vbCrLf & _
                            "                       P.Cliente," & vbCrLf & _
                            "                       P.EndCliente," & vbCrLf & _
                            "                       isnull(Case" & vbCrLf & _
                            "  			          	          when SO.EntradaSaida = 'S'" & vbCrLf & _
                            "  							        then Rom.PesoSaida" & vbCrLf & _
                            "  							        else Rom.PesoEntrada" & vbCrLf & _
                            "  						        end,0)," & vbCrLf & _
                            "  				         isnull(Case" & vbCrLf & _
                            "  						          when SO.EntradaSaida = 'S'" & vbCrLf & _
                            "  							        then Rom.PesoEntrada" & vbCrLf & _
                            "  							        else Rom.PesoSaida" & vbCrLf & _
                            "  						        end,0)" & vbCrLf & _
                            "          ) ApPedido" & vbCrLf & _
                            "      on ApPedido.Empresa_id     = Tr.Empresa_id" & vbCrLf & _
                            "     and ApPedido.EndEmpresa_Id  = Tr.EndEmpresa_Id" & vbCrLf & _
                            "     and ApPedido.Pedido_id      = Tr.Pedido_id" & vbCrLf & _
                            "     and ApPedido.autorizado - ApPedido.Embarcado > 0" & vbCrLf & _
                            "    inner join Clientes C" & vbCrLf & _
                            "       on ApPedido.cliente    = C.cliente_id" & vbCrLf & _
                            "      and ApPedido.endCliente = C.endereco_id" & vbCrLf & _
                            "   where Tr.Transportador_Id  = '" & transportador(0) & "'" & vbCrLf & _
                            "      and Tr.EndTransportar_Id = '" & transportador(1) & "'" & vbCrLf & _
                            "      and tr.Ativo = 1" & vbCrLf & _
                            "      and (pf.ValorFrete > 0 or pf.ValorTon > 0)" & vbCrLf & _
                            "      Group By tr.empresa_id, tr.endempresa_id, tr.pedido_id, ApPedido.cliente, ApPedido.endcliente, c.nome"


        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "Pedidos")

        For Each row In ds.Tables(0).Rows
            Dim p As New EmbarquePedido
            p.CodigoEmpresa = row("Empresa_Id")
            p.EndEmpresa = row("EndEmpresa_Id")
            p.CodigoPedido = row("Pedido_Id")
            p.CodigoCliente = row("Cliente")
            p.EndCliente = row("endCliente")
            p.NomeCliente = row("nome")
            Me.Add(p)
        Next
    End Sub

End Class

<Serializable()> _
Public Class EmbarquePedido
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New()
    End Sub

#End Region

#Region "Fields"

    Private _ProdutoCodigoNome As String = ""
    Private _CnpjEmpresa As String = ""
    Private _Entrega_Id As String = ""
    Private _EndEntrega_Id As Integer
    Private _Entrega As String = ""

    Private _CodigoEmpresa As String = ""
    Private _EndEmpresa As Integer
    Private _Empresa As Cliente

    Private _CodigoCliente As String = ""
    Private _EndCliente As Integer
    Private _NomeCliente As String = ""
    Private _ClienteFormatado As String = ""
    Private _Cliente As Cliente

    Private _CodigoSafra As String = ""

    Private _CodigoPedido As Integer
    Private _Pedido As Pedido

    Private _DataEntrega As Date
    Private _EmbarqueAtivo As String = ""
    Private _DescOperacao As String = ""
    Private _ES As String
    Private _CidadeEntrega As String = ""
    Private _EstadoEntrega As String = ""

    Private _CodigoProduto As Integer
    Private _NomeProduto As String = ""
    Private _Produto As Produto

    Private _QtdePedido As Decimal
    Private _QtdeNormal As Decimal
    Private _QtdeComplemento As Decimal
    Private _QtdeEstorno As Decimal
    Private _QtdeAutorizadoEmbarque As Decimal
    'QTDECARREGAMENTO É O MESMO QUE (AUTORIZADO - CARREGADO)
    Private _QtdeCarregamento As Decimal
    Private _MaximoAutorizar As Decimal
    Private _MaximoEstorno As Decimal
    Private _Embarcado As Decimal
    Private _Devolvido As Decimal
    Private _LocaisDeEntrega As ListEmbarqueXEntrega
#End Region

#Region "Property"

    Public Property ProdutoCodigoNome() As String
        Get
            Return _ProdutoCodigoNome
        End Get
        Set(ByVal value As String)
            _ProdutoCodigoNome = value
        End Set
    End Property

    Public Property CnpjEmpresa() As String
        Get
            Return CodigoEmpresa & "-" & EndEmpresa
        End Get
        Set(ByVal value As String)
            _CnpjEmpresa = value
        End Set
    End Property

    Public Property CodigoEmpresa As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property EndEmpresa As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property Empresa As Cliente
        Get
            If _Empresa Is Nothing And Me.CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(Me.CodigoEmpresa, Me.EndEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property CodigoCliente As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
        End Set
    End Property

    Public Property EndCliente As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
        End Set
    End Property

    Public Property NomeCliente As String
        Get
            '_NomeCliente = Funcoes.FormatarListItemCliente(Cliente).Text
            Return _NomeCliente
        End Get
        Set(ByVal value As String)
            _NomeCliente = value
        End Set
    End Property

    Public ReadOnly Property ClienteFormatado() As String
        Get
            _ClienteFormatado = Me.CodigoPedido & " -Cliente-" & NomeCliente & "-" & Me.CodigoCliente & "-" & Me.EndCliente
            Return _ClienteFormatado
        End Get
    End Property

    Public ReadOnly Property Cliente As Cliente
        Get
            If _Cliente Is Nothing And Me.CodigoCliente.Length > 0 Then _Cliente = New Cliente(Me.CodigoCliente, Me.EndCliente)
            Return _Cliente
        End Get
    End Property

    Public Property CodigoSafra As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
        End Set
    End Property

    Public Property CodigoPedido As Integer
        Get
            Return _CodigoPedido
        End Get
        Set(ByVal value As Integer)
            _CodigoPedido = value
        End Set
    End Property

    Public ReadOnly Property Pedido As Pedido
        Get
            If (_Pedido Is Nothing OrElse Not _Pedido.Codigo > 0) And Me.CodigoPedido > 0 Then _Pedido = New Pedido(Me.CodigoEmpresa, Me.EndEmpresa, Me.CodigoPedido)
            Return _Pedido
        End Get
    End Property

    Public Property DataEntrega As Date
        Get
            Return _DataEntrega
        End Get
        Set(ByVal value As Date)
            _DataEntrega = value
        End Set
    End Property

    Public Property EmbarqueAtivo() As String
        Get
            Return _EmbarqueAtivo
        End Get
        Set(ByVal value As String)
            _EmbarqueAtivo = value
        End Set
    End Property

    Public Property DescOperacao As String
        Get
            Return _DescOperacao
        End Get
        Set(value As String)
            _DescOperacao = value
        End Set
    End Property

    Public Property ES As String
        Get
            Return _ES
        End Get
        Set(ByVal value As String)
            _ES = value
        End Set
    End Property

    Public Property CidadeEntrega As String
        Get
            Return _CidadeEntrega
        End Get
        Set(ByVal value As String)
            _CidadeEntrega = value
        End Set
    End Property

    Public Property EstadoEntrega As String
        Get
            Return _EstadoEntrega
        End Get
        Set(ByVal value As String)
            _EstadoEntrega = value
        End Set
    End Property

    Public Property CodigoProduto As Integer
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As Integer)
            _CodigoProduto = value
        End Set
    End Property

    Public Property NomeProduto As String
        Get
            Return _NomeProduto
        End Get
        Set(ByVal value As String)
            _NomeProduto = value
        End Set
    End Property

    Public ReadOnly Property Produto As Produto
        Get
            If _Produto Is Nothing And Me.CodigoProduto > 0 Then _Produto = New Produto(Me.CodigoProduto)
            Return _Produto
        End Get
    End Property

    Public Property QtdeNormal As Decimal
        Get
            Return _QtdeNormal
        End Get
        Set(ByVal value As Decimal)
            _QtdeNormal = value
        End Set
    End Property

    Public Property QtdeComplemento As Decimal
        Get
            Return _QtdeComplemento
        End Get
        Set(ByVal value As Decimal)
            _QtdeComplemento = value
        End Set
    End Property

    Public Property QtdeEstorno As Decimal
        Get
            Return _QtdeEstorno
        End Get
        Set(ByVal value As Decimal)
            _QtdeEstorno = value
        End Set
    End Property

    Public Property QtdeAutorizadoEmbarque As Decimal
        Get
            Return _QtdeAutorizadoEmbarque
        End Get
        Set(ByVal value As Decimal)
            _QtdeAutorizadoEmbarque = value
        End Set
    End Property

    Public Property QtdeCarregamento() As Decimal
        Get
            Return _QtdeCarregamento
        End Get
        Set(ByVal value As Decimal)
            _QtdeCarregamento = value
        End Set
    End Property

    Public Property MaximoAutorizar() As Decimal
        Get
            Return QtdePedido + Devolvido
        End Get
        Set(ByVal value As Decimal)
            _MaximoAutorizar = value
        End Set
    End Property

    Public Property MaximoEstorno() As Decimal
        Get
            Return QtdeAutorizadoEmbarque - Embarcado
        End Get
        Set(ByVal value As Decimal)
            _MaximoEstorno = value
        End Set
    End Property

    Public Property Embarcado As Decimal
        Get
            Return _Embarcado
        End Get
        Set(ByVal value As Decimal)
            _Embarcado = value
        End Set
    End Property

    Public Property Devolvido As Decimal
        Get
            Return _Devolvido
        End Get
        Set(ByVal value As Decimal)
            _Devolvido = value
        End Set
    End Property

    Public Property QtdePedido As Decimal
        Get
            Return _QtdePedido
        End Get
        Set(ByVal value As Decimal)
            _QtdePedido = value
        End Set
    End Property

    Public Property LocaisDeEntrega() As ListEmbarqueXEntrega
        Get
            If _LocaisDeEntrega Is Nothing Then _LocaisDeEntrega = New ListEmbarqueXEntrega(Me)
            Return _LocaisDeEntrega
        End Get
        Set(ByVal value As ListEmbarqueXEntrega)
            _LocaisDeEntrega = value
        End Set
    End Property
#End Region

End Class
