Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCRMxCliente
    Inherits List(Of CRMxCliente)

#Region "Construtor"
    Public Sub New(ByVal pCrmParametros As CRM)
        _CrmParametros = pCrmParametros
        If pCrmParametros.IUD = "I" Then
            CarregarNova()
        Else
            CarregarSalva()
        End If
    End Sub
#End Region

#Region "Fields"
    Private _TotalVendas As Decimal
    Private _TotalCompras As Decimal
    Private _CrmParametros As CRM
#End Region

#Region "Property"
    Private ReadOnly Property TotalVendas() As Decimal
        Get
            Return _TotalVendas
        End Get
    End Property

    Private ReadOnly Property TotalCompras() As Decimal
        Get
            Return _TotalCompras
        End Get
    End Property

    Public ReadOnly Property CrmParametros() As CRM
        Get
            Return _CrmParametros
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub Atualizar()
        '*****************************************
        '******** Levantamento dos Totais ********
        '*****************************************
        _TotalVendas = 0
        _TotalCompras = 0
        For Each row As CRMxCliente In Me
            If row.TipoClienteCRM <> "V" And row.TipoClienteCRM <> "" Then
                _TotalVendas += row.TotalVendas
                _TotalCompras += row.TotalCompras
            End If
        Next

        '*****************************************
        '********* Atualizacao da lista **********
        '*****************************************
        Dim FatAcumulado As Decimal
        For Each row As CRMxCliente In Me
            If row.TipoClienteCRM <> "V" Then
                FatAcumulado += row.TotalVendas
                row.PercentualFaturamentoAcumulado = Math.Round(FatAcumulado * 100 / Me.TotalVendas, 2)
                If _CrmParametros.CRMPercentualCorte > row.PercentualFaturamentoAcumulado Then
                    _CrmParametros.CRMValorCorte = row.TotalVendas
                    row.TipoClienteCRM = "C"
                ElseIf row.TipoClienteCRM <> "P" And row.TipoClienteCRM <> "E" Then
                    row.TipoClienteCRM = "M"
                    row.PontuacaoMBV = 0
                    row.PontuacaoSOC = 0
                    row.PontuacaoRating = 0
                End If

                If row.TipoClienteCRM <> "M" Then
                    Select Case row.PercentualMargemBruta
                        Case Is <= _CrmParametros.MBVPercentualMenor : row.PontuacaoMBV = _CrmParametros.MBVPontuacaoMenor
                        Case Is >= _CrmParametros.MBVPercentualMaior : row.PontuacaoMBV = _CrmParametros.MBVPontuacaoMaior
                        Case Else : row.PontuacaoMBV = _CrmParametros.MBVPontuacaoEntre
                    End Select

                    Select Case ((row.PercentualVendas * _CrmParametros.SOCPesoVenda) + (row.PercentualCompras * _CrmParametros.SOCPesoCompra)) / 100
                        Case Is <= _CrmParametros.SOCPercentualMenor : row.PontuacaoSOC = _CrmParametros.SOCPontuacaoMenor
                        Case Is >= _CrmParametros.SOCPercentualMaior : row.PontuacaoSOC = _CrmParametros.SOCPontuacaoMaior
                        Case Else : row.PontuacaoSOC = _CrmParametros.SOCPontuacaoEntre
                    End Select

                    Select Case row.LimiteCredito
                        Case "A" : row.PontuacaoRating = _CrmParametros.RatingPontuacaoA
                        Case "B" : row.PontuacaoRating = _CrmParametros.RatingPontuacaoB
                        Case "C" : row.PontuacaoRating = _CrmParametros.RatingPontuacaoC
                        Case Else : row.PontuacaoRating = 0
                    End Select
                End If
            Else
                row.PercentualFaturamentoAcumulado = 0
                row.TipoClienteCRM = "V"
                row.PontuacaoMBV = 0
                row.PontuacaoSOC = 0
                row.PontuacaoRating = 0
            End If
        Next
    End Sub

    Public Sub CarregarNova()
        Dim sql As String = ""
        '******** MARGEM BRUTA DE VENDA *********
        sql = "Select P.Cliente," & vbCrLf & _
              " 	  P.EndCliente," & vbCrLf & _
              " 	  Cli.Nome," & vbCrLf & _
              "       p.Pedido_Id," & vbCrLf & _
              "       PxI.Produto_Id," & vbCrLf & _
              "       max(isnull(PxI.UnitarioOficialCompra,0)) as UnitarioOficialCompra," & vbCrLf & _
              "       max(isnull(PxI.UnitarioMoedaCompra,0))   as UnitarioMoedaCompra," & vbCrLf & _
              " 	  Sum(case" & vbCrLf & _
              " 			 When PxI.TipoDeLancamento = 'E'" & vbCrLf & _
              " 			   then PxI.Quantidade * - 1" & vbCrLf & _
              " 			   else PxI.Quantidade" & vbCrLf & _
              " 		   end) as Quantidade," & vbCrLf & _
              " 	  Sum(case" & vbCrLf & _
              " 		  	 When PxI.TipoDeLancamento = 'E'" & vbCrLf & _
              " 			   then PxI.TotalOficial * - 1" & vbCrLf & _
              " 			   else PxI.TotalOficial" & vbCrLf & _
              " 		   end) as TotalOficial," & vbCrLf & _
              " 	  Sum(case" & vbCrLf & _
              " 		     When PxI.TipoDeLancamento = 'E'" & vbCrLf & _
              " 			   then PxI.TotalMoeda * - 1" & vbCrLf & _
              " 			   else PxI.TotalMoeda" & vbCrLf & _
              " 		   end) as TotalMoeda" & vbCrLf & _
              "  into #temp" & vbCrLf & _
              "  FROM Pedidos P" & vbCrLf & _
              " Inner Join Operacoes OP" & vbCrLf & _
              "    on OP.Operacao_Id     = P.Operacao" & vbCrLf & _
              " Inner Join Suboperacoes SO" & vbCrLf & _
              "    on SO.Operacao_Id     = P.Operacao" & vbCrLf & _
              "   and SO.Suboperacoes_Id = P.Suboperacao" & vbCrLf & _
              " Inner Join PedidoXItemxLancamento PxI" & vbCrLf & _
              "    on P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf & _
              "   and P.EndEmpresa_Id = PXI.EndEmpresa_Id" & vbCrLf & _
              "   and P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf & _
              " INNER JOIN Produtos Prod" & vbCrLf & _
              "    ON PxI.Produto_Id = Prod.Produto_Id" & vbCrLf & _
              " INNER JOIN Clientes Cli" & vbCrLf & _
              "    ON P.Cliente    = Cli.Cliente_Id" & vbCrLf & _
              "   AND P.EndCliente = Cli.Endereco_Id" & vbCrLf & _
              " WHERE P.Situacao      = 1" & vbCrLf

        If CrmParametros.Consolidado Then
            sql &= "   AND left(P.Empresa_id,8) = '" & CrmParametros.CodigoEmpresa.Substring(0, 8) & "'" & vbCrLf
        Else
            sql &= "   AND P.Empresa_id    ='" & CrmParametros.CodigoEmpresa & "'" & vbCrLf & _
                   "   AND P.EndEmpresa_Id = " & CrmParametros.EndEmpresa & vbCrLf
        End If

        If CrmParametros.Mercado = "I" Then
            sql &= "   AND OP.Classe = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf
        ElseIf CrmParametros.Mercado = "E" Then
            sql &= "   AND OP.Classe = '" & eClassesOperacoes.EXPORTACOES.ToString & "'" & vbCrLf
        Else
            sql &= "   AND OP.Classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "') " & vbCrLf
        End If

        sql &= "   AND P.Safra in (Select Safra_Id from safras where year(vencimento) = " & CrmParametros.Ano & ")" & vbCrLf & _
               " GROUP BY P.Cliente," & vbCrLf & _
               "		  P.EndCliente," & vbCrLf & _
               "		  Cli.Nome," & vbCrLf & _
               "		  p.Pedido_Id," & vbCrLf & _
               "		  PxI.Produto_Id," & vbCrLf & _
               "		  having Sum(case" & vbCrLf & _
               "   		 			   When PxI.TipoDeLancamento = 'E'" & vbCrLf & _
               "					     then PxI.Quantidade * - 1" & vbCrLf & _
               "					     else PxI.Quantidade" & vbCrLf & _
               "				     end) > 0;" & vbCrLf

        '******** Conta os clientes de cada analise para descubrir quais tem clientes agrupados *********
        sql &= "Select AC.Analise_Id, count(*) nr_clientes" & vbCrLf & _
               "  into #Agrupamento" & vbCrLf & _
               "  from AnaliseDeCredito AC" & vbCrLf & _
               " Inner Join AnaliseDeCreditoCliente ACC" & vbCrLf & _
               "    ON AC.Analise_Id      = ACC.Analise_Id" & vbCrLf & _
               "   and AC.Ano_Id          = ACC.Ano_Id" & vbCrLf & _
               "   and AC.DefinicaoAno_Id = ACC.DefinicaoAno_Id" & vbCrLf & _
               " Where AC.Ano_Id   =" & CrmParametros.Ano & vbCrLf & _
               "   and AC.Situacao ='LIBERADA'" & vbCrLf & _
               " Group By AC.Analise_Id" & vbCrLf

        '******** Calcula o % de margem bruta na venda agrupando os clientes conforme a analise de credito *********
        sql &= "select isnull(AC.Analise_Id,0) as Analise_Id," & vbCrLf & _
               "       isnull(AC.DefinicaoAno_Id," & CrmParametros.Definicao & ") as DefinicaoAno," & vbCrLf & _
               "       isnull(AC.LimiteCredito,'D') as LimiteCredito," & vbCrLf & _
               "       case" & vbCrLf & _
               "          when isnull(A.nr_clientes,0) > 1" & vbCrLf & _
               "            then ''" & vbCrLf & _
               "            else case" & vbCrLf & _
               "				   when len(T.Cliente) = 14" & vbCrLf & _
               "					 then left(T.Cliente,8)" & vbCrLf & _
               "					 else T.Cliente" & vbCrLf & _
               "			     end" & vbCrLf & _
               "       end as Cliente," & vbCrLf & _
               "       convert(nvarchar(max),'') as Agrupamento," & vbCrLf & _
               "       Convert(numeric(5,2),sum(T.Totaloficial) * 100 / sum(T.UnitarioOficialCompra * T.Quantidade)  -100) as PercCliente" & vbCrLf & _
               "  Into #PartCliente" & vbCrLf & _
               "  from #temp T" & vbCrLf & _
               "  left join AnaliseDeCredito AC" & vbCrLf & _
               "  left Join AnaliseDeCreditoCliente ACC" & vbCrLf & _
               "    ON AC.Analise_Id      = ACC.Analise_Id" & vbCrLf & _
               "   and AC.Ano_Id          = ACC.Ano_Id" & vbCrLf & _
               "   and AC.DefinicaoAno_Id = ACC.DefinicaoAno_Id" & vbCrLf & _
               "    on Left(T.Cliente,8)  = Left(ACC.Cliente_Id,8)" & vbCrLf & _
               "  left join #Agrupamento A" & vbCrLf & _
               "    on A.Analise_id = isnull(AC.Analise_Id,0)" & vbCrLf & _
               " Where isnull(AC.Ano_Id," & CrmParametros.Ano & ")   = " & CrmParametros.Ano & vbCrLf & _
               "   and isnull(AC.Situacao,'LIBERADA') = 'LIBERADA'" & vbCrLf & _
               "   and T.UnitarioOficialCompra > 0" & vbCrLf & _
               " group by isnull(AC.Analise_Id,0)," & vbCrLf & _
               "          isnull(AC.DefinicaoAno_Id," & CrmParametros.Definicao & ")," & vbCrLf & _
               "          isnull(AC.LimiteCredito,'D')," & vbCrLf & _
               "          case" & vbCrLf & _
               "            when isnull(A.nr_clientes,0) > 1" & vbCrLf & _
               "              then ''" & vbCrLf & _
               "              else case" & vbCrLf & _
               "				     when len(T.Cliente) = 14" & vbCrLf & _
               "					   then left(T.Cliente,8)" & vbCrLf & _
               "					   else T.Cliente" & vbCrLf & _
               "			       end" & vbCrLf & _
               "          end;" & vbCrLf

        '******* Desmembra os clientes q foram agrupado na simulacao, alimentando com os valores apurados no agrupamento
        sql &= "insert into #PartCliente(Analise_Id,DefinicaoAno,LimiteCredito, Cliente,Agrupamento,PercCliente)" & vbCrLf & _
               "(select isnull(P.Analise_id,0), isnull(ACC.DefinicaoAno_Id,0), isnull(AC.LimiteCredito,'D'), ACC.Cliente_Id, dbo.ClientesAnalise(P.Analise_id),P.PercCliente" & vbCrLf & _
               "   from #PartCliente P" & vbCrLf & _
               "  inner join AnaliseDeCredito AC" & vbCrLf & _
               "     on P.Analise_Id  =  AC.Analise_Id" & vbCrLf & _
               "  inner Join AnaliseDeCreditoCliente ACC" & vbCrLf & _
               "     ON AC.Analise_Id      = ACC.Analise_Id" & vbCrLf & _
               "    and AC.Ano_Id          = ACC.Ano_Id" & vbCrLf & _
               "    and AC.DefinicaoAno_Id = ACC.DefinicaoAno_Id" & vbCrLf & _
               "  where P.cliente   =''" & vbCrLf & _
               "    and AC.Ano_Id   =" & CrmParametros.Ano & vbCrLf & _
               "    and AC.Situacao ='LIBERADA'" & vbCrLf & _
               ");" & vbCrLf

        '**********  Apaga os Agrupamentos de Clientes  ***************
        sql &= "delete #PartCliente where Cliente = '';" & vbCrLf


        '********* Calcula o total de vendas e compras + o limite de Credito definido na analise de credito **********
        sql &= "Select isnull(PC.Analise_Id,0) as Analise_Id," & vbCrLf & _
               "       isnull(PC.DefinicaoAno," & CrmParametros.Definicao & ") as DefinicaoAno," & vbCrLf & _
               "       isnull(PC.LimiteCredito,'D') as LimiteCredito," & vbCrLf & _
               "       case " & vbCrLf & _
               "          when len(sb.Cliente_Id) = 14" & vbCrLf & _
               "            then left(Sb.Cliente_Id,8)" & vbCrLf & _
               "            else Sb.Cliente_id" & vbCrLf & _
               "       end as cliente," & vbCrLf & _
               "       isnull(PC.Agrupamento,'') as Agrupamento," & vbCrLf & _
               "       sum(sb.NotasVendas)  as NotasVendas," & vbCrLf & _
               "       sum(sb.TotalVendas)  as TotalVendas," & vbCrLf & _
               "       sum(sb.NotasCompras) as NotasCompras," & vbCrLf & _
               "       sum(sb.TotalCompras) as TotalCompras," & vbCrLf & _
               "       isnull(PC.PercCliente,0) as PercMargemBruta" & vbCrLf & _
               " Into #Pre" & vbCrLf & _
               " from (" & vbCrLf & _
               "		 SELECT NF.Cliente_Id," & vbCrLf & _
               "				NF.EndCliente_Id," & vbCrLf & _
               "				COUNT(NF.Nota_Id) AS NotasVendas," & vbCrLf & _
               "				0 as NotasCompras," & vbCrLf & _
               "				sum(case" & vbCrLf & _
               "					  when SO.devolucao = 'S'" & vbCrLf & _
               "						then NFxI.Valor * - 1" & vbCrLf & _
               "						else NFxI.Valor" & vbCrLf & _
               "					end) TotalVendas," & vbCrLf & _
               "				0.00 as TotalCompras" & vbCrLf & _
               "		  FROM NotasFiscais NF" & vbCrLf & _
               "		 INNER JOIN Pedidos P" & vbCrLf & _
               "			ON P.Pedido_Id = NF.Pedido" & vbCrLf & _
               "		 INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
               "			ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
               "		   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
               "		   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
               "		   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
               "		   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
               "		   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
               "		   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
               "		 INNER JOIN Produtos Prd" & vbCrLf & _
               "			ON NFxI.Produto_Id = Prd.Produto_Id" & vbCrLf & _
               "		 Inner Join Operacoes OP" & vbCrLf & _
               "			on op.operacao_id = NFxI.Operacao" & vbCrLf & _
               "		 INNER JOIN SubOperacoes SO" & vbCrLf & _
               "			ON NFxI.Operacao    = SO.Operacao_Id" & vbCrLf & _
               "		   AND NFxI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
               "		 WHERE NF.situacao                     = 1" & vbCrLf & _
               "		   And P.Safra in (Select Safra_Id from safras where year(vencimento) = " & CrmParametros.Ano & ")" & vbCrLf & _
               "		   AND SO.classe                      <> '" & eClassesOperacoes.GLOBAL.ToString & "'" & vbCrLf

        If CrmParametros.Consolidado Then
            sql &= "   AND left(NF.Empresa_id,8) = '" & CrmParametros.CodigoEmpresa.Substring(0, 8) & "'" & vbCrLf
        Else
            sql &= "   AND NF.Empresa_id    ='" & CrmParametros.CodigoEmpresa & "'" & vbCrLf & _
                   "   AND NF.EndEmpresa_Id = " & CrmParametros.EndEmpresa & vbCrLf
        End If

        If CrmParametros.Mercado = "I" Then
            sql &= "   AND OP.Classe = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf
        ElseIf CrmParametros.Mercado = "E" Then
            sql &= "   AND OP.Classe = '" & eClassesOperacoes.EXPORTACOES.ToString & "'" & vbCrLf
        Else
            sql &= "   AND OP.Classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "') " & vbCrLf
        End If

        '****************************************************************************
        '****************** GRUPOS E PRODUTOS VENDAS ********************************
        '****************************************************************************
        Dim xv As String() = CrmParametros.ProdutosVenda.Split(";")
        Dim sqlAuxPv As String = ""

        '*************************  GRUPOS ******************************************
        If xv(0).Length > 0 Then
            Dim xv1 As String() = xv(0).Split("|")
            Dim xvGrd As String() = xv1(1).Split(",")
            Dim grp As String = ""

            For x As Integer = 0 To xvGrd.Length - 1
                grp += IIf(grp.Length > 0, ",", "") & "'" & xvGrd(x) & "'"
            Next
            sqlAuxPv &= "  LEFT(Prd.Grupo," & xv1(0) & ") in (" & grp & ")"
        End If

        '*************************  PRODUTO ******************************************
        If xv(1).Length > 0 Then
            Dim xvPrd As String() = xv(1).Split(",")
            Dim prd As String = ""
            For x As Integer = 0 To xvPrd.Length - 1
                prd += IIf(prd.Length > 0, ",", "") & "'" & xvPrd(x) & "'"
            Next
            sql &= IIf(sqlAuxPv.Length > 0, "AND (" & sqlAuxPv & " OR NFxi.Produto_Id in (" & prd & "))", " AND NFxi.Produto_Id in (" & prd & ")")
        Else
            sql &= " AND " & sqlAuxPv
        End If

        sql &= "		 GROUP BY NF.Cliente_Id, NF.EndCliente_Id" & vbCrLf

        If CrmParametros.ConsideraCompra Then
            sql &= "	 	 UNION" & vbCrLf & _
                   "		 SELECT NF.Cliente_Id," & vbCrLf & _
                   "				NF.EndCliente_Id," & vbCrLf & _
                   "               0 AS NotasVendas," & vbCrLf & _
                   "				COUNT(NF.Nota_Id) AS NotasCompras," & vbCrLf & _
                   "               0.00 as TotalVendas," & vbCrLf & _
                   "				sum(case" & vbCrLf & _
                   "					  when SO.devolucao = 'S'" & vbCrLf & _
                   "						then NFxI.Valor * - 1" & vbCrLf & _
                   "						else NFxI.Valor" & vbCrLf & _
                   "					end) Totalcompras" & vbCrLf & _
                   "		  FROM NotasFiscais NF" & vbCrLf & _
                   "		 INNER JOIN Pedidos P" & vbCrLf & _
                   "			ON P.Pedido_Id = NF.Pedido" & vbCrLf & _
                   "		 INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
                   "			ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
                   "		   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
                   "		   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
                   "		   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
                   "		   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                   "		   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                   "		   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                   "		 INNER JOIN Produtos Prd" & vbCrLf & _
                   "			ON NFxI.Produto_Id = Prd.Produto_Id" & vbCrLf & _
                   "		 Inner Join Operacoes OP" & vbCrLf & _
                   "			on op.operacao_id = NFxI.Operacao" & vbCrLf & _
                   "		 INNER JOIN SubOperacoes SO" & vbCrLf & _
                   "			ON NFxI.Operacao    = SO.Operacao_Id" & vbCrLf & _
                   "		   AND NFxI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
                   "		 WHERE NF.situacao                     = 1" & vbCrLf & _
                   "		   And P.Safra in (Select Safra_Id from safras where year(vencimento) = " & CrmParametros.Ano & ")" & vbCrLf & _
                   "		   AND SO.classe                      <> '" & eClassesOperacoes.GLOBAL.ToString & "'" & vbCrLf

            If CrmParametros.Consolidado Then
                sql &= "   AND left(NF.Empresa_id,8) = '" & CrmParametros.CodigoEmpresa.Substring(0, 8) & "'" & vbCrLf
            Else
                If CrmParametros.CodigoEmpresaCompra.Length > 0 Then
                    sql &= "   AND NF.Empresa_id    ='" & CrmParametros.CodigoEmpresaCompra & "'" & vbCrLf & _
                           "   AND NF.EndEmpresa_Id = " & CrmParametros.EndEmpresaCompra & vbCrLf
                Else
                    sql &= "   AND NF.Empresa_id    ='" & CrmParametros.CodigoEmpresa & "'" & vbCrLf & _
                           "   AND NF.EndEmpresa_Id = " & CrmParametros.EndEmpresa & vbCrLf
                End If
            End If

            If CrmParametros.Mercado = "I" Then
                sql &= "   AND OP.Classe = '" & eClassesOperacoes.COMPRAS.ToString & "'" & vbCrLf
            ElseIf CrmParametros.Mercado = "E" Then
                sql &= "   AND OP.Classe = '" & eClassesOperacoes.IMPORTACOES.ToString & "'" & vbCrLf
            Else
                sql &= "   AND OP.Classe in ('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.IMPORTACOES.ToString & "') " & vbCrLf
            End If

            '****************************************************************************
            '****************** GRUPOS E PRODUTOS COMPRAS *******************************
            '****************************************************************************
            Dim xc As String() = CrmParametros.ProdutosCompra.Split(";")
            Dim sqlAuxPC As String = ""

            '*************************  GRUPOS ******************************************
            If xc(0).Length > 0 Then
                Dim xc1 As String() = xc(0).Split("|")
                Dim xcGrd As String() = xc1(1).Split(",")
                Dim grp As String = ""

                For x As Integer = 0 To xcGrd.Length - 1
                    grp += IIf(grp.Length > 0, ",", "") & "'" & xcGrd(x) & "'"
                Next
                sqlAuxPC &= "   LEFT(Prd.Grupo," & xc1(0) & ") in (" & grp & ")"
            End If

            '*************************  PRODUTO ******************************************
            If xc(1).Length > 0 Then
                Dim xcPrd As String() = xc(1).Split(",")
                Dim prd As String = ""
                For x As Integer = 0 To xcPrd.Length - 1
                    prd += IIf(prd.Length > 0, ",", "") & "'" & xcPrd(x) & "'"
                Next
                sql &= IIf(sqlAuxPC.Length > 0, "AND (" & sqlAuxPC & " OR NFxi.Produto_Id in (" & prd & "))", " AND NFxi.Produto_Id in (" & prd & ")")
            Else
                sql &= " AND " & sqlAuxPC
            End If

            sql &= "		 GROUP BY NF.Cliente_Id, NF.EndCliente_Id" & vbCrLf
        End If

        sql &= "   ) As SB" & vbCrLf & _
               " INNER JOIN Clientes C" & vbCrLf & _
               "    ON SB.Cliente_Id    = C.Cliente_Id" & vbCrLf & _
               "   AND SB.EndCliente_Id = C.Endereco_Id" & vbCrLf & _
               "  Left Join #PartCliente PC" & vbCrLf & _
               "    on case " & vbCrLf & _
               "        when len(PC.Cliente) = 14" & vbCrLf & _
               "          then left(PC.Cliente,8)" & vbCrLf & _
               "          else pc.Cliente" & vbCrLf & _
               "       end = case" & vbCrLf & _
               "		 	  when len(sb.Cliente_Id) = 14" & vbCrLf & _
               "			   then left(Sb.Cliente_Id,8)" & vbCrLf & _
               "			   else Sb.Cliente_id" & vbCrLf & _
               "			 end" & vbCrLf & _
               " group by isnull(PC.Analise_Id,0)," & vbCrLf & _
               "          isnull(PC.DefinicaoAno," & CrmParametros.Definicao & ")," & vbCrLf & _
               "          isnull(PC.LimiteCredito,'D')," & vbCrLf & _
               "          case" & vbCrLf & _
               "			  when len(sb.Cliente_Id) = 14" & vbCrLf & _
               "				then left(Sb.Cliente_Id,8)" & vbCrLf & _
               "				else Sb.Cliente_id" & vbCrLf & _
               "		    end," & vbCrLf & _
               "           isnull(PC.Agrupamento,'')," & vbCrLf & _
               "           isnull(PC.PercCliente,0);" & vbCrLf

        '************  Calcula os % de participacao na venda e compra e ja ordena os clientes levando em consideracao o agrupamento
        sql &= "Select sbordem.ordem," & vbCrLf & _
               "       P.Cliente," & vbCrLf & _
               "       (Select top 1 C.Nome" & vbCrLf & _
               "          from Clientes C" & vbCrLf & _
               "         where case" & vbCrLf & _
               "			     when len(C.Cliente_Id) = 14" & vbCrLf & _
               "				   then left(C.Cliente_Id,8)" & vbCrLf & _
               "				   else C.Cliente_id" & vbCrLf & _
               "		       end = P.Cliente) as NomeCliente," & vbCrLf & _
               "       P.Agrupamento," & vbCrLf & _
               "       P.Analise_Id," & vbCrLf & _
               "       P.DefinicaoAno," & vbCrLf & _
               "       P.LimiteCredito," & vbCrLf & _
               "       P.PercMargemBruta," & vbCrLf & _
               "       isnull(SB_Portifolio.PotencialVenda,0) as PotencialVenda," & vbCrLf & _
               "       P.NotasVendas," & vbCrLf & _
               "       P.TotalVendas," & vbCrLf & _
               "       case" & vbCrLf & _
               "          when isnull(SB_Portifolio.PotencialVenda,0) = 0" & vbCrLf & _
               "            then 0" & vbCrLf & _
               "            else P.TotalVendas * 100 /  isnull(SB_Portifolio.PotencialVenda,0)" & vbCrLf & _
               "       end as PercVenda," & vbCrLf & _
               "       isnull(SB_Portifolio.PotencialCompra,0) as PotencialCompra," & vbCrLf & _
               "       P.NotasCompras," & vbCrLf & _
               "       P.TotalCompras," & vbCrLf & _
               "       case" & vbCrLf & _
               "          when isnull(SB_Portifolio.PotencialCompra,0) = 0" & vbCrLf & _
               "            then 0" & vbCrLf & _
               "            else P.TotalCompras * 100 /  isnull(SB_Portifolio.PotencialCompra,0)" & vbCrLf & _
               "       end as PercCompra" & vbCrLf & _
               "  from #pre P" & vbCrLf & _
               " left join(" & vbCrLf & _
               "            select ACPC.DefinicaoAno_Id," & vbCrLf & _
               "                   case" & vbCrLf & _
               "					  when len(CS.Cliente_Id) = 14" & vbCrLf & _
               "						then left(CS.Cliente_Id,8)" & vbCrLf & _
               "						else CS.Cliente_id" & vbCrLf & _
               "				   end as Cliente," & vbCrLf & _
               "                   sum(CS.AreaPlantada * ACPC.CustoPortifolioHa) PotencialVenda," & vbCrLf & _
               "                   sum(CS.AreaPlantada * ACPC.Produtividade * ACPC.PrecoSaco) PotencialCompra" & vbCrLf & _
               "			  from clienteXSafra CS" & vbCrLf & _
               "             Inner Join AnaliseDeCreditoParametroCultura ACPC" & vbCrLf & _
               "                on ACPC.cultura_Id = CS.Cultura_id" & vbCrLf & _
               "               and ACPC.Safra_Id   = CS.Safra_Id" & vbCrLf & _
               "			 where CS.Safra_id in (Select Safra_Id from safras where year(vencimento) =" & CrmParametros.Ano & ")" & vbCrLf & _
               "               and ACPC.Ano_Id =" & CrmParametros.Ano & vbCrLf & _
               "             group by ACPC.DefinicaoAno_Id," & vbCrLf & _
               "                    case" & vbCrLf & _
               "					  when len(CS.Cliente_Id) = 14" & vbCrLf & _
               "						then left(CS.Cliente_Id,8)" & vbCrLf & _
               "						else CS.Cliente_id" & vbCrLf & _
               "				    end" & vbCrLf & _
               "            )as SB_Portifolio" & vbCrLf & _
               "    on SB_Portifolio.DefinicaoAno_Id = P.DefinicaoAno" & vbCrLf & _
               "   and SB_Portifolio.Cliente         = P.Cliente" & vbCrLf & _
               " Inner join (" & vbCrLf & _
               "			select ROW_NUMBER() over (order by sum(P.TotalVendas) desc) as ordem," & vbCrLf & _
               "				   P.analise_Id," & vbCrLf & _
               "				   case" & vbCrLf & _
               "					 when not A.analise_Id is null" & vbCrLf & _
               "					   then ''" & vbCrLf & _
               "					   else P.cliente" & vbCrLf & _
               "				   end Cliente" & vbCrLf & _
               "			 from #pre P" & vbCrLf & _
               "			 left join #Agrupamento A" & vbCrLf & _
               "			   on P.Analise_Id = A.Analise_Id" & vbCrLf & _
               "			group by P.analise_Id," & vbCrLf & _
               "					 case" & vbCrLf & _
               "					   when not A.analise_Id is null" & vbCrLf & _
               "						 then ''" & vbCrLf & _
               "						 else P.cliente" & vbCrLf & _
               "					 end" & vbCrLf & _
               "            )sbordem" & vbCrLf & _
               "   on case" & vbCrLf & _
               "         when sbordem.Cliente = ''" & vbCrLf & _
               "           then ''" & vbCrLf & _
               "           else P.Cliente" & vbCrLf & _
               "      end = sbordem.Cliente" & vbCrLf & _
               "   and case" & vbCrLf & _
               "         when sbordem.Cliente = ''" & vbCrLf & _
               "           then sbordem.analise_Id" & vbCrLf & _
               "           else 0" & vbCrLf & _
               "       end = p.Analise_Id" & vbCrLf & _
               "  order by sbordem.ordem" & vbCrLf

        Dim banco As New AcessaBanco
        Dim ds As DataSet
        ds = banco.ConsultaDataSet(sql, "CRM")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim C As New CRMxCliente(Me)
            C.IUD = "I"
            C.TipoClienteCRM = "C"
            C.Ordem = row("ordem")
            C.CodigoCliente = row("Cliente")
            C.NomeCliente = row("NomeCliente")
            C.Agrupamento = row("Agrupamento")
            C.LimiteCredito = row("LimiteCredito")
            C.PercentualMargemBruta = row("PercMargemBruta")

            C.PotencialVenda = row("PotencialVenda")
            C.NumNotasVenda = row("NotasVendas")
            C.TotalVendas = row("TotalVendas")
            C.PercentualVendas = row("PercVenda")

            C.PotencialCompra = row("PotencialCompra")
            C.NumNotasCompra = row("NotasCompras")
            C.TotalCompras = row("TotalCompras")
            C.PercentualCompras = row("PercCompra")
            Me.Add(C)
        Next
    End Sub

    Public Sub CarregarSalva()
        Dim sql As String = ""
        sql = "SELECT CxC.Empresa_Id, CxC.EndEmpresa_Id, CxC.Ano_Id, CxC.Consolidado_Id, CxC.Cliente_Id, " & vbCrLf & _
              "       (Select top 1 C.Nome" & vbCrLf & _
              "          from Clientes C" & vbCrLf & _
              "         where case" & vbCrLf & _
              "			        when len(C.Cliente_Id) = 14" & vbCrLf & _
              "				      then left(C.Cliente_Id,8)" & vbCrLf & _
              "				      else C.Cliente_id" & vbCrLf & _
              "		          end = CxC.Cliente_Id) as NomeCliente," & vbCrLf & _
              "       CxC.Ordem, CxC.TipoCliente, CxC.TipoClienteQualitativo, CxC.Agrupamento, CxC.LimiteCredito, CxC.PercentualMargemBruta," & vbCrLf & _
              "       CxC.PotencialVenda, CxC.NumNotasVenda, CxC.TotalVendas, CxC.PercentualVendas," & vbCrLf & _
              "       CxC.PotencialCompra, CxC.NumNotasCompra, CxC.TotalCompras, CxC.PercentualCompras," & vbCrLf & _
              "       CxC.PontuacaoMBV, CxC.PontuacaoSOC, CxC.PontuacaoRating," & vbCrLf & _
              "       CxC.Classificacao" & vbCrLf & _
              "  FROM CRMParametroXCliente CxC" & vbCrLf & _
              " WHERE CxC.Empresa_Id     ='" & CrmParametros.CodigoEmpresa & "'" & vbCrLf & _
              "   AND CxC.EndEmpresa_Id  = " & CrmParametros.EndEmpresa & vbCrLf & _
              "   AND CxC.Ano_Id         = " & CrmParametros.Ano & vbCrLf & _
              "   AND CxC.Consolidado_Id = " & IIf(CrmParametros.Consolidado, "1", "0") & vbCrLf & _
              " Order by CxC.Ordem"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "CRMxCliente")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim C As New CRMxCliente(Me)
            C.IUD = "U"

            C.Ordem = row("ordem")
            C.TipoClienteCRM = row("TipoCliente")
            C.TipoClienteQualitativo = row("TipoClienteQualitativo")
            C.Agrupamento = row("Agrupamento")

            C.CodigoCliente = row("Cliente_Id")
            C.NomeCliente = row("NomeCliente")

            C.LimiteCredito = row("LimiteCredito")
            C.PercentualMargemBruta = row("PercentualMargemBruta")

            C.PotencialVenda = row("PotencialVenda")
            C.NumNotasVenda = row("NumNotasVenda")
            C.TotalVendas = row("TotalVendas")
            C.PercentualVendas = row("PercentualVendas")

            C.PotencialCompra = row("PotencialCompra")
            C.NumNotasCompra = row("NumNotasCompra")
            C.TotalCompras = row("TotalCompras")
            C.PercentualCompras = row("PercentualCompras")

            C.PontuacaoMBV = row("PontuacaoMBV")
            C.PontuacaoRating = row("PontuacaoRating")
            C.PontuacaoSOC = row("PontuacaoSOC")
            Me.Add(C)
        Next

    End Sub

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each row In Me
            If CrmParametros.IUD = "D" Or CrmParametros.IUD = "I" Then row.IUD = CrmParametros.IUD

            If row.IUD <> "" Then
                row.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class CRMxCliente
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pLista As ListCRMxCliente)
        _Lista = pLista
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Lista As ListCRMxCliente

    Private _Parametros As CRM

    Private _Ordem As Integer
    Private _TipoClienteCRM As String = "V"
    Private _TipoClienteQualitativo As String = "V"

    Private _CodigoCliente As String = ""
    Private _NomeCliente As String = ""
    Private _Agrupamento As String = ""

    '******* Limite de Credito ********
    Private _LimiteCredito As String = ""

    '******* Margem Bruta ********
    Private _PercentualMargemBruta As Decimal

    '********** VENDAS ***********
    Private _PotencialVenda As Decimal
    Private _NumNotasVenda As Integer
    Private _TotalVendas As Decimal
    Private _PercentualVendas As Decimal

    '********** COMPRAS ***********
    Private _PotencialCompra As Decimal
    Private _NumNotasCompra As Integer
    Private _TotalCompras As Decimal
    Private _PercentualCompras As Decimal

    '************** Pontuacao e Classificacao **************
    Private _PontuacaoMBV As Decimal
    Private _PontuacaoSOC As Decimal
    Private _PontuacaoRating As Decimal
    Private _Classificacao As String

    '********** Percentual Faturamento Acumulado ***********
    Private _PercentualFaturamentoAcumulado As Decimal
#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public ReadOnly Property Lista() As ListCRMxCliente
        Get
            Return _Lista
        End Get
    End Property

    Public Property Parametros() As CRM
        Get
            Return _Parametros
        End Get
        Set(ByVal value As CRM)
            _Parametros = value
        End Set
    End Property

    Public Property Ordem() As Integer
        Get
            Return _Ordem
        End Get
        Set(ByVal value As Integer)
            _Ordem = value
        End Set
    End Property

    Public Property TipoClienteCRM() As String
        Get
            Return _TipoClienteCRM
        End Get
        Set(ByVal value As String)
            _TipoClienteCRM = value
        End Set
    End Property

    Public Property TipoClienteQualitativo() As String
        Get
            Return _TipoClienteQualitativo
        End Get
        Set(ByVal value As String)
            _TipoClienteQualitativo = value
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
        End Set
    End Property

    Public Property NomeCliente() As String
        Get
            Return _NomeCliente
        End Get
        Set(ByVal value As String)
            _NomeCliente = value
        End Set
    End Property

    Public Property Agrupamento() As String
        Get
            Return _Agrupamento
        End Get
        Set(ByVal value As String)
            _Agrupamento = value
        End Set
    End Property

    '******* Limite de Credito ********
    Public Property LimiteCredito() As String
        Get
            Return _LimiteCredito
        End Get
        Set(ByVal value As String)
            _LimiteCredito = value
        End Set
    End Property

    '******* Margem Bruta ********
    Public Property PercentualMargemBruta() As Decimal
        Get
            Return _PercentualMargemBruta
        End Get
        Set(ByVal value As Decimal)
            _PercentualMargemBruta = value
        End Set
    End Property

    '********** VENDAS ***********
    Public Property PotencialVenda() As Decimal
        Get
            Return _PotencialVenda
        End Get
        Set(ByVal value As Decimal)
            _PotencialVenda = value
        End Set
    End Property

    Public Property NumNotasVenda() As Integer
        Get
            Return _NumNotasVenda
        End Get
        Set(ByVal value As Integer)
            _NumNotasVenda = value
        End Set
    End Property

    Public Property TotalVendas() As Decimal
        Get
            Return _TotalVendas
        End Get
        Set(ByVal value As Decimal)
            _TotalVendas = value
        End Set
    End Property

    Public Property PercentualVendas() As Decimal
        Get
            Return _PercentualVendas
        End Get
        Set(ByVal value As Decimal)
            _PercentualVendas = value
        End Set
    End Property

    '********** COMPRAS ***********
    Public Property PotencialCompra() As Decimal
        Get
            Return _PotencialCompra
        End Get
        Set(ByVal value As Decimal)
            _PotencialCompra = value
        End Set
    End Property

    Public Property NumNotasCompra() As Integer
        Get
            Return _NumNotasCompra
        End Get
        Set(ByVal value As Integer)
            _NumNotasCompra = value
        End Set
    End Property

    Public Property TotalCompras() As Decimal
        Get
            Return _TotalCompras
        End Get
        Set(ByVal value As Decimal)
            _TotalCompras = value
        End Set
    End Property

    Public Property PercentualCompras() As Decimal
        Get
            Return _PercentualCompras
        End Get
        Set(ByVal value As Decimal)
            _PercentualCompras = value
        End Set
    End Property

    '************** Pontuacao e Classificacao **************
    Public Property PontuacaoMBV() As Decimal
        Get
            Return _PontuacaoMBV
        End Get
        Set(ByVal value As Decimal)
            _PontuacaoMBV = value
        End Set
    End Property

    Public Property PontuacaoSOC() As Decimal
        Get
            Return _PontuacaoSOC
        End Get
        Set(ByVal value As Decimal)
            _PontuacaoSOC = value
        End Set
    End Property

    Public Property PontuacaoRating() As Decimal
        Get
            Return _PontuacaoRating
        End Get
        Set(ByVal value As Decimal)
            _PontuacaoRating = value
        End Set
    End Property

    Public ReadOnly Property Classificacao() As String
        Get
            If TipoClienteCRM = "V" Then Return ""

            Select Case PontuacaoMBV + PontuacaoSOC + PontuacaoRating
                Case Is <= 18 : Return "BRONZE"
                Case Is >= 24 : Return "OURO"
                Case Else : Return "PRATA"
            End Select
        End Get
    End Property

    '********** Percentual Faturamento Acumulado ***********
    Public Property PercentualFaturamentoAcumulado() As Decimal
        Get
            Return _PercentualFaturamentoAcumulado
        End Get
        Set(ByVal value As Decimal)
            _PercentualFaturamentoAcumulado = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSQL As String = ""
        Select Case Me.IUD
            Case "I"
                strSQL &= "insert into CRMParametroXCliente(Empresa_Id, EndEmpresa_Id, Ano_Id, Consolidado_Id, Cliente_Id, Ordem, TipoCliente, TipoClienteQualitativo, " & vbCrLf & _
                          "                                 Agrupamento, LimiteCredito, PercentualMargemBruta, PotencialVenda, NumNotasVenda, TotalVendas, " & vbCrLf & _
                          "                                 PercentualVendas, PotencialCompra, NumNotasCompra, TotalCompras, PercentualCompras, " & vbCrLf & _
                          "                                 PontuacaoMBV, PontuacaoSOC, PontuacaoRating, Classificacao)" & vbCrLf & _
                          "  values('" & Me.Lista.CrmParametros.CodigoEmpresa & "'," & Me.Lista.CrmParametros.EndEmpresa & "," & Me.Lista.CrmParametros.Ano & "," & IIf(Me.Lista.CrmParametros.Consolidado, "1", "0") & ",'" & Me.CodigoCliente & "'," & Me.Ordem & ",'" & Me.TipoClienteCRM & "','" & Me.TipoClienteQualitativo & "'," & vbCrLf & _
                          "'" & Me.Agrupamento & "','" & Me.LimiteCredito & "'," & Str(Me.PercentualMargemBruta) & "," & Str(Me.PotencialVenda) & "," & Str(Me.NumNotasVenda) & "," & Str(Me.TotalVendas) & "," & vbCrLf & _
                          Str(Math.Round(Me.PercentualVendas, 2)) & "," & Str(Me.PotencialCompra) & "," & Str(Me.NumNotasCompra) & "," & Str(Me.TotalCompras) & "," & Str(Math.Round(Me.PercentualCompras, 2)) & "," & vbCrLf & _
                          Str(Me.PontuacaoMBV) & "," & Str(Me.PontuacaoSOC) & "," & Str(Me.PontuacaoRating) & ",'" & Me.Classificacao & "')"
                Sqls.Add(strSQL)
            Case "U"
                strSQL = " Update CRMParametroXCliente set " & vbCrLf & _
                         "    Ordem                   = " & Me.Ordem & vbCrLf & _
                         "   ,TipoCliente             ='" & Me.TipoClienteCRM & "'" & vbCrLf & _
                         "   ,TipoClienteQualitativo  ='" & Me.TipoClienteQualitativo & "'" & vbCrLf & _
                         "   ,Agrupamento             ='" & Me.Agrupamento & "'" & vbCrLf & _
                         "   ,LimiteCredito           ='" & Me.LimiteCredito & "'" & vbCrLf & _
                         "   ,PercentualMargemBruta   = " & Str(Me.PercentualMargemBruta) & vbCrLf & _
                         "   ,PotencialVenda          = " & Str(Me.PotencialVenda) & vbCrLf & _
                         "   ,NumNotasVenda           = " & Str(Me.NumNotasVenda) & vbCrLf & _
                         "   ,TotalVendas             = " & Str(Me.TotalVendas) & vbCrLf & _
                         "   ,PercentualVendas        = " & Str(Math.Round(Me.PercentualVendas, 2)) & vbCrLf & _
                         "   ,PotencialCompra         = " & Str(Me.PotencialCompra) & vbCrLf & _
                         "   ,NumNotasCompra          = " & Str(Me.NumNotasCompra) & vbCrLf & _
                         "   ,TotalCompras            = " & Str(Me.TotalCompras) & vbCrLf & _
                         "   ,PercentualCompras       = " & Str(Math.Round(Me.PercentualCompras, 2)) & vbCrLf & _
                         "   ,PontuacaoMBV            = " & Str(Me.PontuacaoMBV) & vbCrLf & _
                         "   ,PontuacaoSOC            = " & Str(Me.PontuacaoSOC) & vbCrLf & _
                         "   ,PontuacaoRating         = " & Str(Me.PontuacaoRating) & vbCrLf & _
                         "   ,Classificacao           ='" & Me.Classificacao & "'" & vbCrLf & _
                         " Where Empresa_Id     ='" & Me.Lista.CrmParametros.CodigoEmpresa & "'" & vbCrLf & _
                         "   and EndEmpresa_Id  = " & Me.Lista.CrmParametros.EndEmpresa & vbCrLf & _
                         "   and Ano_Id         = " & Me.Lista.CrmParametros.Ano & vbCrLf & _
                         "   and Consolidado_Id = " & IIf(Me.Lista.CrmParametros.Consolidado, "1", "0") & vbCrLf & _
                         "   and Cliente_Id     ='" & Me.CodigoCliente & "'"
                Sqls.Add(strSQL)
            Case "D"
                strSQL = " Delete CRMParametroXCliente" & vbCrLf & _
                         "  Where Empresa_Id     ='" & Me.Lista.CrmParametros.CodigoEmpresa & "'" & vbCrLf & _
                         "    and EndEmpresa_Id  = " & Me.Lista.CrmParametros.EndEmpresa & vbCrLf & _
                         "    and Ano_Id         = " & Me.Lista.CrmParametros.Ano & vbCrLf & _
                         "    and Consolidado_Id = " & IIf(Me.Lista.CrmParametros.Consolidado, "1", "0") & vbCrLf & _
                         "    and Cliente_Id     ='" & Me.CodigoCliente & "'"
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class
