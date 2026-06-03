Imports System.Data
Imports System.IO
Imports System.Web
Imports System.Web.UI
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class NotaFiscalEspelho

    Public Sub ExibirEspelho(ByVal page As System.Web.UI.Page, ByVal Empresa As String, ByVal endEmpresa As Integer, ByVal Cliente As String, ByVal endCliente As Integer, ByVal EntradaSaida As String, ByVal nota As Integer, ByVal serie As String)
        Dim nf As New NotaFiscal
        nf.CodigoEmpresa = Empresa
        nf.EnderecoEmpresa = endEmpresa
        nf.CodigoCliente = Cliente
        nf.EnderecoCliente = endCliente
        nf.EntradaSaida = IIf(EntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
        nf.Codigo = nota
        nf.Serie = serie
        ExibirEspelho(page, nf)
    End Sub

    Public Sub ExibirEspelho(ByVal page As System.Web.UI.Page, ByVal nf As NotaFiscal)
        Dim sql As String
        sql = "SELECT NF.UsuarioInclusao," & vbCrLf & _
              "       NF.UsuarioInclusaoData," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then NF.Empresa_Id       else NF.Cliente_Id       end Emitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then NF.EndEmpresa_Id    else NF.EndCliente_Id    end EndEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Empresa.Nome        else Cliente.Nome        end NomeEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Empresa.Cidade      else Cliente.Cidade      end CidadeEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Empresa.Estado      else Cliente.Estado      end EstadoEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Empresa.Complemento else Cliente.Complemento end ComplementoEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Empresa.Inscricao   else Cliente.Inscricao   end InscricaoEmitente," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Cliente.Cliente_Id  else NF.Empresa_Id       end Destinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Cliente.Endereco_Id else NF.EndEmpresa_Id    end EndDestinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Cliente.Nome        else Empresa.Nome        end NomeDestinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Cliente.Cidade      else Empresa.Cidade      end CidadeDestinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Cliente.Estado      else Empresa.Estado      end EstadoDestinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Cliente.Complemento else Empresa.Complemento end ComplementoDestinatario," & vbCrLf & _
              "       case When NF.EntradaSaida_Id = 'S' then Cliente.Inscricao   else Empresa.Inscricao   end InscricaoDestinatario," & vbCrLf & _
              "       NF.EntradaSaida_Id," & vbCrLf & _
              "       NF.Serie_Id," & vbCrLf & _
              "       NF.Nota_Id," & vbCrLf & _
              "       NF.Pedido," & vbCrLf & _
              "       NF.Operacao," & vbCrLf & _
              "       NF.SubOperacao," & vbCrLf & _
              "       SO.Descricao as DescSubOperacao," & vbCrLf & _
              "       NF.Movimento," & vbCrLf & _
              "       NF.DataDaNota," & vbCrLf & _
              "       NFxI.Produto_Id," & vbCrLf & _
              "       P.Unidade," & vbCrLf & _
              "       P.Nome + ' (NCM '+ NCM +')' as NomeProduto," & vbCrLf & _
              "       NFxI.CFOP_Id," & vbCrLf & _
              "       NFxI.Sequencia_Id," & vbCrLf & _
              "       NFxI.QuantidadeFiscal," & vbCrLf & _
              "       NFxI.Unitario," & vbCrLf & _
              "       NFxI.Valor as ValorItem," & vbCrLf & _
              "       NFxE.Encargo_Id," & vbCrLf & _
              "       OE.STICMS as SituacaoTributaria," & vbCrLf & _
              "       ST.Descricao as DescTributaria," & vbCrLf & _
              "       OE.STPISCOFINS as SituacaoTributariaPISCOFINS," & vbCrLf & _
              "       StPC.Descricao as DescTributariaPisCofins," & vbCrLf & _
              "       NFxE.Base," & vbCrLf & _
              "       NFxE.Percentual," & vbCrLf & _
              "       NFxE.Valor as ValorEncargo," & vbCrLf & _
              "       (P.Grupo + ' - ' + (Select Descricao from gruposdeestoques where grupo_Id = left(P.grupo,1)) + ' | ' + (Select Descricao from gruposdeestoques where grupo_Id = left(P.grupo,2)) + ' | ' + (Select Descricao from gruposdeestoques where grupo_Id = left(P.grupo,3)) + ' | ' + (Select Descricao from gruposdeestoques where grupo_Id = left(P.grupo,5)) " & vbCrLf & _
              "        ) as DescGrupoProduto, " & vbCrLf & _
              "       isnull(NF.nfg, 0) as nfg, " & vbCrLf & _
              "       Case when isnull(NF.NFG,0) = 1 " & vbCrLf & _
              "       	then NF.Observacoes" & vbCrLf & _
              "       	else ''" & vbCrLf & _
              "       end AS Observacao," & vbCrLf & _
              "       case when NF.nossaemissao = 'S' then 'SIM' else 'NAO' end NossaEmissao," & vbCrLf & _
              "       case when NF.eletronica   = 'S' then 'SIM' else 'NAO' end Eletronica," & vbCrLf & _
              "       isnull(NFE.chavenfe,'-') as chavenfe," & vbCrLf & _
              "       nfxi.OperacaoXEstado as ConfOperacao" & vbCrLf & _
              "  FROM NotasFiscais NF" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
              "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
              "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
              " Inner join OperacaoXEstado OE" & vbCrLf & _
              "    on OE.Codigo_Id = nfxi.OperacaoXEstado" & vbCrLf & _
              " INNER JOIN NotasFiscaisXEncargos NFxE" & vbCrLf & _
              "    ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf & _
              "   AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf & _
              "   AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf & _
              "   AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf & _
              "   AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf & _
              "   AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf & _
              "   AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf & _
              "   AND NFxI.Produto_Id      = NFxE.Produto_Id" & vbCrLf & _
              "   AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf & _
              "   AND NFxI.Sequencia_Id    = NFxE.Sequencia_id" & vbCrLf & _
              "  LEFT JOIN nferealizadas NFE" & vbCrLf & _
              "    ON NFE.Empresa_Id      = NF.Empresa_Id" & vbCrLf & _
              "   AND NFE.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf & _
              "   AND NFE.Cliente_Id      = NF.Cliente_Id" & vbCrLf & _
              "   AND NFE.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf & _
              "   AND NFE.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf & _
              "   AND NFE.Serie_Id        = NF.Serie_Id" & vbCrLf & _
              "   AND NFE.Nota_Id         = NF.Nota_Id" & vbCrLf & _
              " INNER JOIN Clientes Empresa" & vbCrLf & _
              "    on Empresa.cliente_id  = NF.Empresa_Id" & vbCrLf & _
              "   and Empresa.Endereco_Id = NF.EndEmpresa_Id" & vbCrLf & _
              " INNER JOIN Clientes Cliente" & vbCrLf & _
              "    ON Cliente.Cliente_Id  = CASE WHEN NF.TipoDeDocumentoFrete = 1 THEN NF.Deposito ELSE NF.Cliente_Id END " & vbCrLf & _
              "   and Cliente.Endereco_Id = CASE WHEN NF.TipoDeDocumentoFrete = 1 THEN NF.EndDeposito ELSE NF.EndCliente_Id END " & vbCrLf & _
              " INNER JOIN SubOperacoes SO" & vbCrLf & _
              "    on SO.Operacao_id     = NF.Operacao" & vbCrLf & _
              "   and SO.SubOperacoes_id = NF.SubOperacao" & vbCrLf & _
              " INNER JOIN Produtos P" & vbCrLf & _
              "    ON P.Produto_id = NFxI.Produto_Id" & vbCrLf & _
              "  LEFT JOIN SituacaoTributaria ST" & vbCrLf & _
              "    ON ST.SituacaoTributaria_id =  OE.STICMS" & vbCrLf & _
              "  LEFT join SituacaoTributariaPISCOFINS StPC" & vbCrLf & _
              "    ON STPC.SituacaoTributariaPISCOFINS_id = OE.STPISCOFINS" & vbCrLf & _
              " where NF.Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf & _
              "   and NF.EndEmpresa_id   = " & nf.EnderecoEmpresa & vbCrLf & _
              "   and NF.Cliente_id      ='" & nf.CodigoCliente & "'" & vbCrLf & _
              "   and NF.EndCliente_id   = " & nf.EnderecoCliente & vbCrLf & _
              "   and NF.EntradaSaida_id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "   and NF.Nota_id         = " & nf.Codigo & vbCrLf & _
              "   and NF.Serie_Id        ='" & nf.Serie & "'" & vbCrLf & _
              " order by NF.Operacao, NF.SubOperacao, NFxI.Produto_Id, " & vbCrLf & _
              "            case" & vbCrLf & _
              "            When NFxE.Encargo_id = 'PRODUTO'" & vbCrLf & _
              "              then 1" & vbCrLf & _
              "            When NFxE.Encargo_id = 'LIQUIDO'" & vbCrLf & _
              "              then 3" & vbCrLf & _
              "            Else 2" & vbCrLf & _
              "          end, NFxE.Encargo_id" & vbCrLf

        Dim db As New AcessaBanco
        Dim ds As DataSet
        ds = db.ConsultaDataSet(sql, "Espelho")

        sql = "select r.Conta_id, Pc.Titulo, r.Produto, isnull(cc.CentroDeCusto_id,'') as CentroDeCusto, isnull(cc.Descricao,'') as DescCusto, r.DebitoOficial, r.CreditoOficial, r.Historico," & vbCrLf & _
              "       isnull(r.produto_NF,'') as produto_NF, isnull(OE.CodigoFiscal,0) as cfop_Nf, isnull(r.Sequencia_NF,0) as Sequencia_NF, isnull(Encargo_NF,'') as Encargo_NF  " & vbCrLf & _
              "  from razao r" & vbCrLf & _
              " Inner Join PlanoDeContas PC" & vbCrLf & _
              "    on r.conta_id = PC.Conta_id" & vbCrLf & _
              "  left join CentrosDeCustos cc" & vbCrLf & _
              "    on cc.CentroDeCusto_id  = r.custo" & vbCrLf & _
              "  Left Join NotasFiscaisXItens nfi" & vbCrLf & _
              "    on nfi.Empresa_Id      = r.Empresa_Id" & vbCrLf & _
              "   and nfi.EndEmpresa_Id   = r.EndEmpresa_Id" & vbCrLf & _
              "   and nfi.Cliente_Id      = r.Cliente_Nf" & vbCrLf & _
              "   and nfi.EndCliente_Id   = r.EndCliente_Nf" & vbCrLf & _
              "   and nfi.EntradaSaida_Id = r.EntradaSaida_Nf" & vbCrLf & _
              "   and nfi.Nota_Id         = r.Numero_Nf" & vbCrLf & _
              "   and nfi.Serie_Id        = r.Serie_Nf" & vbCrLf & _
              "   and nfi.Produto_Id      = r.Produto_NF" & vbCrLf & _
              "   and nfi.Sequencia_Id    = r.Sequencia_NF" & vbCrLf & _
              "  Left Join OperacaoXEstado OE" & vbCrLf & _
              "    on OE.Codigo_Id = nfi.OperacaoXEstado" & vbCrLf & _
              " Where r.Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf & _
              "   and r.EndEmpresa_id   = " & nf.EnderecoEmpresa & vbCrLf & _
              "   and r.Cliente_Nf      ='" & nf.CodigoCliente & "'" & vbCrLf & _
              "   and r.EndCliente_Nf   = " & nf.EnderecoCliente & vbCrLf & _
              "   and r.EntradaSaida_Nf ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "   and r.Numero_Nf       = " & nf.Codigo & vbCrLf & _
              "   and r.Serie_Nf        ='" & nf.Serie & "'"
        ds.Merge(db.ConsultaDataSet(sql, "Razao"))

        sql = "select r.Conta_id, Pc.Titulo, isnull(cc.CentroDeCusto_id,'') as CentroDeCusto, isnull(cc.Descricao,'') as DescCusto, r.DebitoOficial, r.CreditoOficial, r.Historico" & vbCrLf & _
              "  from razao r" & vbCrLf & _
              " Inner Join PlanoDeContas PC" & vbCrLf & _
              "    on r.conta_id = PC.Conta_id" & vbCrLf & _
              "  left join CentrosDeCustos cc" & vbCrLf & _
              "    on cc.CentroDeCusto_id  = r.custo" & vbCrLf & _
              " where r.Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf & _
              "   and r.EndEmpresa_id   = " & nf.EnderecoEmpresa & vbCrLf & _
              "   and r.Cliente_Nf      ='" & nf.CodigoCliente & "'" & vbCrLf & _
              "   and r.EndCliente_Nf   = " & nf.EnderecoCliente & vbCrLf & _
              "   and r.EntradaSaida_Nf ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "   and r.Numero_Nf       = " & nf.Codigo & vbCrLf & _
              "   and r.Serie_Nf        ='" & nf.Serie & "'" & vbCrLf & _
              "   And (isnull(r.Produto_NF, '0') = '0' or r.produto_NF = '')"
        ds.Merge(db.ConsultaDataSet(sql, "RazaoSemProduto"))

        sql = "select case when Encargo_Id = 'PRODUTO' then 'VALOR BRUTO' else Encargo_id end as Descricao," & vbCrLf & _
              "       sum(Valor) as Valor" & vbCrLf & _
              "  from NotasfiscaisXEncargos" & vbCrLf & _
              " where Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf & _
              "   and EndEmpresa_id   = " & nf.EnderecoEmpresa & vbCrLf & _
              "   and Cliente_id      ='" & nf.CodigoCliente & "'" & vbCrLf & _
              "   and EndCliente_id   = " & nf.EnderecoCliente & vbCrLf & _
              "   and EntradaSaida_id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "   and Nota_id         = " & nf.Codigo & vbCrLf & _
              "   and Serie_Id        ='" & nf.Serie & "'" & vbCrLf & _
              " group by Encargo_Id" & vbCrLf & _
              " order by case" & vbCrLf & _
              "            When Encargo_id = 'PRODUTO'" & vbCrLf & _
              "              then 1" & vbCrLf & _
              "            When Encargo_id = 'LIQUIDO'" & vbCrLf & _
              "              then 3" & vbCrLf & _
              "            Else 2" & vbCrLf & _
              "          end, Encargo_id" & vbCrLf
        ds.Merge(db.ConsultaDataSet(sql, "Resumo"))

        'Vencimentos da Nota Fiscal
        sql = "SELECT T.Registro_Id Titulo, Pr.Descricao as Situacao, T.Prorrogacao Vencimento, " & vbCrLf & _
              "       T.ValorDoDocumento, T.Descontos, T.Deducoes, T.Juros, T.Acrescimos, T.ValorLiquido" & vbCrLf & _
              "  FROM NotaFiscalXTitulo NFxT" & vbCrLf & _
              "  LEFT JOIN (" & vbCrLf & _
              "             SELECT CP.Registro_Id, CP.Provisao, CP.Prorrogacao, " & vbCrLf & _
              "                    CP.ValorDoDocumento, CP.Descontos, CP.Deducoes, CP.Juros, CP.Acrescimos, CP.ValorLiquido " & vbCrLf & _
              "               FROM ContasAPagar CP" & vbCrLf & _
              "              Where cp.situacao = 1" & vbCrLf & _
              "              UNION ALL" & vbCrLf & _
              "             SELECT CR.Registro_Id, CR.Provisao, CR.Prorrogacao," & vbCrLf & _
              "                    CR.ValorDoDocumento, CR.Descontos, CR.Deducoes, CR.Juros, CR.Acrescimos, CR.ValorLiquido " & vbCrLf & _
              "               FROM ContasAReceber CR" & vbCrLf & _
              "              Where cR.situacao = 1" & vbCrLf & _
              "            ) AS T" & vbCrLf & _
              "     ON NFxT.Titulo_Id = T.Registro_id" & vbCrLf & _
              "   JOIN Provisoes Pr" & vbCrLf & _
              "     ON T.Provisao = Pr.Provisao_Id" & vbCrLf & _
              "  WHERE Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf & _
              "    AND EndEmpresa_id   = " & nf.EnderecoEmpresa & vbCrLf & _
              "    AND Cliente_id      ='" & nf.CodigoCliente & "'" & vbCrLf & _
              "    AND EndCliente_id   = " & nf.EnderecoCliente & vbCrLf & _
              "    AND EntradaSaida_id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "    AND Serie_id        ='" & nf.Serie & "'" & vbCrLf & _
              "    AND Nota_id         = " & nf.Codigo & vbCrLf

        ds.Merge(db.ConsultaDataSet(sql, "VencimentosNota"))

        Funcoes.BindReport(page, ds, "Cr_EspelhoNota", eExportType.PDF)
    End Sub

End Class