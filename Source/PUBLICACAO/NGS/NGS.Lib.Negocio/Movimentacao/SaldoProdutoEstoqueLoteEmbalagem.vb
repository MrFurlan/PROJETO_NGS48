Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListSaldoProdutoEstoqueLoteEmbalagem
    Inherits List(Of SaldoProdutoEstoqueLoteEmbalagem)

#Region "Construtor"
    Public Sub New(ByVal pProduto As String)
        _CodigoProduto = pProduto
    End Sub
#End Region

#Region "Fields"
    Private _CodigoProduto As String = ""
#End Region

#Region "Property"
    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public ReadOnly Property QtdeDeEmbalagens() As Integer
        Get
            Dim qtde As Integer = 0
            For Each row As SaldoProdutoEstoqueLoteEmbalagem In Me
                qtde += row.QtdeEmbalagem
            Next
            Return qtde
        End Get
    End Property

    Public ReadOnly Property QtdeDeProduto() As Decimal
        Get
            Dim qtde As Decimal = 0
            For Each row As SaldoProdutoEstoqueLoteEmbalagem In Me
                qtde += row.QtdeEmbXCapacidade
            Next
            Return qtde
        End Get
    End Property

    Public ReadOnly Property SaldoFisico() As Decimal
        Get
            Dim SFisico As Decimal = 0
            For Each row As SaldoProdutoEstoqueLoteEmbalagem In Me
                SFisico += row.SaldoFisico
            Next
            Return SFisico
        End Get
    End Property

    Public ReadOnly Property SaldoFiscal() As Decimal
        Get
            Dim SFiscal As Decimal = 0
            For Each row As SaldoProdutoEstoqueLoteEmbalagem In Me
                SFiscal += row.SaldoFiscal
            Next
            Return SFiscal
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub CarregarResumoSaldoEmEstoque(ByVal pCodigoEmpresa As String, ByVal pEndEmpresa As Integer, pSO As SubOperacao)
        Dim sql As String
        Dim andtipoOperacao As String = ""

        If pSO.Consignacao Then
            andtipoOperacao = " and isnull(SO.consignacao,0) = 1" & vbCrLf

            If (pSO.EntradaSaida = eEntradaSaida.Entrada And pSO.Devolucao = False) Or (pSO.EntradaSaida = eEntradaSaida.Saida And pSO.Devolucao = True) Then
                'Se for uma entrada que nao é uma devolucao ou se for uma saida que for devolucao entao estamos pegando e devolvendo produto consignado
                andtipoOperacao = " and ((SO.EntradaSaida = 'E' and SO.devolucao = 'N') or (SO.EntradaSaida = 'S' and SO.devolucao = 'S')) " & vbCrLf
            Else
                'Agora se for uma entrada que é devolucao ou se for uma saida que nao é devolucao entao estamos mandando consignado e recebendo devolucao de produtos consignados
                andtipoOperacao = " and ((SO.EntradaSaida = 'E' and SO.devolucao = 'S') or (SO.EntradaSaida = 'S' and SO.devolucao = 'N')) " & vbCrLf
            End If

        ElseIf pSO.AmostraGratis Then
            andtipoOperacao = " and isnull(SO.AmostraGratis,0) = 1" & vbCrLf
        Else
            andtipoOperacao = " and isnull(SO.consignacao,0)   = 0" & vbCrLf & _
                              " and isnull(SO.AmostraGratis,0) = 0" & vbCrLf & _
                              "  or (isnull(SO.consignacao,0)  = 1 " & vbCrLf & _
                              "       and ( " & vbCrLf & _
                              "                (SO.EntradaSaida = 'E' and SO.devolucao = 'S') " & vbCrLf & _
                              "              OR " & vbCrLf & _
                              "                (SO.EntradaSaida = 'S' and SO.devolucao = 'N')" & vbCrLf & _
                              "            ) " & vbCrLf & _
                              "     )" & vbCrLf
        End If

        sql = "Select sb.Produto_Id as CodigoProduto," & vbCrLf &
              "       P.Nome as NomeProduto," & vbCrLf &
              "       P.Unidade," & vbCrLf &
              "       sum(isnull(sb.QtdeFisica,0)) SaldoFisico, " & vbCrLf &
              "       sum(isnull(sb.QtdeFiscal,0)) SaldoFiscal" & vbCrLf &
              " From " & vbCrLf &
              "	   (" & vbCrLf &
              "		SELECT NotasFiscais.Empresa_id, NotasFiscais.EndEmpresa_id, NotasFiscaisXItens.Produto_Id," & vbCrLf &
              "			   sum(case" & vbCrLf &
              "					 when SO.EntradaSaida = 'E'" & vbCrLf &
              "					  then CASE WHEN PRO.ControlarRomaneio = 'false' THEN NotasFiscaisXItens.QuantidadeFisica ELSE ROM.PesoLiquido END " & vbCrLf &
              "					  else CASE WHEN PRO.ControlarRomaneio = 'false' THEN NotasFiscaisXItens.QuantidadeFisica ELSE ROM.PesoLiquido END * - 1 " & vbCrLf &
              "				   end) as QtdeFisica," & vbCrLf &
              "			   sum(case" & vbCrLf &
              "					 when SO.EntradaSaida = 'E'" & vbCrLf &
              "					  then NotasFiscaisXItens.QuantidadeFiscal " & vbCrLf &
              "					  else  NotasFiscaisXItens.QuantidadeFiscal * - 1" & vbCrLf &
              "				   end) as QtdeFiscal" & vbCrLf &
              "		  FROM NotasFiscais" & vbCrLf &
              "		 INNER JOIN NotasFiscaisXItens " & vbCrLf &
              "			ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf &
              "		   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
              "		   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf &
              "		   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf &
              "		   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf &
              "		   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf &
              "		   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf &
              "     LEFT JOIN NotasFiscaisXRomaneios NFXR " & vbCrLf &
              "       On NFXR.Empresa_Id                 = NotasFiscais.Empresa_Id  " & vbCrLf &
              "      AND NFXR.EndEmpresa_Id              = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
              "      And NFXR.Cliente_Id                 = NotasFiscais.Cliente_Id  " & vbCrLf &
              "      AND NFXR.EndCliente_Id              = NotasFiscais.EndCliente_Id  " & vbCrLf &
              "      And NFXR.EntradaSaida_Id            = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
              "      AND NFXR.Serie_Id                   = NotasFiscais.Serie_Id  " & vbCrLf &
              "      And NFXR.Nota_Id                    = NotasFiscais.Nota_Id " & vbCrLf &
              "    LEFT JOIN Romaneios ROM" & vbCrLf &
              "       ON ROM.Empresa_Id                 = NFXR.Empresa_Id " & vbCrLf &
              "      AND ROM.EndEmpresa_Id              = NFXR.EndEmpresa_Id " & vbCrLf &
              "      And ROM.Romaneio_Id                = NFXR.Romaneio_Id" & vbCrLf &
              "    INNER JOIN Produtos PRO              " & vbCrLf &
              "       ON PRO.Produto_Id = NotasFiscaisXItens.Produto_Id " & vbCrLf &
              "	   INNER JOIN SubOperacoes SO" & vbCrLf &
              "			ON NotasFiscais.Operacao    = SO.Operacao_Id " & vbCrLf &
              "		   AND NotasFiscais.SubOperacao = SO.SubOperacoes_Id " & vbCrLf &
              "    WHERE NotasFiscais.situacao   In (1,4,7) " & vbCrLf &
              "        And NotasFiscais.TipoDeDocumento  = 1 " & vbCrLf &
              "        And NotasFiscais.Empresa_Id       = '" & pCodigoEmpresa & "'" & vbCrLf &
              "        And NotasFiscais.EndEmpresa_Id    = " & pEndEmpresa & vbCrLf &
              "        And NotasFiscaisXItens.Produto_Id = '" & _CodigoProduto & "'" & vbCrLf

        sql &= "        and (SO.EstoqueFisico = 'S' or SO.EstoqueFiscal = 'S') " & vbCrLf

        sql &= andtipoOperacao

        sql &= " 	 Group by NotasFiscais.Empresa_id, NotasFiscais.EndEmpresa_id, NotasFiscaisXItens.Produto_Id" & vbCrLf

        If pSO.EstoqueFisico Then
            sql &= " 	 union" & vbCrLf &
                   "    SELECT P.empresa_id, P.EndEmpresa_Id, P.Produto_Id, " & vbCrLf &
                   "	       sum(Entradas) - sum(Saidas) as QtdeFisica, 0" & vbCrLf &
                   "	  FROM Producao P" & vbCrLf &
                   "     INNER JOIN Suboperacoes so" & vbCrLf &
                   "	    on p.operacao_id    = so.Operacao_id" & vbCrLf &
                   "	   and p.suboperacao_id = so.Suboperacoes_id" & vbCrLf &
                   "     where P.FisicoFiscal_Id = 1 " & vbCrLf &
                   "       and P.Produto_Id = '" & _CodigoProduto & "'" & vbCrLf &
                   "       and P.Empresa_Id = '" & pCodigoEmpresa & "'" & vbCrLf &
                   "       and P.EndEmpresa_Id = " & pEndEmpresa & vbCrLf

            sql &= andtipoOperacao

            sql &= "     GROUP BY P.empresa_id, P.EndEmpresa_Id, P.Produto_Id" & vbCrLf
        End If

        If pSO.EstoqueFiscal Then
            sql &= " 	 union" & vbCrLf &
                   "    SELECT P.empresa_id, P.EndEmpresa_Id, P.Produto_Id, " & vbCrLf &
                   "	       0, sum(Entradas) - sum(Saidas) as QtdeFiscal" & vbCrLf &
                   "	  FROM Producao P" & vbCrLf &
                   "     INNER JOIN Suboperacoes so" & vbCrLf &
                   "	    on p.operacao_id    = so.Operacao_id" & vbCrLf &
                   "	   and p.suboperacao_id = so.Suboperacoes_id" & vbCrLf &
                   "     where P.FisicoFiscal_Id = 2 " & vbCrLf &
                   "       and P.Produto_Id = '" & _CodigoProduto & "'" & vbCrLf &
                   "       and P.Empresa_Id = '" & pCodigoEmpresa & "'" & vbCrLf &
                   "       and P.EndEmpresa_Id = " & pEndEmpresa & vbCrLf

            sql &= andtipoOperacao

            sql &= "	 GROUP BY P.empresa_id, P.EndEmpresa_Id, P.Produto_Id" & vbCrLf
        End If

        sql &= "	   ) Sb" & vbCrLf & _
               " Inner Join Produtos P" & vbCrLf & _
               "	   on P.Produto_Id     = Sb.Produto_Id" & vbCrLf & _
               "   and P. ControlarEstoque = 'S'--(P. ControlarEstoque = 'S' AND isnull(P.ControlarLote,'N') = 'N' AND isnull(P.ControlarEmbalagem,'N') = 'N')" & vbCrLf & _
               " Where P.Produto_Id     = '" & _CodigoProduto & "'" & vbCrLf

        If pCodigoEmpresa.Length > 0 Then
            sql &= "   and sb.Empresa_id    ='" & pCodigoEmpresa & "'" & vbCrLf & _
                   "   and Sb.EndEmpresa_id = " & pEndEmpresa & vbCrLf
        End If

        sql &= " Group by sb.Produto_Id, P.Nome, P.Unidade"

        Dim Banco As New AcessaBanco
        Dim Ds As DataSet

        Ds = Banco.ConsultaDataSet(sql, "Estoque")

        For Each row As DataRow In Ds.Tables(0).Rows
            Dim PE As New SaldoProdutoEstoqueLoteEmbalagem
            PE.CodigoProduto = row("CodigoProduto")
            PE.NomeProduto = row("NomeProduto")
            PE.SaldoFisico = row("SaldoFisico")
            PE.SaldoFiscal = row("SaldoFiscal")
            PE.Unidade = row("Unidade")
            Me.Add(PE)
        Next

    End Sub

    Public Sub CarregarPossiveisEntradasLoteClassificacao(Optional ByVal CodigoSafra As String = "")
        Dim sql As String
        sql = "SELECT L.Produto_id as CodigoProduto, " & vbCrLf & _
            "       p.Nome as NomeProduto, " & vbCrLf & _
            "       L.Lote_id as Lote," & vbCrLf & _
            "       L.DataValidade," & vbCrLf & _
            "       L.Tipo AS TipoLote," & vbCrLf & _
            "       LXC.Classificacao_Id as Classificacao," & vbCrLf & _
            "       isnull(LXC.PesoSaco,0) as PesoSaco," & vbCrLf & _
            "       isnull(Emb.EmbalagemIndea,'') as EmbalagemIndea, " & vbCrLf & _
            "       isnull(PXE.Embalagem_Id,0) as Embalagem," & vbCrLf & _
            "       isnull(Emb.Descricao,'') as DescEmbalagem," & vbCrLf & _
            "       isnull(PXE.TipoDeEmbalagem_Id,'') as TipoDeEmbalagem, " & vbCrLf & _
            "       isnull(te.Descricao,'') as DescTipoDeEmbalagem, " & vbCrLf & _
            "       isnull(PXE.CapacidadeEmbalagem_Id,0.00000) as CapacidadeEmbalagem, " & vbCrLf & _
            "       P.Unidade, " & vbCrLf & _
            "       isnull(PXE.PesoVariavel,0) as PesoVariavel" & vbCrLf & _
            "  FROM Produtos P " & vbCrLf & _
            " INNer Join Lote L " & vbCrLf & _
            "    on P.Produto_Id = L.Produto_Id" & vbCrLf & _
            " INNER JOIN LoteXClassificacao LXC" & vbCrLf & _
            "    ON L.Produto_id = LXC.Produto_id " & vbCrLf & _
            "   AND L.Lote_id    = LXC.Lote_Id" & vbCrLf & _
            "  LEFT Join ProdutoXEmbalagem PXE" & vbCrLf & _
            "    On PXE.Produto_Id = L.Produto_Id" & vbCrLf & _
            "  LEFT Join Embalagens Emb" & vbCrLf & _
            "    on Emb.Embalagem_Id = PXE.Embalagem_Id" & vbCrLf & _
            "   LEFT Join TipoDeEmbalagem TE" & vbCrLf & _
            "    on TE.TipoDeEmbalagem_Id = PXE.TipoDeEmbalagem_Id " & vbCrLf & _
            " Where isnull(P.ControlarLote,'N') = 'S'" & vbCrLf & _
            "   and L.Produto_id = '" & _CodigoProduto & "'" & vbCrLf & _
            "   and case" & vbCrLf & _
            "         when L.Tipo = 1" & vbCrLf & _
            "           then L.safra" & vbCrLf & _
            "           else '" & CodigoSafra & "'" & vbCrLf & _
            "       end = '" & CodigoSafra & "'" & vbCrLf & _
            "  UNION" & vbCrLf & _
            " SELECT PXE.Produto_id," & vbCrLf & _
            "        P.Nome as NomeProduto," & vbCrLf & _
            "        '' as Lote_id, " & vbCrLf & _
            "        '' as DataValidade, " & vbCrLf & _
            "        '' as TipoLote, " & vbCrLf & _
            "        '' as Classificacao_Id, " & vbCrLf & _
            "         0 as PesoSaco," & vbCrLf & _
            "        Emb.EmbalagemIndea, " & vbCrLf & _
            "        PXE.Embalagem_Id, " & vbCrLf & _
            "        Emb.Descricao as DescEmbalagem," & vbCrLf & _
            "        PXE.TipoDeEmbalagem_Id, " & vbCrLf & _
            "        isnull(te.Descricao,'') as DescTipoDeEmbalagem, " & vbCrLf & _
            "        PXE.CapacidadeEmbalagem_Id, " & vbCrLf & _
            "        P.Unidade, " & vbCrLf & _
            "        isnull(PXE.PesoVariavel,0) as PesoVariavel" & vbCrLf & _
            "  FROM Produtos P  " & vbCrLf & _
            " Inner Join ProdutoXEmbalagem PXE" & vbCrLf & _
            "    on P.Produto_Id = PXE.Produto_id" & vbCrLf & _
            " Inner Join Embalagens Emb" & vbCrLf & _
            "    on Emb.Embalagem_Id = PXE.Embalagem_Id" & vbCrLf & _
            " Inner Join TipoDeEmbalagem TE" & vbCrLf & _
            "    on TE.TipoDeEmbalagem_Id = PXE.TipoDeEmbalagem_Id " & vbCrLf & _
            " Where isnull(P.ControlarLote,'N') = 'N'" & vbCrLf & _
            "   and PXE.Produto_id = '" & _CodigoProduto & "'" & vbCrLf


        Dim Banco As New AcessaBanco
        Dim Ds As DataSet

        Ds = Banco.ConsultaDataSet(sql, "Estoque")

        For Each row As DataRow In Ds.Tables(0).Rows
            Dim PE As New SaldoProdutoEstoqueLoteEmbalagem
            PE.CodigoProduto = row("CodigoProduto")
            PE.NomeProduto = row("NomeProduto")
            PE.Lote = row("Lote")
            PE.DataDeValidade = row("DataValidade")
            PE.TipoDeLote = row("TipoLote")
            PE.Classificacao = row("Classificacao")
            PE.PesoSaco = row("PesoSaco")
            PE.Embalagem = row("Embalagem")
            PE.DescEmbalagem = row("DescEmbalagem")
            PE.EmbalagemIndea = row("EmbalagemIndea")
            PE.TipoDeEmbalagem = row("TipoDeEmbalagem")
            PE.DescTipoDeEmbalagem = row("DescTipoDeEmbalagem")
            PE.CapacidadeEmbalagem = row("CapacidadeEmbalagem")
            PE.Unidade = row("Unidade")
            PE.PesoVariavel = row("PesoVariavel")
            PE.SaldoFisico = 0
            PE.SaldoFiscal = 0
            Me.Add(PE)
        Next
    End Sub

    Public Sub CarregaSaldoDisponivelSaidaLoteClassificao(ByVal objSubOperacoes As [Lib].Negocio.SubOperacao, ByVal pCodigoEmpresa As String, ByVal pEndEmpresa As Integer, Optional ByVal Lote As String = "", Optional ByVal Classificacao As String = "", Optional ByVal CodigoEmbalagem As Integer = 0, Optional ByVal TipoEmbalagem As String = "", Optional ByVal CapacidadeEmbalagem As Decimal = 0)
        Me.Clear()
        Dim sql As String
        Dim andtipoOperacao As String = ""

        If objSubOperacoes.Consignacao Then
            andtipoOperacao = " and isnull(SO.consignacao,0)   = 1" & vbCrLf

            If (objSubOperacoes.EntradaSaida = eEntradaSaida.Entrada And objSubOperacoes.Devolucao = False) Or (objSubOperacoes.EntradaSaida = eEntradaSaida.Saida And objSubOperacoes.Devolucao = True) Then
                'Se for uma entrada que nao é uma devolucao ou se for uma saida que for devolucao entao estamos pegando e devolvendo produto Consignado
                andtipoOperacao = " and ((SO.EntradaSaida = 'E' and SO.devolucao = 'N') or (SO.EntradaSaida = 'S' and SO.devolucao = 'S')) " & vbCrLf
            Else
                'Agora se for uma entrada que é devolucao ou se for uma saida que nao é devolucao entao estamos mandando consignado e recebendo devolucao de produtos consignados
                andtipoOperacao = " and ((SO.EntradaSaida = 'E' and SO.devolucao = 'S') or (SO.EntradaSaida = 'S' and SO.devolucao = 'N')) " & vbCrLf
            End If

        ElseIf objSubOperacoes.AmostraGratis Then
            andtipoOperacao = " and isnull(SO.AmostraGratis,0) = 1" & vbCrLf
        Else
            andtipoOperacao = " and isnull(SO.consignacao,0)   = 0" & vbCrLf & _
                              " and isnull(SO.AmostraGratis,0) = 0" & vbCrLf
        End If

        sql = "Select sb.Produto_Id as CodigoProduto," & vbCrLf & _
              "       P.Nome as NomeProduto," & vbCrLf & _
              "       P.Unidade," & vbCrLf & _
              "       isnull(sb.Lote,'') as Lote," & vbCrLf & _
              "       isnull(sb.DataValidade,'') as DataValidade," & vbCrLf & _
              "       isnull(sb.TipoLote,'') as TipoLote," & vbCrLf & _
              "       isnull(sb.Classificacao,'') as Classificacao," & vbCrLf & _
              "       isnull(sb.PesoSaco,0) as PesoSaco," & vbCrLf & _
              "       isnull(sb.Embalagem,0) as Embalagem, " & vbCrLf & _
              "       isnull(Emb.Descricao,'') as DescEmbalagem," & vbCrLf & _
              "       isnull(Emb.EmbalagemIndea,'') as EmbalagemIndea, " & vbCrLf & _
              "       isnull(sb.TipoDeEmbalagem,'') as TipoDeEmbalagem," & vbCrLf & _
              "       isnull(TE.Descricao,'') as DescTipoDeEmbalagem," & vbCrLf & _
              "       isnull(sb.CapacidadeEmbalagem,0) as CapacidadeEmbalagem," & vbCrLf & _
              "       sum(isnull(sb.QtdeFisica,0)) SaldoFisico," & vbCrLf & _
              "       sum(isnull(sb.QtdeFiscal,0)) SaldoFiscal," & vbCrLf & _
              "       isnull(PXE.PesoVariavel,0) as PesoVariavel" & vbCrLf & _
              " From " & vbCrLf & _
              "	   (" & vbCrLf & _
              "		SELECT NotasFiscais.Empresa_id, NotasFiscais.EndEmpresa_id," & vbCrLf & _
              "            NotasFiscaisXItens.Produto_Id, " & vbCrLf & _
              "			   isnull(NotasFiscaisXItens.Lote,'') as Lote," & vbCrLf & _
              "			   isnull(L.DataValidade,'') as DataValidade," & vbCrLf & _
              "			   isnull(L.Tipo,'') as TipoLote," & vbCrLf & _
              "			   isnull(NotasFiscaisXItens.Classificacao,'') as Classificacao," & vbCrLf & _
              "            isnull(LC.PesoSaco,0) as PesoSaco, " & vbCrLf & _
              "            NotasFiscaisXItens.Embalagem," & vbCrLf & _
              "            NotasFiscaisXItens.TipoDeEmbalagem," & vbCrLf & _
              "            NotasFiscaisXItens.CapacidadeEmbalagem," & vbCrLf & _
              "			   sum(case" & vbCrLf & _
              "					 when so.EntradaSaida = 'E'" & vbCrLf & _
              "					  then NotasFiscaisXItens.QtdeDeEmbalagem " & vbCrLf & _
              "					  else  NotasFiscaisXItens.QtdeDeEmbalagem * - 1 " & vbCrLf & _
              "				   end) as QtdeDeEmbalagem," & vbCrLf & _
              "			   sum(case" & vbCrLf & _
              "					 when so.EntradaSaida = 'E'" & vbCrLf & _
              "					  then NotasFiscaisXItens.QuantidadeFisica " & vbCrLf & _
              "					  else  NotasFiscaisXItens.QuantidadeFisica * - 1" & vbCrLf & _
              "				   end) as QtdeFisica," & vbCrLf & _
              "			   sum(case" & vbCrLf & _
              "					 when so.EntradaSaida = 'E'" & vbCrLf & _
              "					  then NotasFiscaisXItens.QuantidadeFiscal " & vbCrLf & _
              "					  else  NotasFiscaisXItens.QuantidadeFiscal * - 1" & vbCrLf & _
              "				   end) as QtdeFiscal" & vbCrLf & _
              "		  FROM NotasFiscais" & vbCrLf & _
              "		 INNER JOIN NotasFiscaisXItens " & vbCrLf & _
              "			ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
              "		   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
              "		   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
              "		   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
              "		   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
              "		   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
              "		   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
              "		 INNER JOIN SubOperacoes so" & vbCrLf & _
              "			ON NotasFiscais.Operacao    = so.Operacao_Id " & vbCrLf & _
              "		   AND NotasFiscais.SubOperacao = so.SubOperacoes_Id " & vbCrLf & _
              "       Left Join LoteXClassificacao LC" & vbCrLf & _
              "         on NotasFiscaisXItens.Produto_id    = LC.Produto_id" & vbCrLf & _
              "        And NotasFiscaisXItens.Lote          = LC.Lote_Id" & vbCrLf & _
              "        And NotasFiscaisXItens.Classificacao = LC.Classificacao_Id" & vbCrLf & _
              "       LEFT JOIN Lote L " & vbCrLf & _
              "         ON NotasFiscaisXItens.Produto_id   = L.Produto_id" & vbCrLf & _
              "        AND NotasFiscaisXItens.Lote         = L.Lote_id " & vbCrLf & _
              "      Where so.EstoqueFisico   = 'S'" & vbCrLf & _
              "        and NotasFiscais.situacao       in (1,4,7)" & vbCrLf & _
              "        and NotasFiscais.TipoDeDocumento = 1 " & vbCrLf

        sql &= andtipoOperacao

        sql &= " 	 Group by NotasFiscais.Empresa_id, NotasFiscais.EndEmpresa_id," & vbCrLf & _
              "               NotasFiscaisXItens.Produto_Id," & vbCrLf & _
              "				  isnull(NotasFiscaisXItens.Lote,'')," & vbCrLf & _
              "				  isnull(L.DataValidade,'')," & vbCrLf & _
              "				  isnull(L.Tipo,'')," & vbCrLf & _
              "				  isnull(NotasFiscaisXItens.Classificacao,'')," & vbCrLf & _
              "               isnull(LC.PesoSaco,0)," & vbCrLf & _
              "               NotasFiscaisXItens.Embalagem," & vbCrLf & _
              "               NotasFiscaisXItens.TipoDeEmbalagem," & vbCrLf & _
              "               NotasFiscaisXItens.CapacidadeEmbalagem" & vbCrLf & _
              "		 union" & vbCrLf & _
              "		 SELECT P.empresa_id, P.EndEmpresa_Id, P.Produto_Id, P.Lote_Id, L.DataValidade, L.Tipo AS TipoLote, P.Classificacao_Id, isnull(LC.PesoSaco,0) as PesoSaco, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id," & vbCrLf & _
              "                (sum(Entradas) - sum(Saidas)) / case when CapacidadeEmbalagem_Id = 0 then sum(Entradas) - sum(Saidas) else CapacidadeEmbalagem_Id end as QtdedeEmbalagem," & vbCrLf & _
              "				sum(Entradas) - sum(Saidas) as SaldoFisico, 0" & vbCrLf & _
              "		   FROM Producao P" & vbCrLf & _
              "       INNER JOIN SubOperacoes so" & vbCrLf & _
              "          ON P.Operacao_Id    = so.Operacao_Id" & vbCrLf & _
              "         AND P.SubOperacao_id = so.SubOperacoes_Id" & vbCrLf & _
              "        Left Join LoteXClassificacao LC" & vbCrLf & _
              "          on P.Produto_id       = LC.Produto_id" & vbCrLf & _
              "         And P.Lote_id          = LC.Lote_Id" & vbCrLf & _
              "         And P.Classificacao_id = LC.Classificacao_Id" & vbCrLf & _
              "        LEFT JOIN Lote L " & vbCrLf & _
              "          ON P.Produto_Id = L.Produto_id" & vbCrLf & _
              "         AND P.Lote_Id = L.Lote_id " & vbCrLf & _
              "       where P.FisicoFiscal_Id  = 1 " & vbCrLf

        sql &= andtipoOperacao

        sql &= "		  GROUP BY P.empresa_id, P.EndEmpresa_Id, P.Produto_Id, P.Lote_Id, L.DataValidade, L.Tipo, P.Classificacao_Id, isnull(LC.PesoSaco,0), Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id" & vbCrLf & _
              "		 union" & vbCrLf & _
              "		 SELECT P.empresa_id, P.EndEmpresa_Id, P.Produto_Id, P.Lote_Id, L.DataValidade, L.Tipo AS TipoLote, P.Classificacao_Id, isnull(LC.PesoSaco,0) as PesoSaco, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id," & vbCrLf & _
              "                (sum(Entradas) - sum(Saidas)) / case when CapacidadeEmbalagem_Id = 0 then sum(Entradas) - sum(Saidas) else CapacidadeEmbalagem_Id end as QtdedeEmbalagem," & vbCrLf & _
              "				0, sum(Entradas) - sum(Saidas) as SaldoFiscal" & vbCrLf & _
              "		   FROM Producao P" & vbCrLf & _
              "       INNER JOIN SubOperacoes so" & vbCrLf & _
              "          ON P.Operacao_Id    = so.Operacao_Id" & vbCrLf & _
              "         AND P.SubOperacao_id = so.SubOperacoes_Id" & vbCrLf & _
              "        Left Join LoteXClassificacao LC" & vbCrLf & _
              "          on P.Produto_id       = LC.Produto_id" & vbCrLf & _
              "         And P.Lote_Id          = LC.Lote_Id" & vbCrLf & _
              "         And P.Classificacao_id = LC.Classificacao_Id" & vbCrLf & _
              "        LEFT JOIN Lote L " & vbCrLf & _
              "          ON P.Produto_Id = L.Produto_id" & vbCrLf & _
              "         AND P.Lote_Id = L.Lote_id " & vbCrLf & _
              "       where P.FisicoFiscal_Id = 2 " & vbCrLf

        sql &= andtipoOperacao

        sql &= "		  GROUP BY P.empresa_id, P.EndEmpresa_Id, P.Produto_Id, P.Lote_Id, L.DataValidade, L.Tipo, P.Classificacao_Id, isnull(LC.PesoSaco,0), Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id" & vbCrLf & _
               "	   ) Sb" & vbCrLf & _
               " Inner Join Produtos P" & vbCrLf & _
               "	   on P.Produto_Id     = Sb.Produto_Id" & vbCrLf & _
               "   and (P.ControlarLote = 'S' or P.ControlarEmbalagem = 'S')" & vbCrLf & _
               "  Left Join Embalagens Emb" & vbCrLf & _
               "    on Emb.Embalagem_id = Sb.Embalagem " & vbCrLf & _
               "  Left Join TipoDeEmbalagem TE" & vbCrLf & _
               "    on TE.TipoDeEmbalagem_id = Sb.TipoDeEmbalagem" & vbCrLf & _
               "  Left Join ProdutoXEmbalagem PxE" & vbCrLf & _
               "    on PxE.Produto_Id             = sb.Produto_Id" & vbCrLf & _
               "   and PXE.Embalagem_Id           = sb.Embalagem" & vbCrLf & _
               "   and PXE.TipoDeEmbalagem_Id     = sb.TipoDeEmbalagem" & vbCrLf & _
               "   and PXE.CapacidadeEmbalagem_Id = sb.CapacidadeEmbalagem" & vbCrLf & _
               " Where P.Produto_Id     = '" & Me.CodigoProduto & "'" & vbCrLf & _
               "   and Len(sb.Lote) + len(Emb.EmbalagemIndea) > 0 " & vbCrLf

        If pCodigoEmpresa.Length > 0 Then
            sql &= "   and sb.Empresa_id    ='" & pCodigoEmpresa & "'" & vbCrLf & _
                   "   and Sb.EndEmpresa_id = " & pEndEmpresa & vbCrLf
        End If


        If Lote.Length > 0 Then
            sql &= "   and isnull(sb.Lote,'')          = '" & Lote & "'" & vbCrLf & _
                   "   and isnull(sb.Classificacao,'') = '" & Classificacao & "'" & vbCrLf
        End If

        If CodigoEmbalagem > 0 Then
            sql &= "   and isnull(sb.Embalagem,0)           = " & CodigoEmbalagem & vbCrLf & _
                   "   and isnull(sb.TipoDeEmbalagem,'')    ='" & TipoEmbalagem & "'" & vbCrLf & _
                   "   and isnull(sb.CapacidadeEmbalagem,0) = " & Str(CapacidadeEmbalagem) & vbCrLf
        End If

        sql &= "  Group by sb.Produto_Id, P.Nome, P.Unidade, isnull(sb.Lote,''), isnull(sb.DataValidade,''), isnull(sb.TipoLote,''), isnull(sb.Classificacao,''), isnull(sb.PesoSaco,0) , isnull(Emb.EmbalagemIndea,''), isnull(sb.Embalagem,0), isnull(Emb.Descricao,''), isnull(sb.TipoDeEmbalagem,''), isnull(TE.Descricao,''), isnull(sb.CapacidadeEmbalagem,0), isnull(PXE.PesoVariavel,0)" & vbCrLf
        sql &= "  having sum(isnull(sb.QtdeFisica,0)) > 0 or sum(isnull(sb.QtdeFiscal,0)) > 0"

        Dim Banco As New AcessaBanco
        Dim Ds As DataSet

        Ds = Banco.ConsultaDataSet(sql, "Estoque")

        For Each row As DataRow In Ds.Tables(0).Rows
            Dim PE As New SaldoProdutoEstoqueLoteEmbalagem
            PE.CodigoProduto = row("CodigoProduto")
            PE.NomeProduto = row("NomeProduto")
            PE.Lote = row("Lote")
            PE.DataDeValidade = row("DataValidade")
            PE.TipoDeLote = row("TipoLote")
            PE.Classificacao = row("Classificacao")
            PE.PesoSaco = row("PesoSaco")
            PE.Embalagem = row("Embalagem")
            PE.DescEmbalagem = row("DescEmbalagem")
            PE.EmbalagemIndea = row("EmbalagemIndea")
            PE.TipoDeEmbalagem = row("TipoDeEmbalagem")
            PE.DescTipoDeEmbalagem = row("DescTipoDeEmbalagem")
            PE.CapacidadeEmbalagem = row("CapacidadeEmbalagem")
            PE.SaldoFisico = row("SaldoFisico")
            PE.SaldoFiscal = row("SaldoFiscal")
            PE.Unidade = row("Unidade")
            PE.PesoVariavel = row("PesoVariavel")
            Me.Add(PE)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class SaldoProdutoEstoqueLoteEmbalagem

#Region "Fields"
    Private _CodigoProduto As String
    Private _NomeProduto As String
    Private _Unidade As String
    Private _Lote As String
    Private _DataDeValidade As Date
    Private _TipoDeLote As Integer
    Private _Classificacao As String
    Private _PesoSaco As Decimal
    Private _Embalagem As Integer = 0
    Private _QtdeEmbalagem As Integer = 0
    Private _EmbalagemIndea As String
    Private _DescEmbalagem As String
    Private _TipoDeEmbalagem As String
    Private _DescTipoDeEmbalagem As String
    Private _CapacidadeEmbalagem As Double = 0
    Private _QtdeEmbXCapacidade As Double = 0
    Private _SaldoFisico As Decimal
    Private _SaldoFiscal As Decimal
    Private _PesoVariavel As Boolean
#End Region

#Region "Property"
    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
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

    Public Property DataDeValidade() As Date
        Get
            Return _DataDeValidade
        End Get
        Set(ByVal value As Date)
            _DataDeValidade = value
        End Set
    End Property

    Public Property TipoDeLote() As Integer
        Get
            Return _TipoDeLote
        End Get
        Set(ByVal value As Integer)
            _TipoDeLote = value
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

    Public Property PesoSaco() As Decimal
        Get
            Return _PesoSaco
        End Get
        Set(ByVal value As Decimal)
            _PesoSaco = value
        End Set
    End Property

    Public Property Embalagem() As Integer
        Get
            Return _Embalagem
        End Get
        Set(ByVal value As Integer)
            _Embalagem = value
        End Set
    End Property

    Public Property QtdeEmbalagem() As Integer
        Get
            Return _QtdeEmbalagem
        End Get
        Set(ByVal value As Integer)
            _QtdeEmbalagem = value
        End Set
    End Property

    Public Property DescEmbalagem() As String
        Get
            Return _DescEmbalagem
        End Get
        Set(ByVal value As String)
            _DescEmbalagem = value
        End Set
    End Property

    Public Property EmbalagemIndea() As String
        Get
            Return _EmbalagemIndea
        End Get
        Set(ByVal value As String)
            _EmbalagemIndea = value
        End Set
    End Property

    Public Property TipoDeEmbalagem() As String
        Get
            Return _TipoDeEmbalagem
        End Get
        Set(ByVal value As String)
            _TipoDeEmbalagem = value
        End Set
    End Property

    Public Property DescTipoDeEmbalagem() As String
        Get
            Return _DescTipoDeEmbalagem
        End Get
        Set(ByVal value As String)
            _DescTipoDeEmbalagem = value
        End Set
    End Property

    Public Property CapacidadeEmbalagem() As Double
        Get
            Return _CapacidadeEmbalagem
        End Get
        Set(ByVal value As Double)
            _CapacidadeEmbalagem = value
        End Set
    End Property

    Public Property QtdeEmbXCapacidade() As Double
        Get
            Return _QtdeEmbXCapacidade
        End Get
        Set(ByVal value As Double)
            _QtdeEmbXCapacidade = value
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

    Public Property SaldoFiscal() As Decimal
        Get
            Return _SaldoFiscal
        End Get
        Set(ByVal value As Decimal)
            _SaldoFiscal = value
        End Set
    End Property

    Public ReadOnly Property SaldoEmbalagem()
        Get
            If Me.SaldoFiscal = 0 Then Return 0
            If Me.CapacidadeEmbalagem = 0 Then Return 0
            If Me.PesoVariavel And Me.Unidade = "KGS" Then
                Return Me.SaldoFiscal / Me.PesoSaco
            Else
                Return Me.SaldoFiscal / Me.CapacidadeEmbalagem
            End If

        End Get
    End Property

    Public Property PesoVariavel() As Boolean
        Get
            Return _PesoVariavel
        End Get
        Set(ByVal value As Boolean)
            _PesoVariavel = value
        End Set
    End Property
#End Region

End Class