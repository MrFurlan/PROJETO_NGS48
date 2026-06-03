Imports System.Data
Imports System.Web
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()>
Public Class ListNotaFiscalXItemXEncargo
    Inherits List(Of NotaFiscalXItemXEncargo)
    Implements IBaseEntity

#Region "Fields"
    Private _Parent As Negocio.NotaFiscalXItem
    Dim Banco As New AcessaBanco
    Dim _SaldoInicial As Decimal
    Dim _EncProduto As NotaFiscalXItemXEncargo
    Dim _EncLiquido As NotaFiscalXItemXEncargo

#End Region

#Region "Contrutores"
    Public Sub New(ByVal Parent As Negocio.NotaFiscalXItem, ByVal lstEtapa As List(Of eEtapaEncago))
        Me.Parent = Parent

        'TESTANDO - FURLAN - 01/10/2024
        'If Parent.NotaFiscal.IUD = "U" AndAlso HttpContext.Current.Session("ssNomeUsuario") = "FURLAN" Then
        '    CarregarEncargosDaNota()
        '    Exit Sub
        'End If

        'CONHECIMENTO DE TERCEIROS PARA NOTAS GERAIS - FURLAN - 01-04-2013
        If (Not Parent.NotaFiscal.NossaEmissao _
            AndAlso (Parent.NotaFiscal.CodigoTipoDeDocumento = 2 Or Parent.NotaFiscal.CodigoTipoDeDocumento = 10 Or Parent.NotaFiscal.CodigoTipoDeDocumento = 14 Or Parent.NotaFiscal.CodigoTipoDeDocumento = 57) AndAlso Parent.NotaFiscal.TipoDeDocumentoFrete Is Nothing) _
        Or (Parent.NotaFiscal.CodigoTipoDeDocumento = 2 AndAlso Parent.NotaFiscal.TipoDeDocumentoFrete IsNot Nothing AndAlso Parent.NotaFiscal.CarregandoNota) _
        Or (Parent.NotaFiscal.CodigoTipoDeDocumento = 14 AndAlso Parent.NotaFiscal.TipoDeDocumentoFrete IsNot Nothing AndAlso Parent.NotaFiscal.CarregandoNota) Then
            CarregaListaConhecimento()
        Else
            If Not Parent.Operacao Is Nothing Then
                If Parent.Operacao.UFDepositoDestino Then
                    If Parent.NotaFiscal.Empresa.CodigoEstado = Parent.NotaFiscal.Destino.CodigoEstado AndAlso Parent.UsarRegiao = False Then
                        'PARA CLIENTE DO MESMO ESTADO DA EMPRESA PROCURA APENAS NO ESTADO
                        CarregaLista(lstEtapa, Parent.NotaFiscal.Destino.CodigoEstado)
                    Else
                        CarregaLista(lstEtapa, Parent.NotaFiscal.Destino.CodigoEstado, Parent.NotaFiscal.Destino.Estado.Regiao)
                    End If
                Else
                    If Parent.NotaFiscal.Empresa.CodigoEstado = Parent.NotaFiscal.Cliente.CodigoEstado AndAlso Parent.UsarRegiao = False Then
                        'PARA CLIENTE DO MESMO ESTADO DA EMPRESA PROCURA APENAS NO ESTADO
                        CarregaLista(lstEtapa, Parent.NotaFiscal.Cliente.CodigoEstado)
                    Else
                        CarregaLista(lstEtapa, Parent.NotaFiscal.Cliente.CodigoEstado, Parent.NotaFiscal.Cliente.Estado.Regiao)
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub New()
    End Sub

    Public Sub CarregarEncargosDaNota()
        Dim sql As String = String.Empty
        Dim objEncargoXTaxa As New EncargoXTaxa()

        sql = "Select OEE.Encargo_id, OE.STICMS, OE.STIPI, OE.STPISCOFINS, ISNULL(OE.UsarCalculadoraDeImposto, 0) AS UsarCalculadoraDeImposto,  ISNULL(OE.STIBSCBS, 0) AS STIBSCBS, ISNULL(OE.ClassificacaoIBSCBS, 0) AS ClassificacaoIBSCBS, ISNULL(OE.ReducaoIBS_Perc, 0) AS ReducaoIBS_Perc ISNULL(OE.ReducaoCBS_Perc, 0) AS ReducaoCBS_Perc, isnull(OE.ObsPISCOFINS,0) as ObsPISCOFINS,  OE.Operacao, OE.SubOperacao, OE.GrupoProduto, OE.Produto, OE.GrupoFiscal, OE.CodigoFiscal, " & vbCrLf &
               "       OE.EstadoOrigem, OE.EstadoDestino, OEE.AliquotaBase, OEE.AliquotaLimite, isnull(encXtx.Percentual,0) as PercEncxTaxa, isnull(encXtx.SimplesNacional,0) as SimplesNacional, OEE.Sinal, En.ValorouPeso, isnull(NxE.centrodecusto,0) as Custo, " & vbCrLf &
               "       Case" & vbCrLf &
               "         When OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
               "           Then 1" & vbCrLf &
               "         When OEE.Encargo_Id = 'LIQUIDO' or OEE.Encargo_Id = 'LIQUIDOAPAGAR'" & vbCrLf &
               "           Then 3" & vbCrLf &
               "         Else 2" & vbCrLf &
               "       end ordem," & vbCrLf &
               "	   isnull(case" & vbCrLf &
               "	          when OEE.Encargo_Id = 'ICMS'" & vbCrLf &
               "	            then (SELECT Descricao FROM ObservacoesTributarias WHERE Codigo_Id = OE.ObsICMS AND Encargo = OEE.Encargo_Id)" & vbCrLf &
               "	            else (SELECT Descricao FROM ObservacoesTributarias WHERE Codigo_Id = OEE.ObservacaoTributaria AND Encargo = OEE.Encargo_Id)" & vbCrLf &
               "	          end,'') AS ObservacaoFiscal," & vbCrLf &
               "       isnull(NxE.Base,0) as Base, isnull(NxE.Percentual,0) as Percentual, isnull(NxE.Valor,0) as Valor, isnull(NxE.PercentualExibicao,0) as AliquotaExibicao, " & vbCrLf &
               "       isnull(NxE.BaseNova,0) as BaseNova, isnull(NxE.ValorNovo,0) as ValorNovo" & vbCrLf &
               "  from NotasFiscaisXEncargos NxE" & vbCrLf &
               "    Inner Join OperacaoXEstado OE" & vbCrLf &
               "            on OE.Codigo_Id = " & Parent.CodigoOperacaoEstado & vbCrLf &
               "    Inner Join OperacaoXEstadoXEncargo OEE" & vbCrLf &
               "            on OEE.Codigo_Id   = OE.Codigo_Id" & vbCrLf &
               "            and OEE.Encargo_Id = NxE.Encargo_Id" & vbCrLf &
               "    Inner Join Encargos En" & vbCrLf &
               "            on En.Encargo_id =  OEE.Encargo_Id" & vbCrLf &
               "    Left join (SELECT ext.Encargo_Id as Encargo, ext.Percentual, ext.Estado_Id, ext.Produto_Id, isnull(ext.SimplesNacional,0) as SimplesNacional" & vbCrLf &
               "               FROM EncargosXTaxas ext" & vbCrLf &
               "               WHERE Data_Id = (SELECT MAX(Data_Id)" & vbCrLf &
               "               				     FROM EncargosXTaxas" & vbCrLf &
               "               				    WHERE EncargosXTaxas.Encargo_Id = ext.Encargo_Id" & vbCrLf &
               "               				      AND EncargosXTaxas.Estado_Id  = ext.Estado_Id" & vbCrLf &
               "               				      AND EncargosXTaxas.Produto_Id = '" & Parent.CodigoProduto.ToString & "'" & vbCrLf &
               "               				      AND Data_id <= '" & Parent.NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "')" & vbCrLf &
               "                              ) encXtx" & vbCrLf &
               "            on encXtx.Encargo     = OEE.Encargo_id" & vbCrLf &
               "            AND encXtx.Estado_Id  = OE.EstadoOrigem" & vbCrLf &
               "            AND encXtx.Produto_Id = '" & Parent.CodigoProduto.ToString & "'" & vbCrLf &
               " where NxE.Empresa_id      ='" & Parent.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
               "   and NxE.EndEmpresa_id   = " & Parent.NotaFiscal.EnderecoEmpresa & vbCrLf &
               "   and NxE.Cliente_Id      ='" & Parent.NotaFiscal.CodigoCliente & "'" & vbCrLf &
               "   and NxE.EndCliente_id   = " & Parent.NotaFiscal.EnderecoCliente & vbCrLf &
               "   and NxE.EntradaSaida_Id ='" & Parent.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
               "   and NxE.Serie_id        ='" & Parent.NotaFiscal.Serie & "'" & vbCrLf &
               "   and NxE.Nota_id         = " & Parent.NotaFiscal.Codigo & vbCrLf &
               "   and NxE.Produto_id      ='" & Parent.CodigoProduto & "'" & vbCrLf &
               "   and NxE.Sequencia_Id    = " & Parent.Sequencia & vbCrLf &
               " order by ordem, OEE.Encargo_id"

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Encargos")

        If ds.Tables(0).Rows.Count > 0 Then

            If Parent.CodigoOperacaoEstado = 0 Then
                Parent.CodigoOperacaoEstado = ds.Tables(0).Rows(0)("Codigo_Id")
            End If

            For Each row As DataRow In ds.Tables(0).Rows
                Dim objNFEncargo As New Negocio.NotaFiscalXItemXEncargo(Me.Parent)
                With objNFEncargo
                    If row("Encargo_id") = "PRODUTO" Then
                        Parent.CFOP = row("CodigoFiscal")
                    End If

                    .Existe = 1
                    .Codigo = row("Encargo_id")

                    .SituacaoTributaria = row("STICMS")
                    .SituacaoTributariaIPI = row("STIPI")
                    .SituacaoTributariaPISCOFINS = row("STPISCOFINS")
                    .SituacaoTributariaPISCOFINSOBS = row("ObsPISCOFINS")

                    .UsarCalculadoraDeImposto = row("UsarCalculadoraDeImposto")
                    .SituacaoTributariaIBSCBS = row("STIBSCBS")
                    .ClassificacaoIBSCBS = row("ClassificacaoIBSCBS")
                    .ReducaoIBS_Perc = row("ReducaoIBS_Perc")
                    .ReducaoCBS_Perc = row("ReducaoCBS_Perc")

                    .CodigoOperacao = row("Operacao")
                    .CodigoSubOperacao = row("SubOperacao")
                    .CodigoGrupoProduto = row("GrupoProduto")
                    .CodigoProduto = row("Produto")
                    .EstadoOrigem = row("EstadoOrigem")
                    .EstadoDestino = row("EstadoDestino")
                    .BasePercentual = row("AliquotaBase")
                    .PercentualLimite = row("AliquotaLimite")
                    .PercentualEncagosXTaxa = row("PercEncxTaxa")
                    .PercentualSimplesNacional = row("SimplesNacional")
                    .Sinal = row("Sinal")

                    .ObservacaoFiscal = row("ObservacaoFiscal")
                    .ValorPeso = IIf(row("ValorouPeso"), 1, 0)

                    If Not Parent.CentroDeCustoInformado Is Nothing AndAlso Parent.CentroDeCustoInformado.Length > 0 Then
                        .CentroDeCusto = Parent.CentroDeCustoInformado
                    Else
                        .CentroDeCusto = row("Custo")
                    End If

                    .BaseNova = row("BaseNova")
                    .ValorNovo = row("ValorNovo")

                    .Base = row("Base")
                    .Percentual = row("Percentual")
                    .PercentualExibicao = row("AliquotaExibicao")
                    .Valor = row("Valor")

                    If .Codigo = "ICMS-ST" Then
                        objEncargoXTaxa.SelecionarVigente(Parent.NotaFiscal.Cliente.CodigoEstado, "ICMS-ST", Parent.NotaFiscal.Movimento, "")
                        .PercentualSimplesNacional = objEncargoXTaxa.SimplesNacional
                    End If

                    'Carrega o percentual de exibição
                    If .PercentualExibicao = 0 Then .PercentualExibicao = .Percentual

                    Dim Tp As eTipoPessoa
                    If Parent.NotaFiscal.Cliente.CodigoEstado = "EX" Then
                        Parent.Retencao = False
                        Tp = eTipoPessoa.Ambas
                    ElseIf Parent.NotaFiscal.CodigoCliente.Length = 14 OrElse Parent.NotaFiscal.CodigoTipoDeDocumento = 57 Then
                        Tp = eTipoPessoa.Juridica
                    Else
                        Tp = eTipoPessoa.Fisica
                    End If

                    If Not objNFEncargo.Encargo.PodeSofreRetencao Then
                        Me.Add(objNFEncargo)
                    ElseIf (objNFEncargo.Encargo.TipoPessoaRetencao = Tp Or objNFEncargo.Encargo.TipoPessoaRetencao = 3) Then
                        If Parent.Retencao Then
                            Me.Add(objNFEncargo)
                        End If
                    Else
                        Me.Add(objNFEncargo)
                    End If

                    If objNFEncargo.Codigo = "PRODUTO" Then Me.EncProduto = objNFEncargo
                    If objNFEncargo.Codigo = "LIQUIDO" Then Me.EncLiquido = objNFEncargo

                End With
            Next
        Else
            HttpContext.Current.Session("ssMessage") = Parent.Produto.ToString & " - " & Parent.Produto.Nome & ", Erro na Consulta dos Encargos da Nota Fiscal"
        End If
    End Sub


    Public Sub CarregaLista(ByVal lstEtapa As List(Of eEtapaEncago), Optional ByVal Destino As String = "", Optional Regiao As String = "")
        Dim sql As String = ""
        Dim uf As String = Destino
        Dim Rg As String = Regiao
        Dim AtualLiquido As Boolean = False

        Dim objEncargoXTaxa As New EncargoXTaxa()

        sql = "Select OE.Codigo_Id," & vbCrLf &
               "       case" & vbCrLf &
               "         when OE.Produto <> '' and OE.EstadoDestino = '" & uf & "'" & vbCrLf &
               "           then 1" & vbCrLf &
               "	     when OE.Produto =  '' and OE.EstadoDestino = '" & uf & "'" & vbCrLf &
               "	       then 2" & vbCrLf &
               "         when OE.Produto <> '' and OE.EstadoDestino = '" & Rg & "'" & vbCrLf &
               "           then 3" & vbCrLf &
               "	     when OE.Produto =  '' and OE.EstadoDestino = '" & Rg & "'" & vbCrLf &
               "	       then 4" & vbCrLf &
               "	     else " & vbCrLf &
               "	       0 " & vbCrLf &
               "       end Prioridade," & vbCrLf &
               "       Case" & vbCrLf &
               "         When OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
               "           Then 1" & vbCrLf &
               "         When OEE.Encargo_Id = 'LIQUIDO' or OEE.Encargo_Id = 'LIQUIDOAPAGAR'" & vbCrLf &
               "           Then 3" & vbCrLf &
               "         Else 2" & vbCrLf &
               "       end ordem," & vbCrLf &
               "       OE.GrupoProduto," & vbCrLf &
               "	   OE.Produto," & vbCrLf &
               "	   OE.Operacao," & vbCrLf &
               "	   OE.SubOperacao," & vbCrLf &
               "	   OE.EstadoOrigem," & vbCrLf &
               "	   OE.EstadoDestino," & vbCrLf &
               "	   OEE.Encargo_id," & vbCrLf &
               "	   En.ValorOuPeso," & vbCrLf &
               "       OE.STICMS, isnull(STIPI,0) as STIPI," & vbCrLf &
               "	   isnull(STPISCOFINS,0) as STPISCOFINS," & vbCrLf &
               "       ISNULL(OE.UsarCalculadoraDeImposto,0) AS UsarCalculadoraDeImposto," & vbCrLf &
               "       ISNULL(OE.STIBSCBS,0) AS STIBSCBS," & vbCrLf &
               "       ISNULL(OE.ClassificacaoIBSCBS,0) AS ClassificacaoIBSCBS," & vbCrLf &
               "       ISNULL(OE.ReducaoIBS_Perc,0) AS ReducaoIBS_Perc," & vbCrLf &
               "       ISNULL(OE.ReducaoCBS_Perc,0) AS ReducaoCBS_Perc," & vbCrLf &
               "	   isnull(OE.ObsPISCOFINS,0) as ObsPISCOFINS," & vbCrLf &
               "       OE.GrupoFiscal," & vbCrLf &
               "	   OE.CodigoFiscal," & vbCrLf &
               "       ISNULL(OEE.DebitaConta,'') DebitaConta," & vbCrLf &
               "	   ISNULL(OEE.CreditaConta, '') CreditaConta," & vbCrLf &
               "	   OEE.Sinal," & vbCrLf &
               "	   OEE.AliquotaBase," & vbCrLf &
               "       isnull(encXtx.Percentual, OEE.Aliquota) AS Aliquota," & vbCrLf &
               "       encXtx.Percentual," & vbCrLf &
               "	   encXtx.SimplesNacional," & vbCrLf &
               "	   OEE.AliquotaExibicao," & vbCrLf &
               "       isnull(OEE.AliquotaLimite,100) as AliquotaLimite," & vbCrLf &
               "	   case" & vbCrLf &
               "	        when OEE.Encargo_Id = 'ICMS'" & vbCrLf &
               "	            then (SELECT Descricao FROM ObservacoesTributarias WHERE Codigo_Id = OE.ObsICMS AND Encargo = OEE.Encargo_Id)" & vbCrLf &
               "	            else (SELECT Descricao FROM ObservacoesTributarias WHERE Codigo_Id = OEE.ObservacaoTributaria AND Encargo = OEE.Encargo_Id)" & vbCrLf &
               "	        end AS ObservacaoFiscal," & vbCrLf &
               "	   case" & vbCrLf &
               "	        when OEE.Encargo_Id = 'ICMS DESONERADO'" & vbCrLf &
               "	            then isnull(PDPD.Preco,0)" & vbCrLf &
               "	            else isnull(PDP.Preco,0)" & vbCrLf &
               "	        end AS PrecoDePauta" & vbCrLf &
               "  into #Temp" & vbCrLf &
               "  from " & IIf(Parent.CodigoOperacaoEstado > 0, "OperacaoXEstado", "VW_OperacoesVigentes") & " OE" & vbCrLf &
               " Inner Join OperacaoXEstadoXEncargo OEE" & vbCrLf &
               "    on OE.Codigo_Id = OEE.Codigo_Id" & vbCrLf &
               "  Left join (SELECT ext.Encargo_Id as Encargo, ext.Percentual, ext.Estado_Id, ext.Produto_Id, isnull(ext.SimplesNacional,0) as SimplesNacional" & vbCrLf &
               "               FROM EncargosXTaxas ext" & vbCrLf &
               "              WHERE Data_Id = (SELECT MAX(Data_Id)" & vbCrLf &
               "               				     FROM EncargosXTaxas" & vbCrLf &
               "               				    WHERE EncargosXTaxas.Encargo_Id = ext.Encargo_Id" & vbCrLf &
               "               				      AND EncargosXTaxas.Estado_Id  = ext.Estado_Id" & vbCrLf &
               "               				      AND EncargosXTaxas.Produto_Id = '" & Parent.CodigoProduto.ToString & "'" & vbCrLf

        If Parent.NotaFiscal.SubOperacao.Devolucao Then
            sql &= "               				      AND Data_id <= '" & Parent.NotaFiscal.DataDevolucao.ToString("yyyy-MM-dd") & "')" & vbCrLf
        Else
            sql &= "               				      AND Data_id <= '" & Parent.NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "')" & vbCrLf
        End If

        sql &= "                              ) encXtx" & vbCrLf &
               "    on encXtx.Encargo     = OEE.Encargo_id" & vbCrLf &
               "   AND encXtx.Estado_Id   = OE.EstadoOrigem" & vbCrLf &
               "   AND encXtx.Produto_Id  = '" & Parent.CodigoProduto.ToString & "'" & vbCrLf &
               "  Left join (SELECT 'ICMS' as Encargo, PP.Produto_id, PP.Estado_Id, PP.Preco" & vbCrLf &
               "               FROM PrecoDePauta PP" & vbCrLf &
               "              WHERE Data_Id = (SELECT MAX(Data_Id)" & vbCrLf &
               "               				     FROM PrecoDePauta" & vbCrLf &
               "               				    WHERE PrecoDePauta.Estado_Id  = PP.Estado_Id" & vbCrLf &
               "                                  AND PrecoDePauta.Produto_Id = PP.Produto_Id" & vbCrLf &
               "               				      AND PrecoDePauta.Data_id   <= '" & Parent.NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "')" & vbCrLf &
               "                               ) PDP" & vbCrLf &
               "    ON PDP.Encargo    = OEE.Encargo_id" & vbCrLf &
               "   AND PDP.Estado_Id  = OE.EstadoOrigem" & vbCrLf &
               "   AND PDP.Produto_Id ='" & Parent.CodigoProduto.ToString & "'" & vbCrLf &
               "  Left join (SELECT 'ICMS DESONERADO' as Encargo, PPD.Produto_id, PPD.Estado_Id, PPD.Preco" & vbCrLf &
               "               FROM PrecoDePauta PPD" & vbCrLf &
               "              WHERE Data_Id = (SELECT MAX(Data_Id)" & vbCrLf &
               "               				     FROM PrecoDePauta" & vbCrLf &
               "               				    WHERE PrecoDePauta.Estado_Id  = PPD.Estado_Id" & vbCrLf &
               "                                  AND PrecoDePauta.Produto_Id = PPD.Produto_Id" & vbCrLf &
               "               				      AND PrecoDePauta.Data_id   <= '" & Parent.NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "')" & vbCrLf &
               "                               ) PDPD" & vbCrLf &
               "    ON PDPD.Encargo    = OEE.Encargo_id" & vbCrLf &
               "   AND PDPD.Estado_Id  = OE.EstadoOrigem" & vbCrLf &
               "   AND PDPD.Produto_Id ='" & Parent.CodigoProduto.ToString & "'" & vbCrLf &
               "  Inner Join Encargos En" & vbCrLf &
               "     on En.Encargo_id =  OEE.Encargo_Id" & vbCrLf &
               "   LEFT JOIN Produtos Prd" & vbCrLf &
               "     on OEE.Encargo_Id = 'IPI'" & vbCrLf &
               "    and Prd.Produto_Id = '" & Parent.CodigoProduto.ToString & "'" & vbCrLf &
               "  Inner Join Operacoes OP" & vbCrLf &
               "     on OP.Operacao_id = OE.Operacao" & vbCrLf

        If Parent.CodigoOperacaoEstado > 0 Then
            sql &= "  where OE.Codigo_id =" & Parent.CodigoOperacaoEstado
        ElseIf Parent.CFOP > 0 Then
            sql &= "  where (OE.Empresa = '99999999' OR OE.Empresa = '" & Left(Parent.NotaFiscal.CodigoEmpresa, 8) & "')" & vbCrLf &
                   "    and OE.CodigoFiscal  = " & Parent.CFOP & vbCrLf &
                   "    and OE.GrupoProduto  ='" & Parent.Produto.CodigoGrupo & "'" & vbCrLf &
                   "    and (OE.Produto      ='' or OE.Produto = '" & Parent.Produto.Codigo & "')" & vbCrLf &
                   "    and OE.EstadoOrigem  ='" & Parent.NotaFiscal.Empresa.CodigoEstado & "'" & vbCrLf &
                   "    and OE.EstadoDestino in ('" & uf & "','" & Rg & "')" & vbCrLf
        ElseIf Parent.NotaFiscal.NotaDeTerceiro Then
            sql &= "  where 1=1 " & vbCrLf &
                    "    and OE.Operacao      = " & Parent.CodigoOperacao & vbCrLf &
                    "    and OE.SubOperacao   = " & Parent.CodigoSubOperacao & vbCrLf &
                    "    and OE.GrupoProduto  ='" & Parent.Produto.CodigoGrupo & "'" & vbCrLf &
                    "    and (OE.Produto      ='' or OE.Produto = '" & Parent.Produto.Codigo & "')" & vbCrLf &
                    "    and OE.EstadoOrigem  ='" & Parent.NotaFiscal.Empresa.CodigoEstado & "'" & vbCrLf &
                    "    and OE.EstadoDestino in ('" & uf & "','" & Rg & "')" & vbCrLf
        Else
            sql &= "  where (OE.Empresa = '99999999' OR OE.Empresa = '" & Left(Parent.NotaFiscal.CodigoEmpresa, 8) & "')" & vbCrLf &
                   "    and OE.Operacao      = " & Parent.CodigoOperacao & vbCrLf &
                   "    and OE.SubOperacao   = " & Parent.CodigoSubOperacao & vbCrLf &
                   "    and OE.GrupoProduto  ='" & Parent.Produto.CodigoGrupo & "'" & vbCrLf &
                   "    and (OE.Produto      ='' or OE.Produto = '" & Parent.Produto.Codigo & "')" & vbCrLf &
                   "    and OE.EstadoOrigem  ='" & Parent.NotaFiscal.Empresa.CodigoEstado & "'" & vbCrLf &
                   "    and OE.EstadoDestino in ('" & uf & "','" & Rg & "')" & vbCrLf
        End If

        Dim InscricaoCliente As String = Parent.NotaFiscal.Cliente.InscricaoEstadual.ToUpper()

        If InscricaoCliente.Contains("ISENTO") Then
            If Left(Parent.NotaFiscal.CodigoEmpresa, 8) = "24450490" OrElse Left(Parent.NotaFiscal.CodigoEmpresa, 8) = "40938762" OrElse Left(Parent.NotaFiscal.CodigoEmpresa, 8) = "49673784" Then
                'NÃO FAZ NADA PARA BAXI - FURLAN - 26/06/2024
                'NÃO FAZ NADA PARA RT - FURLAN - 19/08/2024
            ElseIf Not Parent.SubOperacao.Devolucao AndAlso Parent.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso Parent.NotaFiscal.Empresa.CodigoEstado = "PR" AndAlso Parent.NotaFiscal.Cliente.CodigoEstado = "PR" Then
                'NÃO FAZ VERIFICAÇÃO PARA O ESTADO DO PARANÁ POIS A EMPRESA TEM OBRIGAÇÃO DE PAGAR O ICMS INDEPENDENTE SE O CLIENTE TEM OU NÃO INSCRIÇÃO - FURLAN - 19/12/2023
            Else
                sql &= "    and OE.STICMS NOT IN(00,100)" & vbCrLf
            End If
        End If

        If Parent.NotaFiscal.Cliente.CodigoEstado <> "EX" AndAlso ((lstEtapa IsNot Nothing AndAlso lstEtapa.Count > 0 AndAlso lstEtapa.Where(Function(x) x = eEtapaEncago.Liquidacao).Count = 0) OrElse lstEtapa Is Nothing) Then
            If Parent.NotaFiscal.CodigoCliente.Length = 14 OrElse Parent.NotaFiscal.CodigoTipoDeDocumento = 57 Then
                If Not Parent.NotaFiscal.Cliente Is Nothing AndAlso Parent.NotaFiscal.Cliente.Codigo > 0 AndAlso Parent.NotaFiscal.Cliente.CodigoCategoria = 3 Then
                    ''SE CATEGORIA DE CLIENTE FOR PRODUTOR EMPRESA = 3 DEVE VIR TODOS OS ENCARGOS - FURLAN - 22/04/2024
                    sql &= " and (isnull(En.TipoDePessoa,3) in (1,2,3) " & IIf(Parent.NotaFiscal.Empresa.Empresa.ObrigaEncargo, "or (En.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & " and OEE.Encargo_Id not in('FUNRURAL','SENAR'))" & vbCrLf
                ElseIf Parent.NotaFiscal.Empresa.CodigoEstado = "MT" AndAlso Not Parent.NotaFiscal.SubOperacao Is Nothing AndAlso Parent.NotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso Not Parent.NotaFiscal.SubOperacao.Devolucao Then
                    ''SE EMPRESA MT E FOR SAIDA DEVE VIR ENCARGO FETHAB - FURLAN - 02/05/2024
                    sql &= " and (isnull(En.TipoDePessoa,3) in (1,2,3) " & IIf(Parent.NotaFiscal.Empresa.Empresa.ObrigaEncargo, "or (En.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & " and OEE.Encargo_Id not in('FUNRURAL','SENAR'))" & vbCrLf
                ElseIf Parent.NotaFiscal.Empresa.CodigoEstado = "MT" AndAlso Not Parent.NotaFiscal.SubOperacao Is Nothing AndAlso Parent.NotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso Parent.NotaFiscal.SubOperacao.Devolucao Then
                    ''DEVE VIR NA DEVOLUÇÃO TAMBÉM - FURLAN - 01/11/2024
                    sql &= " and (isnull(En.TipoDePessoa,3) in (1,2,3) " & IIf(Parent.NotaFiscal.Empresa.Empresa.ObrigaEncargo, "or (En.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & " and OEE.Encargo_Id not in('FUNRURAL','SENAR'))" & vbCrLf
                Else
                    sql &= " and (isnull(En.TipoDePessoa,3) in (2,3) " & IIf(Parent.NotaFiscal.Empresa.Empresa.ObrigaEncargo, "or (En.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & ")" & vbCrLf
                End If
            Else
                If Not Parent.SubOperacao.Devolucao AndAlso Parent.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                    sql &= " and (isnull(En.TipoDePessoa,3) in (1,3) " & IIf(Parent.NotaFiscal.Empresa.Empresa.ObrigaEncargo, "or (En.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & " or En.Encargo_id in('PIS','COFINS') )" & vbCrLf
                Else
                    sql &= " and (isnull(En.TipoDePessoa,3) in (1,3) " & IIf(Parent.NotaFiscal.Empresa.Empresa.ObrigaEncargo, "or (En.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & ")" & vbCrLf
                End If
            End If
        End If


        If lstEtapa IsNot Nothing AndAlso lstEtapa.Count > 0 Then
            Dim index = 0
            Dim strEtapa = String.Empty
            For Each e As eEtapaEncago In lstEtapa
                strEtapa &= IIf(Not String.IsNullOrEmpty(strEtapa), ",", "") & CInt(e)
            Next
            sql &= "    And En.Etapa             in (" & strEtapa & ");" & vbCrLf
        End If


        sql &= "Select Codigo_id," & vbCrLf &
               "       case when NXE.Empresa_id is null then 0 else 1 end Existe, T.ordem," & vbCrLf &
               "       T.GrupoProduto, T.Produto, T.Operacao, T.SubOperacao, T.EstadoOrigem, T.EstadoDestino," & vbCrLf &
               "       T.Encargo_Id," & vbCrLf &
               "       T.STICMS," & vbCrLf &
               "       T.STIPI," & vbCrLf &
               "       T.STPISCOFINS," & vbCrLf &
               "       T.UsarCalculadoraDeImposto," & vbCrLf &
               "       T.STIBSCBS," & vbCrLf &
               "       T.ClassificacaoIBSCBS," & vbCrLf &
               "       T.ReducaoIBS_Perc," & vbCrLf &
               "       T.ReducaoCBS_Perc," & vbCrLf &
               "       ISNULL(T.ObsPISCOFINS,0) ObsPISCOFINS," & vbCrLf &
               "       T.ValorOuPeso, T.GrupoFiscal, T.CodigoFiscal," & vbCrLf &
               "       T.DebitaConta, T.CreditaConta,  T.Sinal, T.AliquotaBase, T.Aliquota, isnull(T.AliquotaExibicao,0) as AliquotaExibicao,  T.AliquotaLimite," & vbCrLf &
               "       isnull(NxE.Base,0) as Base, isnull(NxE.Percentual,0) as Percentual,  isnull(NxE.Valor,0) as Valor," & vbCrLf &
               "       isnull(T.ObservacaoFiscal,'') as ObservacaoFiscal," & vbCrLf &
               "       isnull(T.Percentual,0) as PercEncxTaxa, isnull(T.SimplesNacional,0) AS SimplesNacional, isnull(NxE.centrodecusto,0) as custo, T.PrecoDePauta," & vbCrLf &
               "       isnull(NxE.BaseNova,0) as BaseNova, isnull(NxE.ValorNovo,0) as ValorNovo" & vbCrLf &
               "  Into #Pre " & vbCrLf &
               "  from #Temp T " & vbCrLf &
               "  left Join NotasFiscaisXEncargos NxE" & vbCrLf &
               "    on NxE.Empresa_id      ='" & Parent.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
               "   and NxE.EndEmpresa_id   = " & Parent.NotaFiscal.EnderecoEmpresa & vbCrLf &
               "   and NxE.Cliente_Id      ='" & Parent.NotaFiscal.CodigoCliente & "'" & vbCrLf &
               "   and NxE.EndCliente_id   = " & Parent.NotaFiscal.EnderecoCliente & vbCrLf &
               "   and NxE.EntradaSaida_Id ='" & Parent.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
               "   and NxE.Serie_id        ='" & Parent.NotaFiscal.Serie & "'" & vbCrLf &
               "   and NxE.Nota_id         = " & Parent.NotaFiscal.Codigo & vbCrLf &
               "   and NxE.Produto_id      ='" & Parent.CodigoProduto & "'" & vbCrLf &
               "   and NxE.Sequencia_Id    = " & Parent.Sequencia & vbCrLf &
               "   and NxE.Encargo_Id      = T.Encargo_Id" & vbCrLf &
               "   and NxE.Nota_id        > 0 "

        If Parent.IUD = "U" Then
            sql &= " And 1 = 2"
        End If

        sql &= "  where T.prioridade  = (Select min(prioridade) from #Temp) "

        If Parent.CodigoOperacaoEstado = 0 And Parent.CodigoOperacao > 0 And Parent.CodigoSubOperacao > 0 Then
            sql &= " AND T.Operacao = " & Parent.CodigoOperacao & " and T.SubOperacao = " & Parent.CodigoSubOperacao & ""
        End If

        sql &= "  order by ordem, encargo_id;" & vbCrLf &
               " Select * from #pre " & vbCrLf

        If Parent.IUD <> "U" Then sql &= " where Existe = (Select max(Existe) from #pre)" & vbCrLf

        sql &= " order by ordem, encargo_id "

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Encargos")

        If ds.Tables(0).Rows.Count > 0 Then

            If Parent.CodigoOperacaoEstado = 0 Then
                Parent.CodigoOperacaoEstado = ds.Tables(0).Rows(0)("Codigo_Id")
            End If

            Dim EncICMS As NotaFiscalXItemXEncargo = Nothing
            Dim EncADUANEIRAS As NotaFiscalXItemXEncargo = Nothing
            Dim EncFrete As NotaFiscalXItemXEncargo = Nothing
            Dim EncICMSDesonerado As NotaFiscalXItemXEncargo = Nothing
            Dim EncIPI As NotaFiscalXItemXEncargo = Nothing

            'PEGANDO IPI ANTES DE PERCORRER TODOS ENCARGOS PARA SOMAR NO ICMS-ST - FURLAN - 16/08/2024
            For Each row As DataRow In ds.Tables(0).Rows
                If row("Encargo_id") = "IPI" Then
                    Dim objNFEncargo As New Negocio.NotaFiscalXItemXEncargo(Me.Parent)

                    objNFEncargo.Codigo = row("Encargo_id")
                    objNFEncargo.BasePercentual = row("AliquotaBase")
                    objNFEncargo.Base = Math.Round((Parent.ValorTotal * objNFEncargo.BasePercentual) / 100, 2, MidpointRounding.AwayFromZero)
                    objNFEncargo.Percentual = row("Aliquota")
                    objNFEncargo.Valor = Math.Round((objNFEncargo.Base * objNFEncargo.Percentual) / 100, 2, MidpointRounding.AwayFromZero)

                    EncIPI = objNFEncargo
                End If
            Next

            For Each row As DataRow In ds.Tables(0).Rows
                Dim objNFEncargo As New Negocio.NotaFiscalXItemXEncargo(Me.Parent)
                With objNFEncargo
                    If row("Encargo_id") = "PRODUTO" Then
                        Parent.CFOP = row("CodigoFiscal")
                        If row("Existe") = 1 Then
                            AtualLiquido = False
                        Else
                            AtualLiquido = True
                        End If
                    End If

                    .Existe = row("Existe")
                    .Codigo = row("Encargo_id")

                    .SituacaoTributaria = row("STICMS")
                    .SituacaoTributariaIPI = row("STIPI")
                    .SituacaoTributariaPISCOFINS = row("STPISCOFINS")
                    .SituacaoTributariaPISCOFINSOBS = row("ObsPISCOFINS")

                    .UsarCalculadoraDeImposto = row("UsarCalculadoraDeImposto")
                    .SituacaoTributariaIBSCBS = row("STIBSCBS")
                    .ClassificacaoIBSCBS = row("ClassificacaoIBSCBS")
                    .ReducaoIBS_Perc = row("ReducaoIBS_Perc")
                    .ReducaoCBS_Perc = row("ReducaoCBS_Perc")

                    .CodigoOperacao = row("Operacao")
                    .CodigoSubOperacao = row("SubOperacao")
                    .CodigoGrupoProduto = row("GrupoProduto")
                    .CodigoProduto = row("Produto")
                    .EstadoOrigem = row("EstadoOrigem")
                    .EstadoDestino = row("EstadoDestino")
                    .BasePercentual = row("AliquotaBase")
                    .PercentualLimite = row("AliquotaLimite")
                    .PercentualEncagosXTaxa = row("PercEncxTaxa")
                    .PercentualSimplesNacional = row("SimplesNacional")
                    .Sinal = row("Sinal")

                    .ObservacaoFiscal = row("ObservacaoFiscal")
                    .ValorPeso = IIf(row("ValorouPeso"), 1, 0)

                    If Not Parent.CentroDeCustoInformado Is Nothing AndAlso Parent.CentroDeCustoInformado.Length > 0 Then
                        .CentroDeCusto = Parent.CentroDeCustoInformado
                    Else
                        .CentroDeCusto = row("Custo")
                    End If

                    .BaseNova = row("BaseNova")
                    .ValorNovo = row("ValorNovo")

                    If row("Valor") > 0 Then
                        .Base = row("Base")
                        .Percentual = row("Percentual")
                        .PercentualExibicao = row("AliquotaExibicao")
                        .Valor = row("Valor")
                    Else
                        If .ValorPeso = eValorPeso.Valor Then

                            If (Left(Parent.NotaFiscal.CodigoEmpresa, 8) = "40938762" OrElse Left(Parent.NotaFiscal.CodigoEmpresa, 8) = "49673784") AndAlso .Codigo = "FUNDO FECP" Then
                                objEncargoXTaxa.SelecionarVigente(Parent.NotaFiscal.Cliente.CodigoEstado, "ICMS-ST", Parent.NotaFiscal.Movimento, "")

                                .PercentualSimplesNacional = objEncargoXTaxa.SimplesNacional

                                Dim vst As Decimal = Math.Round((Me.EncProduto.Valor * .PercentualSimplesNacional) / 100, 2, MidpointRounding.AwayFromZero)

                                .Base = Me.EncProduto.Valor + vst
                                .Percentual = row("Aliquota")
                                .Valor = Math.Round((.Base * .Percentual) / 100, 2, MidpointRounding.AwayFromZero)

                            ElseIf (Left(Parent.NotaFiscal.CodigoEmpresa, 8) = "40938762" OrElse Left(Parent.NotaFiscal.CodigoEmpresa, 8) = "49673784") AndAlso .Codigo = "ICMS-ST" Then

                                Dim vst As Decimal = 0
                                Dim vlrST As Decimal = 0

                                If Parent.NotaFiscal.Cliente.CodigoEstado = "MG" Then

                                    If Parent.QuantidadeFiscal > 0 Then
                                        objEncargoXTaxa.SelecionarVigente(Parent.NotaFiscal.Cliente.CodigoEstado, "ICMS ST MG", Parent.NotaFiscal.Movimento, Parent.CodigoProduto)

                                        Dim undComercializao As String = String.Empty

                                        For Each itemPed In Parent.NotaFiscal.Pedido.Itens
                                            If itemPed.CodigoProduto = row("Produto") Then
                                                undComercializao = itemPed.CodigoUnidadeComercializacao
                                                Exit For
                                            End If
                                        Next

                                        Dim fatorConversao As Decimal

                                        If Parent.Produto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = undComercializao).Count() > 0 Then
                                            fatorConversao = Parent.Produto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = undComercializao).First.FatorConversao()
                                        Else
                                            fatorConversao = Parent.Produto.UnidadesDeComercializacao.FirstOrDefault().FatorConversao
                                        End If

                                        If fatorConversao > 0 Then
                                            vst = Math.Round(((Parent.QuantidadeFiscal * fatorConversao) * objEncargoXTaxa.SimplesNacional), 2, MidpointRounding.AwayFromZero)
                                        Else
                                            vst = Math.Round((Parent.QuantidadeFiscal * objEncargoXTaxa.SimplesNacional), 2, MidpointRounding.AwayFromZero)
                                        End If

                                        .Base = vst

                                        vlrST = Math.Round((.Base * objEncargoXTaxa.Percentual / 100), 2, MidpointRounding.AwayFromZero)
                                    Else
                                        Continue For
                                    End If

                                Else
                                    objEncargoXTaxa.SelecionarVigente(Parent.NotaFiscal.Cliente.CodigoEstado, .Codigo, Parent.NotaFiscal.Movimento, "")

                                    .PercentualSimplesNacional = objEncargoXTaxa.SimplesNacional

                                    Dim totalPRD As Decimal = Me.EncProduto.Valor

                                    If Not EncIPI Is Nothing AndAlso EncIPI.Valor > 0 Then
                                        totalPRD += EncIPI.Valor
                                    End If

                                    vst = Math.Round((totalPRD * .PercentualSimplesNacional) / 100, 2, MidpointRounding.AwayFromZero)

                                    .Base = totalPRD + vst

                                    vlrST = Math.Round((.Base * objEncargoXTaxa.Percentual / 100), 2, MidpointRounding.AwayFromZero)
                                End If

                                Dim perST As Decimal = 0
                                If Not EncICMS Is Nothing AndAlso EncICMS.Valor > 0 Then
                                    perST = Math.Round(((vlrST - EncICMS.Valor) / .Base) * 100, 2, MidpointRounding.AwayFromZero)
                                Else
                                    If Me.EncProduto.Valor > 0 Then
                                        perST = Math.Round(((Me.EncProduto.Valor - vlrST) / .Base) * 100, 2, MidpointRounding.AwayFromZero)
                                    End If
                                End If

                                .Percentual = perST

                                If Parent.SubOperacao.FinalidadeDaNota = 2 AndAlso EncICMS.Valor = 0 Then
                                    .Valor = vlrST

                                ElseIf Not EncICMS Is Nothing AndAlso EncICMS.Valor > 0 Then
                                    .Valor = vlrST - EncICMS.Valor

                                Else
                                    .Valor = vlrST - EncICMSDesonerado.Valor

                                End If
                            ElseIf row("Encargo_id") = "ICMS DIFERENCIAL" Then
                                If Not EncICMS Is Nothing Then
                                    .Base = Math.Round((Parent.ValorTotal * .BasePercentual) / 100, 2, MidpointRounding.AwayFromZero)
                                    .Valor = Math.Round(((.Base * row("Aliquota")) / 100) - EncICMS.Valor, 2, MidpointRounding.AwayFromZero)
                                    .Percentual = (.Valor / .Base) * 100
                                End If

                            ElseIf (row("Encargo_id") = "CBS" Or row("Encargo_id") = "IBS") AndAlso .ReducaoCBS_Perc > 0 Then

                                .Base = Math.Round((Parent.ValorTotal * .BasePercentual) / 100, 2, MidpointRounding.AwayFromZero)

                                Dim percentual = .BasePercentual - .ReducaoCBS_Perc
                                .Percentual = Math.Round((row("Aliquota") * percentual) / 100, 2, MidpointRounding.AwayFromZero)
                                .Valor = Math.Round((.Base * .Percentual) / 100, 2, MidpointRounding.AwayFromZero)

                            ElseIf row("PrecoDePauta") > 0 _
                             AndAlso Parent.QuantidadeFiscal > 0 _
                             AndAlso (Parent.QuantidadeFiscal * row("PrecoDePauta")) > Parent.ValorTotal _
                             AndAlso (Parent.Operacao.CodigoClasse = eClassesOperacoes.VENDAS.ToString Or Parent.Operacao.CodigoClasse = eClassesOperacoes.DEPOSITOS.ToString) _
                             AndAlso ((Parent.SubOperacao.EntradaSaida = eEntradaSaida.Saida And Not Parent.SubOperacao.Devolucao) OrElse (Parent.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso Parent.NotaFiscal.NossaEmissao)) Then
                                .Base = Math.Round(((Parent.QuantidadeFiscal * row("PrecoDePauta")) * .BasePercentual) / 100, 2, MidpointRounding.AwayFromZero)
                                .Percentual = row("Aliquota")
                                .PercentualExibicao = row("AliquotaExibicao")
                                .Valor = Math.Round((.Base * .Percentual) / 100, 2, MidpointRounding.AwayFromZero)
                            Else
                                .Base = Math.Round((Parent.ValorTotal * .BasePercentual) / 100, 2, MidpointRounding.AwayFromZero)
                                .Percentual = row("Aliquota")
                                .PercentualExibicao = row("AliquotaExibicao")
                                .Valor = Math.Round((.Base * .Percentual) / 100, 2, MidpointRounding.AwayFromZero)
                            End If
                        Else
                            .Percentual = row("Aliquota")
                            .PercentualExibicao = row("AliquotaExibicao")

                            If Parent.NotaFiscal.Empresa.CodigoEstado = "MT" AndAlso (.Codigo.Contains("FETHAB") Or .Codigo = "IAGRO") Then
                                If Parent.NotaFiscal.SubOperacao.Devolucao Then
                                    objEncargoXTaxa.SelecionarVigente(Parent.NotaFiscal.Empresa.CodigoEstado, .Codigo, Parent.NotaFiscal.DataDevolucao, .CodigoProduto)
                                Else
                                    objEncargoXTaxa.SelecionarVigente(Parent.NotaFiscal.Empresa.CodigoEstado, .Codigo, Parent.NotaFiscal.Movimento, .CodigoProduto)
                                End If

                                .Base = Parent.QuantidadeFiscal
                                .Valor = Math.Round(((CDec(objEncargoXTaxa.SimplesNacional) * .Percentual) / 100) * (Parent.QuantidadeFiscal / 1000), 2, MidpointRounding.AwayFromZero)
                            ElseIf .Codigo = "FUNDERSUL" OrElse .Codigo = "FUNDEMS" Then
                                objEncargoXTaxa.SelecionarVigente(Parent.NotaFiscal.Empresa.CodigoEstado, "UFERMS", Parent.NotaFiscal.Movimento, .CodigoProduto)
                                .Base = Parent.QuantidadeFiscal
                                .Valor = Math.Round(((CDec(objEncargoXTaxa.SimplesNacional) * .Percentual) / 100) * (Parent.QuantidadeFiscal / 1000), 2, MidpointRounding.AwayFromZero)
                            ElseIf .Encargo.Operador = "%" Then
                                If row("PrecoDePauta") > 0 AndAlso Parent.QuantidadeFiscal > 0 AndAlso (Parent.QuantidadeFiscal * row("PrecoDePauta")) > Parent.ValorTotal AndAlso (Parent.Operacao.CodigoClasse = eClassesOperacoes.VENDAS.ToString Or Parent.Operacao.CodigoClasse = eClassesOperacoes.DEPOSITOS.ToString) AndAlso Parent.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso Not Parent.SubOperacao.Devolucao Then
                                    .Base = Math.Round(((Parent.QuantidadeFiscal * row("PrecoDePauta")) * .BasePercentual) / 100, 2, MidpointRounding.AwayFromZero)
                                Else
                                    .Base = Math.Round((Parent.QuantidadeFiscal * .BasePercentual) / 100, 2, MidpointRounding.AwayFromZero)
                                End If
                                .Valor = Math.Round((.Base * .Percentual) / 1000, 2, MidpointRounding.AwayFromZero) 'Base para peso, nao mudar - Furlan
                            ElseIf .Encargo.Operador = "*" Then
                                .Base = Parent.QuantidadeFiscal
                                .Valor = Math.Round((.Base * .Percentual), 2, MidpointRounding.AwayFromZero)
                            End If
                        End If
                    End If
                    'Carrega o percentual de exibição
                    If .PercentualExibicao = 0 Then .PercentualExibicao = .Percentual

                    'If Not Parent.Retencao Then
                    '    sql &= "And isnull(En.PodeSofreRetencao,0) = 0"
                    'End If

                    Dim Tp As eTipoPessoa
                    If Parent.NotaFiscal.Cliente.CodigoEstado = "EX" Then
                        Parent.Retencao = False
                        Tp = eTipoPessoa.Ambas
                    ElseIf Parent.NotaFiscal.CodigoCliente.Length = 14 OrElse Parent.NotaFiscal.CodigoTipoDeDocumento = 57 Then
                        Tp = eTipoPessoa.Juridica
                    Else
                        Tp = eTipoPessoa.Fisica
                    End If

                    If objNFEncargo.Codigo = "FRETES" AndAlso Parent.ValorFreteXML > 0 Then
                        objNFEncargo.Valor = Parent.ValorFreteXML
                    End If

                    If objNFEncargo.Codigo = "DESCONTOS" AndAlso Parent.ValorDescontoXML > 0 Then
                        objNFEncargo.Valor = Parent.ValorDescontoXML
                    End If

                    If Not objNFEncargo.Encargo.PodeSofreRetencao Then
                        Me.Add(objNFEncargo)
                    ElseIf (objNFEncargo.Encargo.TipoPessoaRetencao = Tp Or objNFEncargo.Encargo.TipoPessoaRetencao = 3) Then
                        If Parent.Retencao Then
                            Me.Add(objNFEncargo)
                        End If
                    Else
                        Me.Add(objNFEncargo)
                    End If

                    If objNFEncargo.Codigo = "PRODUTO" Then Me.EncProduto = objNFEncargo
                    If objNFEncargo.Codigo = "LIQUIDO" Then Me.EncLiquido = objNFEncargo
                    If objNFEncargo.Codigo = "ICMS" Then EncICMS = objNFEncargo
                    If objNFEncargo.Codigo = "ICMS DESONERADO" Then EncICMSDesonerado = objNFEncargo
                    If objNFEncargo.Codigo = "DESP.ADUANEIRAS" Then EncADUANEIRAS = objNFEncargo
                    If objNFEncargo.Codigo = "FRETES" Then EncFrete = objNFEncargo

                End With
            Next

            If Left(Parent.NotaFiscal.Empresa.Codigo, 8) = "24450490" AndAlso Not Parent.NotaFiscal.SubOperacao Is Nothing AndAlso Parent.NotaFiscal.SubOperacao.Classe = eClassesOperacoes.IMPORTACOES AndAlso Parent.NotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                For Each objEncargo As NotaFiscalXItemXEncargo In Me
                    If objEncargo.Codigo = "ICMS" Then
                        objEncargoXTaxa.SelecionarVigente(Parent.NotaFiscal.Empresa.CodigoEstado, "ICMS", Parent.NotaFiscal.Movimento, "")

                        Dim vlrFrete As Decimal = 0
                        Dim vlrADUANEIRAS As Decimal = 0

                        If Not EncFrete Is Nothing AndAlso EncFrete.Valor > 0 Then vlrFrete = EncFrete.Valor
                        If Not EncADUANEIRAS Is Nothing AndAlso EncADUANEIRAS.Valor > 0 Then vlrADUANEIRAS = EncADUANEIRAS.Valor

                        Dim vTotal As Decimal = EncProduto.Valor + vlrADUANEIRAS + vlrFrete

                        Dim vICMS As Decimal = Math.Round(((vTotal / objEncargoXTaxa.Percentual) * objEncargoXTaxa.SimplesNacional) / 100, 2, MidpointRounding.AwayFromZero)

                        objEncargo.Valor = vICMS

                        objEncargo.Base = Math.Round((vICMS * 100) / objEncargo.Percentual, 2, MidpointRounding.AwayFromZero)

                        EncICMS = objEncargo
                    End If
                Next

                For Each objEncargo As NotaFiscalXItemXEncargo In Me
                    If objEncargo.Codigo = "CUSTO ICMS" Then
                        objEncargo.Valor = EncICMS.Valor
                    End If
                Next
            End If

            'VALOR DA BASE DE CÁLCULO DO FUNDO FECP DEVE SER IGUAL AO ICMS ST - FURLAN - 21/01/2025.

            If Parent.NotaFiscal.IUD = "I" Then
                Dim BaseST As Decimal = 0
                Dim temFUNDOFECP As Boolean = False
                For Each enc As NotaFiscalXItemXEncargo In Me
                    If enc.Codigo.Contains("ICMS-ST") Then BaseST = enc.Base
                    If enc.Codigo.Contains("FUNDO FECP") Then temFUNDOFECP = True
                Next

                If temFUNDOFECP AndAlso BaseST > 0 Then
                    For Each enc As NotaFiscalXItemXEncargo In Me
                        If enc.Codigo.Contains("FUNDO FECP") Then
                            'RECALCULA
                            enc.Base = BaseST
                            enc.Valor = Math.Round(enc.Base * (enc.Percentual / 100), 2, MidpointRounding.AwayFromZero)
                        End If
                    Next
                End If
            End If

            If AtualLiquido Then AtualizaLiquido()
        Else
            HttpContext.Current.Session("ssMessage") = Parent.Produto.ToString & " - " & Parent.Produto.Nome & ", Não possui encargos para a Operação Selecionada"
        End If
    End Sub

    Public Sub CarregaListaConhecimento()
        Dim sql As String = ""
        Dim AtualLiquido As Boolean = False

        sql = "Select OE.GrupoProduto," & vbCrLf &
              "       OE.Produto," & vbCrLf &
              "       OE.Operacao," & vbCrLf &
              "       OE.SubOperacao," & vbCrLf &
              "       OE.EstadoOrigem," & vbCrLf &
              "       OE.EstadoDestino," & vbCrLf &
              "       NxE.Encargo_Id," & vbCrLf &
              "       OE.STICMS," & vbCrLf &
              "       OE.STIPI," & vbCrLf &
              "       OE.STPISCOFINS," & vbCrLf &
              "       ISNULL(OE.UsarCalculadoraDeImposto,0) AS UsarCalculadoraDeImposto," & vbCrLf &
              "       ISNULL(OE.STIBSCBS,0) AS STIBSCBS," & vbCrLf &
              "       ISNULL(OE.ClassificacaoIBSCBS,0) AS ClassificacaoIBSCBS," & vbCrLf &
              "       ISNULL(OE.ReducaoIBS_Perc,0) AS ReducaoIBS_Perc," & vbCrLf &
              "       ISNULL(OE.ReducaoCBS_Perc,0) AS ReducaoCBS_Perc," & vbCrLf &
              "       ISNULL(OE.ObsPISCOFINS,0) ObsPISCOFINS, " & vbCrLf &
              "       en.ValorOuPeso," & vbCrLf &
              "       OE.GrupoFiscal," & vbCrLf &
              "       OE.CodigoFiscal," & vbCrLf &
              "       OEE.DebitaConta," & vbCrLf &
              "       OEE.CreditaConta," & vbCrLf &
              "       OEE.Sinal," & vbCrLf &
              "       OEE.AliquotaBase," & vbCrLf &
              "       isnull(encXtx.Percentual, OEE.Aliquota) As Aliquota," & vbCrLf &
              "       isnull(OEE.AliquotaExibicao,0) As AliquotaExibicao," & vbCrLf &
              "       isnull(OEE.AliquotaLimite,0) As AliquotaLimite," & vbCrLf &
              "       NxE.Base," & vbCrLf &
              "       NxE.Percentual," & vbCrLf &
              "       NxE.Valor," & vbCrLf &
              "       isnull(OT.Descricao,'') as ObservacaoFiscal," & vbCrLf &
              "       isnull(encXtx.SimplesNacional,0) as SimplesNacional," & vbCrLf &
              "       isnull(NxE.CentroDeCusto,0) as CentroDeCusto," & vbCrLf &
              "       isnull(PDP.Preco,0) as PrecoDePauta," & vbCrLf &
              "       NxE.CFOP_Id," & vbCrLf &
              "       isnull(NxE.BaseNova,0) as BaseNova, isnull(NxE.ValorNovo,0) as ValorNovo" & vbCrLf &
              "  FROM NotasFiscaisXEncargos NxE" & vbCrLf &
              " INNER JOIN OperacaoXEstado OE" & vbCrLf &
              "    ON OE.Codigo_Id = " & Me.Parent.CodigoOperacaoEstado & vbCrLf &
              " INNER Join OperacaoXEstadoXEncargo OEE" & vbCrLf &
              "    on OEE.Codigo_Id  = Oe.Codigo_Id" & vbCrLf &
              "   and OEE.Encargo_Id = NxE.Encargo_Id" & vbCrLf &
              "  left join (SELECT ext.Encargo_Id as Encargo, ext.Percentual, ext.Estado_Id, isnull(ext.SimplesNacional,0) as SimplesNacional" & vbCrLf &
              "               FROM EncargosXTaxas ext" & vbCrLf &
              "              WHERE Data_Id = (SELECT MAX(Data_Id)" & vbCrLf &
              "                                 FROM EncargosXTaxas" & vbCrLf &
              "                				   WHERE EncargosXTaxas.Encargo_Id = ext.Encargo_Id" & vbCrLf &
              "                					 AND EncargosXTaxas.Estado_Id  = ext.Estado_Id" & vbCrLf &
               "               				     AND EncargosXTaxas.Produto_Id = '" & Parent.CodigoProduto.ToString & "'" & vbCrLf &
              "                					 AND Data_id <= '" & Parent.NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "') " & vbCrLf &
              "                			       ) encXtx " & vbCrLf &
              "    on encXtx.Encargo    = OEE.Encargo_id" & vbCrLf &
              "   AND encXtx.Estado_Id  = OE.EstadoOrigem" & vbCrLf &
              "  Left join (SELECT 'ICMS' as Encargo, PP.Produto_id, PP.Estado_Id, Preco" & vbCrLf &
              "               FROM PrecoDePauta PP" & vbCrLf &
              "              WHERE Data_Id = (SELECT MAX(Data_Id)" & vbCrLf &
              "                                 FROM PrecoDePauta" & vbCrLf &
              "                				   WHERE PrecoDePauta.Estado_Id  = PP.Estado_Id" & vbCrLf &
              "                					 AND PrecoDePauta.Produto_Id = PP.Produto_Id" & vbCrLf &
              "                					 AND PrecoDePauta.Data_id   <= '" & Parent.NotaFiscal.Movimento.ToString("yyyy-MM-dd") & "') " & vbCrLf &
              "                			      ) PDP" & vbCrLf &
              "    on PDP.Encargo    = OEE.Encargo_id" & vbCrLf &
              "   AND PDP.Estado_Id  = OE.EstadoOrigem" & vbCrLf &
              "   AND PDP.Produto_Id = NxE.Produto_Id" & vbCrLf &
              " Inner Join Encargos En" & vbCrLf &
              "    on En.Encargo_id =  NxE.Encargo_Id" & vbCrLf &
              "  left join ObservacoesTributarias OT" & vbCrLf &
              "    on OE.ObsICMS     = OT.Codigo_Id" & vbCrLf &
              "   and OEE.Encargo_Id = OT.Encargo" & vbCrLf &
              " WHERE NxE.Empresa_id      ='" & Parent.NotaFiscal.CodigoEmpresa & "' " & vbCrLf &
              "   AND NxE.EndEmpresa_id   = " & Parent.NotaFiscal.EnderecoEmpresa & vbCrLf &
              "   AND NxE.Cliente_Id      ='" & Parent.NotaFiscal.CodigoCliente & "' " & vbCrLf &
              "   AND NxE.EndCliente_id   = " & Parent.NotaFiscal.EnderecoCliente & vbCrLf &
              "   AND NxE.EntradaSaida_Id ='" & Parent.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "' " & vbCrLf &
              "   AND NxE.Serie_id        ='" & Parent.NotaFiscal.Serie & "' " & vbCrLf &
              "   AND NxE.Nota_id         = " & Parent.NotaFiscal.Codigo & vbCrLf &
              "   AND NxE.Produto_id      ='" & Parent.CodigoProduto & "' " & vbCrLf &
              "   AND NxE.Sequencia_Id    = " & Parent.Sequencia & vbCrLf &
              "   AND NxE.Nota_id         > 0" & vbCrLf &
              "  order by " & vbCrLf &
              "        case " & vbCrLf &
              "           When OEE.Encargo_Id = 'PRODUTO' " & vbCrLf &
              "             Then 1 " & vbCrLf &
              "            When OEE.Encargo_Id = 'LIQUIDO' or OEE.Encargo_Id = 'LIQUIDOAPAGAR' " & vbCrLf &
              "             Then 3 " & vbCrLf &
              "            Else 2 " & vbCrLf &
              "        end "

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Encargos")

        If ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim objNFEncargo As New Negocio.NotaFiscalXItemXEncargo(Me.Parent)

                With objNFEncargo

                    If row("Encargo_id") = "PRODUTO" Then
                        Parent.CFOP = row("CFOP_Id")
                        AtualLiquido = True
                    End If

                    .Existe = True
                    .Codigo = row("Encargo_id")

                    .SituacaoTributaria = row("STICMS")
                    .SituacaoTributariaIPI = row("STIPI")
                    .SituacaoTributariaPISCOFINS = row("STPISCOFINS")
                    .SituacaoTributariaPISCOFINSOBS = row("ObsPISCOFINS")

                    .UsarCalculadoraDeImposto = row("UsarCalculadoraDeImposto")
                    .SituacaoTributariaIBSCBS = row("STIBSCBS")
                    .ClassificacaoIBSCBS = row("ClassificacaoIBSCBS")
                    .ReducaoIBS_Perc = row("ReducaoIBS_Perc")
                    .ReducaoCBS_Perc = row("ReducaoCBS_Perc")

                    .CodigoOperacao = row("Operacao")
                    .CodigoSubOperacao = row("SubOperacao")
                    .CodigoGrupoProduto = row("GrupoProduto")
                    .CodigoProduto = row("Produto")
                    .EstadoOrigem = row("EstadoOrigem")
                    .EstadoDestino = row("EstadoDestino")
                    .BasePercentual = row("AliquotaBase")
                    .PercentualLimite = row("AliquotaLimite")
                    .PercentualEncagosXTaxa = row("Percentual")
                    .PercentualSimplesNacional = row("SimplesNacional")
                    .Sinal = row("Sinal")
                    .ObservacaoFiscal = row("ObservacaoFiscal")
                    .ValorPeso = IIf(row("ValorouPeso"), 1, 0)
                    .CentroDeCusto = row("CentroDeCusto")

                    .BaseNova = row("BaseNova")
                    .ValorNovo = row("ValorNovo")

                    If row("Valor") > 0 Then
                        .Base = row("Base")
                        .Percentual = row("Percentual")
                        .PercentualExibicao = row("AliquotaExibicao")
                        .Valor = row("Valor")
                    Else
                        If .ValorPeso = eValorPeso.Valor Then
                            If row("PrecoDePauta") > 0 AndAlso Parent.QuantidadeFiscal > 0 AndAlso (Parent.QuantidadeFiscal * row("PrecoDePauta")) > Parent.ValorTotal AndAlso Parent.Operacao.CodigoClasse = eClassesOperacoes.VENDAS.ToString AndAlso Parent.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso Not Parent.SubOperacao.Devolucao Then
                                .Base = Math.Round(((Parent.QuantidadeFiscal * row("PrecoDePauta")) * .BasePercentual) / 100, 2, MidpointRounding.AwayFromZero)
                            Else
                                .Base = Math.Round((Parent.ValorTotal * .BasePercentual) / 100, 2, MidpointRounding.AwayFromZero)
                            End If
                            .Percentual = row("Aliquota")
                            .PercentualExibicao = row("AliquotaExibicao")
                            .Valor = Math.Round((.Base * .Percentual) / 100, 2, MidpointRounding.AwayFromZero)
                        Else
                            .Percentual = row("Aliquota")
                            .PercentualExibicao = row("AliquotaExibicao")
                            If .Encargo.Operador = "%" Then
                                If row("PrecoDePauta") > 0 AndAlso Parent.QuantidadeFiscal > 0 AndAlso (Parent.QuantidadeFiscal * row("PrecoDePauta")) > Parent.ValorTotal AndAlso Parent.Operacao.CodigoClasse = eClassesOperacoes.VENDAS.ToString AndAlso Parent.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso Not Parent.SubOperacao.Devolucao Then
                                    .Base = Math.Round(((Parent.QuantidadeFiscal * row("PrecoDePauta")) * .BasePercentual) / 100, 2, MidpointRounding.AwayFromZero)
                                Else
                                    .Base = Math.Round((Parent.QuantidadeFiscal * .BasePercentual) / 100, 2, MidpointRounding.AwayFromZero)
                                End If
                                .Valor = Math.Round((.Base * .Percentual) / 1000, 2, MidpointRounding.AwayFromZero) 'Base para peso, nao mudar - Furlan
                            ElseIf .Encargo.Operador = "*" Then
                                .Base = Parent.QuantidadeFiscal
                                .Valor = Math.Round((.Base * .Percentual), 2, MidpointRounding.AwayFromZero)
                            End If
                        End If
                    End If

                    'If .Codigo = "FACS" Or .Codigo = "FETHAB" Or .Codigo.Contains("FUNRURAL") Or .Codigo.Contains("SENAR") Or .Codigo.Contains("LIMINARAGRICOLA") Or .Codigo = ("FABOV") Or .Codigo = "FETHAB GADO" Then
                    '    If Parent.NotaFiscal.Cliente.CodigoCategoria < 4 Then
                    '        Me.Add(objNFEncargo)
                    '    End If
                    'Else
                    Me.Add(objNFEncargo)
                    If objNFEncargo.Codigo = "PRODUTO" Then Me.EncProduto = objNFEncargo
                    If objNFEncargo.Codigo = "LIQUIDO" Then Me.EncLiquido = objNFEncargo
                    'End If
                End With
            Next
            If AtualLiquido Then AtualizaLiquido()
        Else
            HttpContext.Current.Session("ssMessage") = Parent.Produto.ToString & " - " & Parent.Produto.Nome & ", Não possui encargos para a Operação Selecionada"
        End If
    End Sub
#End Region

#Region "Propriedades"
    Public Property Parent() As Negocio.NotaFiscalXItem
        Get
            Return _Parent
        End Get
        Set(ByVal value As Negocio.NotaFiscalXItem)
            _Parent = value
        End Set
    End Property

    Public Property EncProduto() As NotaFiscalXItemXEncargo
        Get
            Return _EncProduto
        End Get
        Set(ByVal value As NotaFiscalXItemXEncargo)
            _EncProduto = value
        End Set
    End Property

    Public Property EncLiquido() As NotaFiscalXItemXEncargo
        Get
            Return _EncLiquido
        End Get
        Set(ByVal value As NotaFiscalXItemXEncargo)
            _EncLiquido = value
        End Set
    End Property

    Public ReadOnly Property MensagemImpostos As String
        Get
            Dim msg As String = ""
            Dim LEnc As New ListEncargo(True, "ImprimirNFE = 1")

            For Each row As NotaFiscalXItemXEncargo In Me
                Dim codigo As String = row.Codigo
                Dim x = LEnc.Find(Function(s) s.Codigo = codigo)
                If Not x Is Nothing Then
                    msg &= IIf(String.IsNullOrWhiteSpace(msg), "Impostos: ", " , ") & row.Codigo & " = " & row.Valor
                End If
            Next
            Return msg
        End Get
    End Property
#End Region

#Region "Métodos"

    Public Sub AtualizaLiquido(Optional ByVal SaldoInicial As Decimal = 0)
        If SaldoInicial > 0 Then _SaldoInicial = SaldoInicial
        If Me.Count = 0 Then Exit Sub
        Parent.CarregandoEncargos = True

        Dim Total As Decimal = 0
        Dim IPI As Decimal = 0
        Dim SomaProduto As Decimal = 0

        Dim EncargoProduto As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoLiquido As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoLiquidoAPagar As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoIcms As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoIcmsST As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoIPI As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoIPIRecuperar As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoDesconto As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoComissao As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoFretes As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoFunrural As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoSenar As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoPIS As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoPISRecuperar As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoCOFINS As NotaFiscalXItemXEncargo = Nothing
        Dim EncargoCOFINSRecuperar As NotaFiscalXItemXEncargo = Nothing
        Dim EncADUANEIRAS As NotaFiscalXItemXEncargo = Nothing
        Dim EncCustoCofins As NotaFiscalXItemXEncargo = Nothing
        Dim EncCustoPis As NotaFiscalXItemXEncargo = Nothing

        For Each Enc As NotaFiscalXItemXEncargo In Me
            Select Case Enc.Codigo.Trim
                Case "LIQUIDO" : EncargoLiquido = Enc
                Case "DESCONTOS" : EncargoDesconto = Enc
                Case "FUNRURAL JUDICIAL" : EncargoFunrural = Enc
                Case "FUNRURAL" : EncargoFunrural = Enc
                Case "SENAR" : EncargoSenar = Enc
                Case "IPI" : EncargoIPI = Enc
                Case "IPI A RECUP." : EncargoIPIRecuperar = Enc
                Case "ICMS" : EncargoIcms = Enc
                Case "ICMS-ST" : EncargoIcmsST = Enc
                Case "LIQUIDOAPAGAR" : EncargoLiquidoAPagar = Enc
                Case "COMISSAO" : EncargoComissao = Enc
                Case "PRODUTO" : EncargoProduto = Enc
                Case "FRETES" : EncargoFretes = Enc
                Case "PIS" : EncargoPIS = Enc
                Case "CUSTO PIS" : EncCustoPis = Enc
                Case "PIS RECUP." : EncargoPISRecuperar = Enc
                Case "COFINS" : EncargoCOFINS = Enc
                Case "CUSTO COFINS" : EncCustoCofins = Enc
                Case "COFINS RECUP." : EncargoCOFINSRecuperar = Enc
                Case "DESP.ADUANEIRAS" : EncADUANEIRAS = Enc
            End Select

            'If Enc.Sinal = "=+" And Not Enc.Presumido Then
            '    SomaProduto = Enc.Valor
            'End If

            If Left(Parent.NotaFiscal.CodigoEmpresa, 8) = "24450490" AndAlso Parent.NotaFiscal.DiferencaValorNFXProdutoXML <> 0 Then
                If Enc.Codigo.Trim = "DESCONTOS" AndAlso Enc.Valor = 0 AndAlso Parent.NotaFiscal.DiferencaValorNFXProdutoXML < 0 Then
                    Enc.Valor = Math.Abs(Parent.NotaFiscal.DiferencaValorNFXProdutoXML)
                ElseIf Enc.Codigo.Trim = "DESP. ACESSORIA" AndAlso Enc.Valor = 0 AndAlso Parent.NotaFiscal.DiferencaValorNFXProdutoXML > 0 Then
                    Enc.Valor = Math.Abs(Parent.NotaFiscal.DiferencaValorNFXProdutoXML)
                End If
            End If

            If Enc.Sinal = "+" And Enc.Codigo.Trim <> "IPI" Then
                Total += Enc.Valor
            ElseIf Enc.Sinal = "-" And Enc.Codigo.Trim <> "IPI" Then
                Total -= Enc.Valor
            End If

        Next

        If Not EncargoIPI Is Nothing AndAlso Me.Parent.CodigoPedido > 0 Then
            Dim encPed As Negocio.PedidoXEncargo
            Dim itemPed As Negocio.PedidoXItem
            For Each itemPed In Parent.Pedido.Itens
                If Me.Parent.CodigoProduto = itemPed.CodigoProduto Then
                    For Each encPed In itemPed.Encargos
                        If encPed.CodigoEncargo.Trim = "IPI" Then
                            EncargoIPI.Percentual = encPed.Percentual
                            'EncargoIPI.Valor = EncargoIPI.Base * encPed.Percentual / 100
                            If Not EncargoIPIRecuperar Is Nothing Then
                                EncargoIPIRecuperar.Percentual = EncargoIPI.Percentual
                                EncargoIPIRecuperar.Valor = EncargoIPI.Valor
                            End If
                        End If
                    Next
                End If
            Next
            If EncargoIPI.Sinal = "+" Then Total += EncargoIPI.Valor

        ElseIf Not EncargoIPI Is Nothing AndAlso Me.Parent.NotaFiscal.NFG AndAlso EncargoIPI.Sinal = "+" Then

            Total += EncargoIPI.Valor

        End If

        If Not EncargoIcms Is Nothing Then
            Dim vlrfrete As Decimal = 0
            Dim vlrdesconto As Decimal = 0
            If Not EncargoFretes Is Nothing Then vlrfrete = EncargoFretes.Valor
            'If Not EncargoDesconto Is Nothing Then vlrdesconto = EncargoDesconto.Valor - 14/05/2012 - Furlan - Divergência entre operação Curtume e Alvorada
            'If vlrfrete > 0 OrElse vlrdesconto > 0 Then
            If Not EncargoFretes Is Nothing AndAlso EncargoFretes.Encargo.Atualizacao AndAlso vlrfrete > 0 Then
                EncargoIcms.Base = Math.Round((EncargoProduto.Valor + vlrfrete - vlrdesconto) * EncargoIcms.BasePercentual / 100, 2, MidpointRounding.AwayFromZero)
                EncargoIcms.Valor = Math.Round((EncargoIcms.Base * EncargoIcms.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
            End If
        End If

        If Not EncargoFunrural Is Nothing Then
            Dim vlrfrete As Decimal = 0
            Dim vlrdesconto As Decimal = 0
            If Not EncargoFretes Is Nothing Then vlrfrete = EncargoFretes.Valor
            If Not EncargoDesconto Is Nothing Then vlrdesconto = EncargoDesconto.Valor
            If vlrfrete > 0 OrElse vlrdesconto > 0 Then
                Total += EncargoFunrural.Valor
                EncargoFunrural.Base = Math.Round((EncargoProduto.Valor + vlrfrete - vlrdesconto) * EncargoFunrural.BasePercentual / 100, 2, MidpointRounding.AwayFromZero)
                EncargoFunrural.Valor = Math.Round((EncargoFunrural.Base * EncargoFunrural.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
                Total -= EncargoFunrural.Valor
            End If
        End If

        If Not EncargoSenar Is Nothing Then
            Dim vlrfrete As Decimal = 0
            Dim vlrdesconto As Decimal = 0
            If Not EncargoFretes Is Nothing Then vlrfrete = EncargoFretes.Valor
            If Not EncargoDesconto Is Nothing Then vlrdesconto = EncargoDesconto.Valor
            If vlrfrete > 0 OrElse vlrdesconto > 0 Then
                Total += EncargoSenar.Valor
                EncargoSenar.Base = Math.Round((EncargoProduto.Valor + vlrfrete - vlrdesconto) * EncargoSenar.BasePercentual / 100, 2, MidpointRounding.AwayFromZero)
                EncargoSenar.Valor = Math.Round((EncargoSenar.Base * EncargoSenar.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
                Total -= EncargoSenar.Valor
            End If
        End If

        'If Not EncargoIcmsST Is Nothing Then
        '    If EncargoIcmsST.PercentualLimite > 0 Then
        '        EncargoIcmsST.Base = (EncargoProduto.Base + IPI) * EncargoIcmsST.PercentualSimplesNacional
        '        EncargoIcmsST.Valor = Math.Round(((EncargoIcmsST.Base * EncargoIcmsST.Percentual) / 100) - EncargoIcms.Valor, 2, MidpointRounding.AwayFromZero)
        '        'Else
        '        '    EncargoIcmsST.Base = (EncargoProduto.Base + IPI) * EncargoIcmsST.PercentualEncagosXTaxa
        '    End If
        'End If

        If Not EncargoPIS Is Nothing Then
            Dim vlrfrete As Decimal = 0
            Dim vlrdesconto As Decimal = 0
            If Not EncargoFretes Is Nothing Then vlrfrete = EncargoFretes.Valor
            If Not EncargoDesconto Is Nothing Then vlrdesconto = EncargoDesconto.Valor
            If vlrfrete > 0 OrElse vlrdesconto > 0 Then
                EncargoPIS.Base = Math.Round((EncargoProduto.Valor + vlrfrete - vlrdesconto) * EncargoPIS.BasePercentual / 100, 2, MidpointRounding.AwayFromZero)
                EncargoPIS.Valor = Math.Round((EncargoPIS.Base * EncargoPIS.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
            End If
        End If

        If Not EncCustoPis Is Nothing Then
            Dim vlrfrete As Decimal = 0
            Dim vlrdesconto As Decimal = 0
            If Not EncargoFretes Is Nothing Then vlrfrete = EncargoFretes.Valor
            If Not EncargoDesconto Is Nothing Then vlrdesconto = EncargoDesconto.Valor
            If vlrfrete > 0 OrElse vlrdesconto > 0 Then
                EncCustoPis.Base = Math.Round((EncargoProduto.Valor + vlrfrete - vlrdesconto) * EncCustoPis.BasePercentual / 100, 2, MidpointRounding.AwayFromZero)
                EncCustoPis.Valor = Math.Round((EncCustoPis.Base * EncCustoPis.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
            End If
        End If

        If Not EncargoPISRecuperar Is Nothing Then
            Dim vlrfrete As Decimal = 0
            Dim vlrdesconto As Decimal = 0
            If Not EncargoFretes Is Nothing Then vlrfrete = EncargoFretes.Valor
            If Not EncargoDesconto Is Nothing Then vlrdesconto = EncargoDesconto.Valor
            If vlrfrete > 0 OrElse vlrdesconto > 0 Then
                EncargoPISRecuperar.Base = Math.Round((EncargoProduto.Valor + vlrfrete - vlrdesconto) * EncargoPISRecuperar.BasePercentual / 100, 2, MidpointRounding.AwayFromZero)
                EncargoPISRecuperar.Valor = Math.Round((EncargoPISRecuperar.Base * EncargoPISRecuperar.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
            End If
        End If

        If Not EncargoCOFINS Is Nothing Then
            Dim vlrfrete As Decimal = 0
            Dim vlrdesconto As Decimal = 0
            If Not EncargoFretes Is Nothing Then vlrfrete = EncargoFretes.Valor
            If Not EncargoDesconto Is Nothing Then vlrdesconto = EncargoDesconto.Valor
            If vlrfrete > 0 OrElse vlrdesconto > 0 Then
                EncargoCOFINS.Base = Math.Round((EncargoProduto.Valor + vlrfrete - vlrdesconto) * EncargoCOFINS.BasePercentual / 100, 2, MidpointRounding.AwayFromZero)
                EncargoCOFINS.Valor = Math.Round((EncargoCOFINS.Base * EncargoCOFINS.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
            End If
        End If

        If Not EncCustoCofins Is Nothing Then
            Dim vlrfrete As Decimal = 0
            Dim vlrdesconto As Decimal = 0
            If Not EncargoFretes Is Nothing Then vlrfrete = EncargoFretes.Valor
            If Not EncargoDesconto Is Nothing Then vlrdesconto = EncargoDesconto.Valor
            If vlrfrete > 0 OrElse vlrdesconto > 0 Then
                EncCustoCofins.Base = Math.Round((EncargoProduto.Valor + vlrfrete - vlrdesconto) * EncCustoCofins.BasePercentual / 100, 2, MidpointRounding.AwayFromZero)
                EncCustoCofins.Valor = Math.Round((EncCustoCofins.Base * EncCustoCofins.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
            End If
        End If

        If Not EncargoCOFINSRecuperar Is Nothing Then
            Dim vlrfrete As Decimal = 0
            Dim vlrdesconto As Decimal = 0
            If Not EncargoFretes Is Nothing Then vlrfrete = EncargoFretes.Valor
            If Not EncargoDesconto Is Nothing Then vlrdesconto = EncargoDesconto.Valor
            If vlrfrete > 0 OrElse vlrdesconto > 0 Then
                EncargoCOFINSRecuperar.Base = Math.Round((EncargoProduto.Valor + vlrfrete - vlrdesconto) * EncargoCOFINSRecuperar.BasePercentual / 100, 2, MidpointRounding.AwayFromZero)
                EncargoCOFINSRecuperar.Valor = Math.Round((EncargoCOFINSRecuperar.Base * EncargoCOFINSRecuperar.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
            End If
        End If

        If Not EncargoProduto Is Nothing Then EncargoProduto.Valor += SomaProduto
        Total += SomaProduto

        If Not EncargoLiquido Is Nothing Then
            EncargoLiquido.Base = Total
            EncargoLiquido.Valor = Total
        End If

        If Not EncargoLiquidoAPagar Is Nothing Then
            EncargoLiquidoAPagar.Base = Total + _SaldoInicial
            EncargoLiquidoAPagar.Valor = Total + _SaldoInicial
        End If
        Parent.CarregandoEncargos = False

        Me.EncProduto = EncargoProduto
        Me.EncLiquido = EncargoLiquido

        'If TemDesconto = True Then
        '    For Each Enc As NotaFiscalXItemXEncargo In Me
        '        If Enc.Codigo <> "LIQUIDO" And Enc.Codigo <> "PRODUTO" And Enc.Codigo <> "FRETES" And Enc.Codigo <> "DESCONTOS" Then
        '            If Enc.Sinal = "-" Or Enc.Sinal = "=" Then
        '                Parent.CarregandoEncargos = True
        '                Enc.Base = Math.Round(Total, 2, MidpointRounding.AwayFromZero)
        '                Enc.Valor = Math.Round((Enc.Base * Enc.Percentual) / 100, 2, MidpointRounding.AwayFromZero)
        '                Parent.CarregandoEncargos = False
        '            End If
        '        End If
        '    Next
        'End If
    End Sub

    Public Sub SalvarSql(ByVal Sqls As ArrayList)

        'ISSO DEVE FICAR ANTES DO ITEM, DÚVIDAS CHAMAR - FURLAN - 28/08/2025
        'If Parent.IUD = "U" Then

        '    Dim sql As String

        '    If Parent.CodigoProdutoOld Is Nothing OrElse Parent.CodigoProdutoOld.Length = 0 Then

        '        sql = " Delete NotasFiscaisXEncargos " & vbCrLf &
        '              "  Where Empresa_Id      ='" & Parent.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
        '              "    and EndEmpresa_Id   = " & Parent.NotaFiscal.EnderecoEmpresa & vbCrLf &
        '              "	   and Cliente_Id      ='" & Parent.NotaFiscal.CodigoCliente & "'" & vbCrLf &
        '              "    and EndCliente_Id   = " & Parent.NotaFiscal.EnderecoCliente & vbCrLf &
        '              "    and EntradaSaida_Id ='" & Parent.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
        '              "    and Nota_Id         = " & Parent.NotaFiscal.Codigo & vbCrLf &
        '              "    and Serie_Id        ='" & Parent.NotaFiscal.Serie & "'" & vbCrLf &
        '              "    and Produto_Id      ='" & Parent.CodigoProduto & "'" & vbCrLf &
        '              "    and Sequencia_Id    = " & Parent.Sequencia & ";"
        '        '"    and CFOP_Id         = " & Parent.CFOP & vbCrLf & _

        '    Else

        '        sql = " Delete NotasFiscaisXEncargos " & vbCrLf &
        '              "  Where Empresa_Id      ='" & Parent.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
        '              "    and EndEmpresa_Id   = " & Parent.NotaFiscal.EnderecoEmpresa & vbCrLf &
        '              "	   and Cliente_Id      ='" & Parent.NotaFiscal.CodigoCliente & "'" & vbCrLf &
        '              "    and EndCliente_Id   = " & Parent.NotaFiscal.EnderecoCliente & vbCrLf &
        '              "    and EntradaSaida_Id ='" & Parent.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
        '              "    and Nota_Id         = " & Parent.NotaFiscal.Codigo & vbCrLf &
        '              "    and Serie_Id        ='" & Parent.NotaFiscal.Serie & "'" & vbCrLf &
        '              "    --and Produto_Id      ='" & Parent.CodigoProdutoOld & "'" & vbCrLf &
        '              "    and Sequencia_Id    = " & Parent.Sequencia & ";"
        '        '"    and CFOP_Id         = " & Parent.CFOP & vbCrLf & _

        '    End If

        '    Sqls.Add(sql)
        'End If

        For Each Enc As Negocio.NotaFiscalXItemXEncargo In Me
            If Parent.IUD = "U" Then
                Enc.IUD = "I"
            ElseIf Parent.IUD = "D" Or Parent.IUD = "I" Then
                Enc.IUD = Parent.IUD
            End If

            If Enc.IUD <> "" Then
                Enc.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Function Clone() As ListNotaFiscalXItemXEncargo
        Dim encargos As ListNotaFiscalXItemXEncargo = CType(Me.MemberwiseClone(), ListNotaFiscalXItemXEncargo)
        Return encargos
    End Function

#End Region

End Class

<Serializable()>
Public Class NotaFiscalXItemXEncargo
    Implements IBaseEntity

#Region "Construtores"
    Public Sub New()
    End Sub

    Public Sub New(ByVal Item As Negocio.NotaFiscalXItem)
        _ItemNotaFiscal = Item
    End Sub
#End Region

#Region "Fields"
    Private _ItemNotaFiscal As Negocio.NotaFiscalXItem
    Private _IUD As String
    Private _Existe As Boolean
    Private _Codigo As String
    Private _SituacaoTributaria As Integer
    Private _SituacaoTributariaIPI As Integer
    Private _SituacaoTributariaPISCOFINS As Integer
    Private _SituacaoTributariaPISCOFINSOBS As Integer
    Private _SituacaoTributariaIBSCBS As Integer
    Private _CodigoOperacao As Integer
    Private _Operacao As Negocio.Operacao
    Private _CodigoSubOperacao As Integer
    Private _SubOperacao As Negocio.SubOperacao
    Private _CodigoGrupoProduto As String
    Private _GrupoProduto As Negocio.GrupoProduto
    Private _CodigoProduto As String
    Private _Produto As Negocio.Produto
    Private _EstadoOrigem As String
    Private _EstadoDestino As String
    Private _BasePercentual As Decimal
    Private _Base As Decimal
    Private _Percentual As Decimal
    Private _PercentualExibicao As Decimal
    Private _PercentualLimite As Decimal
    Private _PercentualSimplesNacional As Decimal
    Private _PercentualEncagosXTaxa As Decimal
    Private _Valor As Decimal
    Private _CentroDeCusto As String = "0"
    Private _Sinal As String
    Private _ValorPeso As Negocio.eValorPeso
    Private _DescricaoFiscal As String = ""
    Private _ObservacaoFiscal As String
    Private _Presumido As Boolean
    Private _Encargo As Negocio.Encargo
    Private _Sequencia As Integer
    Private _GrupoDeContaDebito As PlanoDeConta
    Private _GrupoDeContaCredito As PlanoDeConta
    Private _BaseNova As Decimal
    Private _ValorNovo As Decimal
    Private _ClassificacaoIBSCBS As String = ""
    Private _UsarCalculadoraDeImposto As Boolean
    Private _ReducaoIBS_Perc As Decimal
    Private _ReducaoCBS_Perc As Decimal

#End Region

#Region "Propriedades"

    Public Property ItemNotaFiscal() As Negocio.NotaFiscalXItem
        Get
            Return _ItemNotaFiscal
        End Get
        Set(ByVal value As Negocio.NotaFiscalXItem)
            _ItemNotaFiscal = value
        End Set
    End Property

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Existe() As Boolean
        Get
            Return _Existe
        End Get
        Set(ByVal value As Boolean)
            _Existe = value
        End Set
    End Property

    Public Property Codigo() As String
        Get
            Return _Codigo
        End Get
        Set(ByVal value As String)
            _Codigo = value
        End Set
    End Property

    Public Property SituacaoTributaria() As Integer
        Get
            Return _SituacaoTributaria
        End Get
        Set(ByVal value As Integer)
            _SituacaoTributaria = value
        End Set
    End Property

    Public Property SituacaoTributariaIPI() As Integer
        Get
            Return _SituacaoTributariaIPI
        End Get
        Set(ByVal value As Integer)
            _SituacaoTributariaIPI = value
        End Set
    End Property

    Public Property SituacaoTributariaPISCOFINS() As Integer
        Get
            Return _SituacaoTributariaPISCOFINS
        End Get
        Set(ByVal value As Integer)
            _SituacaoTributariaPISCOFINS = value
        End Set
    End Property

    Public Property SituacaoTributariaPISCOFINSOBS() As Integer
        Get
            Return _SituacaoTributariaPISCOFINSOBS
        End Get
        Set(ByVal value As Integer)
            _SituacaoTributariaPISCOFINSOBS = value
        End Set
    End Property

    Public Property CodigoOperacao() As Integer
        Get
            Return _CodigoOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoOperacao = value
        End Set
    End Property

    Public Property Operacao() As Negocio.Operacao
        Get
            If _Operacao Is Nothing And _CodigoOperacao > 0 Then _Operacao = New Negocio.Operacao(_CodigoOperacao)
            Return _Operacao
        End Get
        Set(ByVal value As Negocio.Operacao)
            _Operacao = value
        End Set
    End Property

    Public Property CodigoSubOperacao() As Integer
        Get
            Return _CodigoSubOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSubOperacao = value
        End Set
    End Property

    Public Property SubOperacao() As Negocio.SubOperacao
        Get
            If _SubOperacao Is Nothing And _CodigoOperacao > 0 And _CodigoSubOperacao > 0 Then _SubOperacao = New Negocio.SubOperacao(_CodigoOperacao, _CodigoSubOperacao)
            Return _SubOperacao
        End Get
        Set(ByVal value As Negocio.SubOperacao)
            _SubOperacao = value
        End Set
    End Property

    Public Property CodigoGrupoProduto() As String
        Get
            Return _CodigoGrupoProduto
        End Get
        Set(ByVal value As String)
            _CodigoGrupoProduto = value
        End Set
    End Property

    Public Property GrupoProduto() As Negocio.GrupoProduto
        Get
            If _GrupoProduto Is Nothing And _CodigoGrupoProduto > 0 Then _GrupoProduto = New Negocio.GrupoProduto(_CodigoGrupoProduto)
            Return _GrupoProduto
        End Get
        Set(ByVal value As Negocio.GrupoProduto)
            _GrupoProduto = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property Produto() As Negocio.Produto
        Get
            If _Produto Is Nothing And Not String.IsNullOrEmpty(_CodigoProduto) Then _Produto = New Negocio.Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Negocio.Produto)
            _Produto = value
        End Set
    End Property

    Public Property EstadoOrigem() As String
        Get
            Return _EstadoOrigem
        End Get
        Set(ByVal value As String)
            _EstadoOrigem = value
        End Set
    End Property

    Public Property EstadoDestino() As String
        Get
            Return _EstadoDestino
        End Get
        Set(ByVal value As String)
            _EstadoDestino = value
        End Set
    End Property

    Public Property BasePercentual() As Decimal
        Get
            Return _BasePercentual
        End Get
        Set(ByVal value As Decimal)
            _BasePercentual = value
        End Set
    End Property

    Public Property PercentualExibicao() As Decimal
        Get
            'If _PercentualExibicao = 0 Then
            '    Return _Percentual
            'Else
            Return _PercentualExibicao
            'End If
        End Get
        Set(ByVal value As Decimal)
            _PercentualExibicao = value
        End Set
    End Property

    Public Property Base() As Decimal
        Get
            Return _Base
        End Get
        Set(ByVal value As Decimal)
            _Base = value
        End Set
    End Property

    Public Property Percentual() As Decimal
        Get
            Return _Percentual
        End Get
        Set(ByVal value As Decimal)
            _Percentual = value
        End Set
    End Property

    Public Property Valor() As Decimal
        Get
            Return _Valor
        End Get
        Set(ByVal value As Decimal)
            _Valor = value
            If Not ItemNotaFiscal.CarregandoEncargos Then ItemNotaFiscal.Encargos.AtualizaLiquido()
        End Set
    End Property

    Public Property PercentualEncagosXTaxa() As Decimal
        Get
            Return _PercentualEncagosXTaxa
        End Get
        Set(ByVal value As Decimal)
            _PercentualEncagosXTaxa = value
        End Set
    End Property

    Public Property PercentualSimplesNacional() As Decimal
        Get
            Return _PercentualSimplesNacional
        End Get
        Set(ByVal value As Decimal)
            _PercentualSimplesNacional = value
        End Set
    End Property

    Public Property PercentualLimite() As Decimal
        Get
            Return _PercentualLimite
        End Get
        Set(ByVal value As Decimal)
            _PercentualLimite = value
        End Set
    End Property

    Public Property CentroDeCusto() As String
        Get
            Return _CentroDeCusto
        End Get
        Set(ByVal value As String)

            If value.Trim = "" Then
                _CentroDeCusto = "0"
            Else
                _CentroDeCusto = value
            End If

        End Set
    End Property

    Public Property Sinal() As String
        Get
            Return _Sinal
        End Get
        Set(ByVal value As String)
            _Sinal = value
        End Set
    End Property

    Public Property ValorPeso() As eValorPeso
        Get
            Return _ValorPeso
        End Get
        Set(ByVal value As eValorPeso)
            _ValorPeso = value
        End Set
    End Property

    Public Property DescricaoFiscal() As String
        Get
            Return _DescricaoFiscal
        End Get
        Set(ByVal value As String)
            _DescricaoFiscal = value
        End Set
    End Property

    Public Property ObservacaoFiscal() As String
        Get
            Return _ObservacaoFiscal
        End Get
        Set(ByVal value As String)
            _ObservacaoFiscal = value
        End Set
    End Property

    Public ReadOnly Property OperacaoEncargo() As Negocio.OperacaoXEstadoXEncargo
        Get
            If Not ItemNotaFiscal.OperacaoEstado Is Nothing Then
                Return ItemNotaFiscal.OperacaoEstado.Encargos.Where(Function(s) s.CodigoEncargo = Me.Codigo).FirstOrDefault
            Else
                Return New Negocio.OperacaoXEstadoXEncargo()
            End If
        End Get
    End Property

    Public Property Encargo() As Negocio.Encargo
        Get
            If _Encargo Is Nothing And _Codigo.Length > 0 Then _Encargo = New Encargo(_Codigo)
            Return _Encargo
        End Get
        Set(ByVal value As Negocio.Encargo)
            _Encargo = value
        End Set
    End Property

    'Public Property Presumido() As Boolean
    '    Get
    '        Return _Presumido
    '    End Get
    '    Set(ByVal value As Boolean)
    '        _Presumido = value
    '    End Set
    'End Property

    Public ReadOnly Property ContaDeDebito() As String
        Get
            If Not OperacaoEncargo Is Nothing Then
                Return OperacaoEncargo.CodigoDebitaConta
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property GrupoDeContaDebito() As PlanoDeConta
        Get
            If _GrupoDeContaDebito Is Nothing AndAlso Not ContaDeDebito Is Nothing AndAlso ContaDeDebito.Length > 0 Then _GrupoDeContaDebito = New PlanoDeConta("", 0, ContaDeDebito)
            Return _GrupoDeContaDebito
        End Get
    End Property

    Public ReadOnly Property ContaDeCredito() As String
        Get
            Return OperacaoEncargo.CodigoCreditaConta
        End Get
    End Property

    Public ReadOnly Property GrupoDeContaCredito() As PlanoDeConta
        Get
            If _GrupoDeContaCredito Is Nothing And ContaDeCredito.Length > 0 Then _GrupoDeContaCredito = New PlanoDeConta("", 0, ContaDeCredito)
            Return _GrupoDeContaCredito
        End Get
    End Property

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
        End Set
    End Property

    Public Property BaseNova() As Decimal
        Get
            Return _BaseNova
        End Get
        Set(ByVal value As Decimal)
            _BaseNova = value
        End Set
    End Property

    Public Property ValorNovo() As Decimal
        Get
            Return _ValorNovo
        End Get
        Set(ByVal value As Decimal)
            _ValorNovo = value
        End Set
    End Property

    Public Property SituacaoTributariaIBSCBS As Integer
        Get
            Return _SituacaoTributariaIBSCBS
        End Get
        Set(value As Integer)
            _SituacaoTributariaIBSCBS = value
        End Set
    End Property

    Public Property ClassificacaoIBSCBS As String
        Get
            Return _ClassificacaoIBSCBS
        End Get
        Set(value As String)
            _ClassificacaoIBSCBS = value
        End Set
    End Property

    Public Property UsarCalculadoraDeImposto As Boolean
        Get
            Return _UsarCalculadoraDeImposto
        End Get
        Set(value As Boolean)
            _UsarCalculadoraDeImposto = value
        End Set
    End Property

    Public Property ReducaoIBS_Perc As Decimal
        Get
            Return _ReducaoIBS_Perc
        End Get
        Set(value As Decimal)
            _ReducaoIBS_Perc = value
        End Set
    End Property

    Public Property ReducaoCBS_Perc As Decimal
        Get
            Return _ReducaoCBS_Perc
        End Get
        Set(value As Decimal)
            _ReducaoCBS_Perc = value
        End Set
    End Property

#End Region

#Region "Methods"

    Public Function SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case IUD
            Case "I"
                sql = "INSERT INTO NotasFiscaisXEncargos (Empresa_Id, EndEmpresa_Id, " & vbCrLf &
                                                         "Cliente_Id, EndCliente_Id, " & vbCrLf &
                                                         "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf &
                                                         "Produto_Id, CFOP_Id, Sequencia_id, " & vbCrLf &
                                                         "Encargo_Id, Base, Percentual,  Valor, CentroDeCusto, BaseNova, ValorNovo)"
                sql &= "                          VALUES ('" & ItemNotaFiscal.NotaFiscal.CodigoEmpresa & "', " & ItemNotaFiscal.NotaFiscal.EnderecoEmpresa & ", " & vbCrLf &
                                                         "'" & ItemNotaFiscal.NotaFiscal.CodigoCliente & "', " & ItemNotaFiscal.NotaFiscal.EnderecoCliente & ", " & vbCrLf &
                                                         "'" & ItemNotaFiscal.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & ItemNotaFiscal.NotaFiscal.Serie & "', " & ItemNotaFiscal.NotaFiscal.Codigo & ", " & vbCrLf &
                                                         "'" & ItemNotaFiscal.CodigoProduto & "', " & ItemNotaFiscal.CFOP & "," & ItemNotaFiscal.Sequencia & " ," & vbCrLf &
                                                         "'" & Codigo & "'," & Str(Me.Base) & "," & Str(Me.Percentual) & "," & Str(Me.Valor) & ",'" & Me.CentroDeCusto & "'," & Str(Me.BaseNova) & "," & Str(Me.ValorNovo) & ")"
                If Valor > 0 Or Codigo = "PRODUTO" Or Codigo = "LIQUIDO" Then Sqls.Add(sql)
            Case "D"

                If ItemNotaFiscal.CodigoProdutoOld Is Nothing OrElse ItemNotaFiscal.CodigoProdutoOld.Length = 0 Then

                    ItemNotaFiscal.CodigoProdutoOld = ItemNotaFiscal.CodigoProduto

                End If

                If ItemNotaFiscal.CFOPOld = 0 Then

                    ItemNotaFiscal.CFOPOld = ItemNotaFiscal.CFOP

                End If

                sql = " Delete NotasFiscaisXEncargos " & vbCrLf &
                      "  Where Empresa_Id      ='" & ItemNotaFiscal.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                      "    and EndEmpresa_Id   = " & ItemNotaFiscal.NotaFiscal.EnderecoEmpresa & vbCrLf &
                      "	   and Cliente_Id      ='" & ItemNotaFiscal.NotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "    and EndCliente_Id   = " & ItemNotaFiscal.NotaFiscal.EnderecoCliente & vbCrLf &
                      "    and EntradaSaida_Id ='" & ItemNotaFiscal.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "    and Nota_Id         = " & ItemNotaFiscal.NotaFiscal.Codigo & vbCrLf &
                      "    and Serie_Id        ='" & ItemNotaFiscal.NotaFiscal.Serie & "'" & vbCrLf &
                      "    and Produto_Id      ='" & ItemNotaFiscal.CodigoProdutoOld & "'" & vbCrLf &
                      "    and Sequencia_Id    = " & ItemNotaFiscal.Sequencia & vbCrLf &
                      "    and Encargo_Id      ='" & _Codigo & "';"

                Sqls.Add(sql)

        End Select
        Return Sqls
    End Function

    Public Function Clone() As NotaFiscalXItemXEncargo
        Dim encargo As NotaFiscalXItemXEncargo = CType(Me.MemberwiseClone(), NotaFiscalXItemXEncargo)
        Return encargo
    End Function

#End Region

End Class