Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Web.UI.WebControls
Imports NGS.Lib.Uteis
Imports System.Web
Imports System.Text

Public Class CarregarDDL

#Region "Atributos"
    Private ds As DataSet
#End Region

#Region "Propriedades"
    Public ReadOnly Property Banco() As AcessaBanco
        Get
            Return New AcessaBanco()
        End Get
    End Property
#End Region

#Region "ENUM"
    Public Enum Tabela
        AjustesDaApuracaoIcms
        Analises
        Ano
        AtivosXContas
        Bancos
        BancosXContaContabil
        CarteirasXTributos
        CategoriaCliente
        CCustoECF
        CCustoECD
        Cultura
        ClasseAmbiental
        ClasseDeRisco
        ClasseToxicologica
        ClasseDeOperacao
        Clientes
        ClientesXTipos
        CentroDeCusto
        CentroDeCustoDescricao
        Cidades
        ClientesXEmpresas
        CarteiraFinanceira
        CarteiraFinanceiraConta
        CarteiraDoTitulo
        CondicaoDePagamento
        ContasBancarias
        ContasOrcamentarias
        ControleDeQualidade
        CFOP
        CFOPGrupo
        CNAE
        Dacon
        Depositos
        DepositosPedido
        Estados
        EstadoERegiao
        EstadosUF
        Empresas
        EmpresasMonitor
        EmpresasConsolidadas
        Embalagem
        EstadoFisicoIA
        EmbalagemTipoQtdeDoProduto
        Encargos
        EPI
        EspecificacaoDoProduto
        Etapas
        FormulacaoFito
        FiliaisPedido
        Finalidade
        GeneroDoProduto
        GeneroDoProdutoXSub
        GruposDeAtivos
        GrupoQuimico
        GrupoProduto
        GrupoProdutoXConsumo
        GrupoProdutoGeral
        Historico
        HistoricoContabil
        IngredianteAtivo
        Indexador
        IndexGtinProduto
        LoteContabil
        LoteProduto
        LoteClassificacaoProduto
        Marca
        Mes
        MicroRegiao
        Moeda
        Motivo
        ModeloCertidaoNegativa
        NaturezaJuridica
        Navios
        NaviosXInvoice
        NivelGrupoProduto
        Operacao
        OperacaoSubOperacao
        OperacaoSubOperacaoPermitidasNaNota
        ObservacaoTributariaGeral
        Pais
        ProcedimentoParaProducao
        Processos
        Produto
        ProdutoProducao
        ProdutoPorGrupo
        Provisoes
        PlanoDeContas
        PlanoDeContasComEncargos
        PlanoDeCustos
        ReferencialBacen
        Regiao
        ResponsabilidadeDaConta
        Safra
        Seguimento
        Servidores
        Seguimentos
        Situacao
        SituacaoTributariaICMS
        SituacaoTributariaIPI
        SituacaoTributariaPISCOFINS
        SituacaoTributariaIBSCBS
        SituacaoTributariaPISCOFINSObs
        SubOperacaoSemOperacao
        TabelaDeClassificacoes
        TabelaDePreco
        TipoDeCertidao
        TipoDeEmbalagem
        TipoDaContaContabil
        TipoDeDocumento
        TipoDeFaturamento
        TipoDePagamento
        TipoDeCliente
        TipoDeVeiculo
        UnidadeDeNegocio
        UnidadeDeMedida
        Usuarios
        ViaDeTransportes
        ViaDeTransportesSefaz
    End Enum
#End Region

#Region "Contrutor"
    Public Sub New()

    End Sub
#End Region

#Region "Methods Public"
    Public Sub Carregar(ByRef ddl As Object, ByVal tbl As Tabela, Optional ByVal Where As String = "", Optional ByVal LinhaEmBranco As Boolean = True, Optional ByVal Parametros As Hashtable = Nothing)
        Select Case tbl
            Case Tabela.AjustesDaApuracaoIcms : CarregarAjustesDaApuracaoIcms(ddl, Where)
            Case Tabela.Ano : CarregarAno(ddl, Where)
            Case Tabela.Analises : CarregarAnalises(ddl, Where)
            Case Tabela.AtivosXContas : CarregarAtivosXContas(ddl, Where)
            Case Tabela.Bancos : CarregarBancos(ddl, Where)
            Case Tabela.BancosXContaContabil : CarregarBancosXContaContabil(ddl, Where)
            Case Tabela.CarteirasXTributos : CarregarCarteirasXTributos(ddl, Where)
            Case Tabela.CategoriaCliente : CarregarCategoriaCliente(ddl, Where)
            Case Tabela.CCustoECF : carregarCCustoECF(ddl, Where)
            Case Tabela.CCustoECD : carregarCCustoECD(ddl, Where)
            Case Tabela.Cultura : CarregarCultura(ddl, Where)
            Case Tabela.CarteiraDoTitulo : CarregarCarteiraDoTitulo(ddl, Where)
            Case Tabela.CarteiraFinanceira : CarregarCarteiraFinanceira(ddl, Where)
            Case Tabela.CarteiraFinanceiraConta : CarregarCarteiraFinanceiraConta(ddl, Where)
            Case Tabela.CentroDeCusto : CarregarCentroDeCusto(ddl, Where)
            Case Tabela.CentroDeCustoDescricao : CarregarCentroDeCustoDescricao(ddl, Where)
            Case Tabela.Cidades : CarregarCidades(ddl, Where)
            Case Tabela.ClasseToxicologica : CarregarClasseToxicologica(ddl, Where)
            Case Tabela.ClasseAmbiental : CarregarClasseAmbiental(ddl, Where)
            Case Tabela.ClasseDeRisco : CarregarClasseDeRisco(ddl, Where)
            Case Tabela.ClasseDeOperacao : CarregarClasseDeOperacao(ddl, Where)
            Case Tabela.Clientes : CarregarClientes(ddl, Where)
            Case Tabela.ClientesXTipos : CarregarClientesTipo(ddl, Where)
            Case Tabela.ClientesXEmpresas : CarregarClientesXEmpresas(ddl, Where)
            Case Tabela.CondicaoDePagamento : CarregarCondicaoDePagamento(ddl, Where)
            Case Tabela.ContasBancarias : CarregarContasBancarias(ddl, Where)
            Case Tabela.ContasOrcamentarias : CarregarContasOrcamentarias(ddl, Where)
            Case Tabela.ControleDeQualidade : CarregarControleDeQualidade(ddl, Where)
            Case Tabela.CFOP : CarregarCFOP(ddl, Where)
            Case Tabela.CFOPGrupo : CarregarCFOPGrupo(ddl, Where)
            Case Tabela.CNAE : CarregarCNAE(ddl, Where)
            Case Tabela.Dacon : CarregarDacon(ddl, Where)
            Case Tabela.Depositos : CarregarDepositos(ddl, Where)
            Case Tabela.DepositosPedido : CarregarDepositosPedido(ddl, Where, Parametros)
            Case Tabela.Estados : CarregarEstados(ddl, Where)
            Case Tabela.EstadoERegiao : CarregarEstadoERegiao(ddl, Where)
            Case Tabela.EstadosUF : CarregarEstadosUF(ddl, Where)
            Case Tabela.Empresas : CarregarEmpresas(ddl, Where)
            Case Tabela.EmpresasMonitor : CarregarEmpresasMonitor(ddl, Where)
            Case Tabela.EmpresasConsolidadas : CarregarEmpresasConsolidadas(ddl, Where)
            Case Tabela.Embalagem : CarregarEmbalagem(ddl, Where)
            Case Tabela.EstadoFisicoIA : CarregarEstadoFisicoIA(ddl, Where)
            Case Tabela.EmbalagemTipoQtdeDoProduto : CarregarEmbalagemTipoQtdeDoProduto(ddl, Where)
            Case Tabela.EspecificacaoDoProduto : CarregarEspecificacaoDoProduto(ddl, Where)
            Case Tabela.Encargos : CarregarEncargos(ddl, Where)
            Case Tabela.EPI : CarregarEPI(ddl, Where)
            Case Tabela.Etapas : CarregarEtapas(ddl, Where)
            Case Tabela.FormulacaoFito : CarregarFormulacaoDoFito(ddl, Where)
            Case Tabela.FiliaisPedido : CarregarFiliaisPedido(ddl, Where, Parametros)
            Case Tabela.Finalidade : CarregarFinalidade(ddl, Where)
            Case Tabela.GeneroDoProduto : CarregarGeneroDoProduto(ddl, Where)
            Case Tabela.GeneroDoProdutoXSub : CarregarGeneroDoProdutoXSub(ddl, Where)
            Case Tabela.GruposDeAtivos : CarregarGruposDeAtivos(ddl, Where)
            Case Tabela.GrupoQuimico : CarregarGrupoQuimico(ddl, Where)
            Case Tabela.GrupoProduto : CarregarGrupoProduto(ddl, Where)
            Case Tabela.GrupoProdutoXConsumo : CarregarGrupoProdutoXConsumo(ddl, Where)
            Case Tabela.GrupoProdutoGeral : CarregarGrupoProdutoGeral(ddl, Where)
            Case Tabela.HistoricoContabil : CarregarHistoricos(ddl, Where)
            Case Tabela.Historico : CarregarHistoricosDataList(ddl, Where)
            Case Tabela.IngredianteAtivo : CarregarIngredianteAtivo(ddl, Where)
            Case Tabela.Indexador : CarregarIndexador(ddl, Where)
            Case Tabela.IndexGtinProduto : CarregarIndexGtinProduto(ddl, Where)
            Case Tabela.LoteContabil : CarregarLoteContabil(ddl, Where)
            Case Tabela.LoteProduto : CarregarLoteProduto(ddl, Where)
            Case Tabela.LoteClassificacaoProduto : CarregarLoteClassificacaoProduto(ddl, Where)
            Case Tabela.Marca : CarregarMarca(ddl, Where)
            Case Tabela.Mes : CarregarMes(ddl, Where)
            Case Tabela.Regiao : CarregarRegiao(ddl, Where)
            Case Tabela.MicroRegiao : CarregarMicroRegiao(ddl, Where)
            Case Tabela.Moeda : CarregarMoeda(ddl, Where)
            Case Tabela.Motivo : CarregarMotivo(ddl, Where)
            Case Tabela.ModeloCertidaoNegativa : CarregarModeloCertidaoNegativa(ddl, Where)
            Case Tabela.NaturezaJuridica : CarregarNaturezaJuridica(ddl, Where)
            Case Tabela.NivelGrupoProduto : CarregarNivelGrupoProduto(ddl, Where)
            Case Tabela.OperacaoSubOperacao : CarregarOperacaoSubOperacao(ddl, Where)
            Case Tabela.SubOperacaoSemOperacao : CarregarSubOperacaoSemOperacao(ddl, Where)
            Case Tabela.OperacaoSubOperacaoPermitidasNaNota : CarregarOperacaoSubOperacaoPermitidasNaNota(ddl, Where, Parametros)
            Case Tabela.Operacao : CarregarOperacao(ddl, Where)
            Case Tabela.ObservacaoTributariaGeral : CarregaObservacaoTributariaGeral(ddl, Where)
            Case Tabela.Pais : CarregarPais(ddl, Where)
            Case Tabela.ProcedimentoParaProducao : CarregarProcedimentoParaProducao(ddl, Where)
            Case Tabela.Processos : CarregarProcessos(ddl, Where)
            Case Tabela.Produto : CarregarProduto(ddl, Where)
            Case Tabela.ProdutoProducao : CarregarProdutoProducao(ddl, Where)
            Case Tabela.ProdutoPorGrupo : CarregarProdutoPorGrupo(ddl, Where)
            Case Tabela.Provisoes : CarregarProvisoes(ddl, Where)
            Case Tabela.PlanoDeContas : CarregarPlanoDeContas(ddl, Where)
            Case Tabela.PlanoDeContasComEncargos : CarregarPlanoDeContasComEncargos(ddl, Where)
            Case Tabela.PlanoDeCustos : CarregarPlanoDeCustos(ddl, Where)
            Case Tabela.ReferencialBacen : CarregarReferencialBacen(ddl, Where)
            Case Tabela.Regiao : CarregarRegiao(ddl, Where)
            Case Tabela.ResponsabilidadeDaConta : CarregarResponsabilidadeDaConta(ddl, Where)
            Case Tabela.Safra : CarregarSafra(ddl, Where)
            Case Tabela.Seguimento : CarregarSeguimento(ddl, Where)
            Case Tabela.Servidores : CarregarServidores(ddl, Where)
            Case Tabela.Seguimentos : CarregarSeguimentos(ddl, Where)
            Case Tabela.Situacao : CarregarSituacao(ddl, Where)
            Case Tabela.SituacaoTributariaICMS : CarregarSituacaoICMS(ddl, Where)
            Case Tabela.SituacaoTributariaIPI : CarregarSituacaoIPI(ddl, Where)
            Case Tabela.SituacaoTributariaPISCOFINS : CarregarSituacaoPISCOFINS(ddl, Where)
            Case Tabela.SituacaoTributariaIBSCBS : CarregarSituacaoIBSCBS(ddl, Where)
            Case Tabela.SituacaoTributariaPISCOFINSObs : CarregarSituacaoPISCOFINSObs(ddl, Where)
            Case Tabela.TabelaDeClassificacoes : CarregarTabelaDeClassificacoes(ddl, Where)
            Case Tabela.TabelaDePreco : CarregarTabelaDePreco(ddl, Where)
            Case Tabela.TipoDeCliente : CarregarTipoDeCliente(ddl, Where)
            Case Tabela.TipoDaContaContabil : CarregarTipoDaContaContabil(ddl, Where)
            Case Tabela.TipoDeDocumento : CarregarTipoDeDocumento(ddl, Where)
            Case Tabela.TipoDeFaturamento : CarregarTipoDeFaturamento(ddl, Where)
            Case Tabela.TipoDeCertidao : CarregarTipoDeCertidao(ddl, Where)
            Case Tabela.TipoDeEmbalagem : CarregarTipoDeEmbalagem(ddl, Where)
            Case Tabela.TipoDePagamento : CarregarTipoDePagamento(ddl, Where, Parametros)
            Case Tabela.TipoDeVeiculo : CarregarTipoDeVeiculo(ddl, Where)
            Case Tabela.UnidadeDeNegocio : CarregarUnidadeDeNegocio(ddl, Where)
            Case Tabela.UnidadeDeMedida : CarregarUnidadeDeMedida(ddl, Where)
            Case Tabela.Usuarios : CarregarUsuarios(ddl, Where)
            Case Tabela.ViaDeTransportes : CarregarViaDeTransportes(ddl, Where)
            Case Tabela.ViaDeTransportesSefaz : CarregarViaDeTransportesSefaz(ddl, Where)
            Case Tabela.Navios : Navios(ddl, Where)
            Case Tabela.NaviosXInvoice : NaviosXInvoice(ddl, Where)
        End Select
        If LinhaEmBranco Then Funcoes.InserirLinhaEmBranco(ddl)
    End Sub

    Public Sub Carregar(ByRef ddl As Object, ByVal pOpcoes As String, ByVal Separador As String, Optional ByVal SegundoSeparador As String = "")
        Dim op() As String = pOpcoes.Split(Separador)
        For i As Integer = 0 To op.Length - 1
            If SegundoSeparador = String.Empty Then
                ddl.Items.Add(New System.Web.UI.WebControls.ListItem(op(i), op(i)))
            Else
                ddl.Items.Add(New System.Web.UI.WebControls.ListItem(op(i).Split(SegundoSeparador)(0) & " - " & op(i).Split(SegundoSeparador)(1), op(i).Split(SegundoSeparador)(0)))
            End If
        Next
    End Sub
#End Region

#Region "Methods Private"
    Private Sub CarregarAnalises(ByRef ddl As Object, ByVal Where As String)
        Dim objListAnalises As New ListAnalise()

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As Analise In objListAnalises
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr.Descricao, 50, ".") & " - " & Dr.Codigo, Dr.Codigo.ToString()))
        Next
    End Sub

    Private Sub CarregarAtivosXContas(ByRef ddl As Object, ByVal Where As String)
        Dim sql As String = " select pl.Conta_Id, pl.Titulo " & vbCrLf &
                            "   from AtivosXContas ac" & vbCrLf &
                            "  Inner Join PlanoDeContas pl" & vbCrLf &
                            "     on pl.Conta_Id = ac.Conta_Id" & vbCrLf
        If Where.Length > 0 Then
            sql &= "  where ac.Empresa_Id = '" & Where & "'"
        End If

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "AtivosXContas").Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Conta_Id") & " - " & Dr("Titulo"), Dr("Conta_Id")))
        Next
    End Sub

    Private Sub carregarCCustoECF(ByRef ddl As Object, ByVal where As String)
        Dim strSQL As String = "Select Codigo_Id, Descricao" &
                               "  from PlanoDeCustosECF"
        If where.Length > 0 Then
            strSQL &= " Where " & where
        End If

        strSQL &= " ORDER BY Codigo_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "PlanoDeCustosECF").Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Codigo_Id") & " - " & Dr("Descricao"), Dr("Codigo_Id")))
        Next
    End Sub

    Private Sub carregarCCustoECD(ByRef ddl As Object, ByVal where As String)
        Dim strSQL As String = "Select Custo_Id, Descricao" &
                               "  from PlanoDeCustosECD"
        If where.Length > 0 Then
            strSQL &= " Where " & where
        End If

        strSQL &= " ORDER BY Custo_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "PlanoDeCustosECD").Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Custo_Id") & " - " & Dr("Descricao"), Dr("Custo_Id")))
        Next
    End Sub

    Private Sub CarregarEncargos(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Encargo_Id, Descricao " &
                               "FROM Encargos "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        ds = Banco.ConsultaDataSet(strSQL, "Encargos")
        ddl.Items.Clear()

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Encargo_Id") & " - " & Dr("Descricao"), Dr("Encargo_Id")))
        Next
    End Sub

    Private Sub CarregarViaDeTransportes(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Codigo_Id, Descricao " &
                               "FROM ViaDeTransportes "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        ds = Banco.ConsultaDataSet(strSQL, "ViaDeTransportes")
        ddl.Items.Clear()

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Codigo_Id") & " - " & Dr("Descricao"), Dr("Codigo_Id")))
        Next
    End Sub

    Private Sub CarregarViaDeTransportesSefaz(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT CodigoSefaz_Id, Descricao " &
                               "FROM ViaDeTransportes"

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " Order By CodigoSefaz_Id"

        ds = Banco.ConsultaDataSet(strSQL, "ViaDeTransportes")
        ddl.Items.Clear()

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("CodigoSefaz_Id") & " - " & Dr("Descricao"), Dr("CodigoSefaz_Id")))
        Next
    End Sub

    Private Sub Navios(ByRef ddl As Object, ByVal Where As String)

        Dim sql As String = "SELECT Codigo_Id, Descricao" & vbCrLf &
                            "FROM Navios" & vbCrLf &
                            "WHERE Ativo = 1" & vbCrLf &
                            " Order By Codigo_Id"

        ds = Banco.ConsultaDataSet(sql, "Navios")
        ddl.Items.Clear()

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Codigo_Id") & " - " & Dr("Descricao")))
        Next
    End Sub

    Private Sub NaviosXInvoice(ByRef ddl As Object, ByVal Where As String)

        Dim sql As String = "SELECT nxi.Navio_Id as Codigo_Id, n.Descricao " & vbCrLf & _
                            "FROM NaviosXInvoice nxi " & vbCrLf & _
                            "	INNER JOIN Navios n " & vbCrLf & _
                            "			ON n.Codigo_Id = nxi.Navio_Id " & vbCrLf & _
                            "	INNER JOIN NavioXInvoiceXProduto nxixp " & vbCrLf & _
                            "			ON nxixp.Codigo_Id  = nxi.Codigo_Id " & vbCrLf & _
                            "		   AND nxixp.Produto_Id = '" & Where & "'" & vbCrLf & _
                            "WHERE nxi.Ativo = 1 " & vbCrLf & _
                            "group by nxi.Navio_Id, n.Descricao"

        ds = Banco.ConsultaDataSet(sql, "Navios")
        ddl.Items.Clear()

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Codigo_Id") & " - " & Dr("Descricao")))
        Next
    End Sub

    Private Sub CarregarCondicaoDePagamento(ByRef ddl As Object, ByVal Where As String)
        Dim SQL As String

        SQL = "SELECT Pagamento_id, Descricao + '(' + convert(nvarchar,parcelas)  + ' Parcelas)' as Descricao, " &
                  "Case" & vbCrLf &
                  "   When isnull(AVista,0) = 0" & vbCrLf &
                  "      then 2" & vbCrLf &
                  "	     else 1" & vbCrLf &
                  "   End As Ordem" & vbCrLf &
                  "into #Pagamentos" & vbCrLf &
                  "  FROM Pagamentos " & vbCrLf &
                  "" & vbCrLf &
                  "select Pagamento_id, Descricao from #Pagamentos" & vbCrLf &
                  "order by Ordem, Pagamento_id" & vbCrLf

        If Where.Length > 0 Then
            SQL &= " Where " & Where
        End If

        ds = Banco.ConsultaDataSet(SQL, "CondicaoDePagamento")
        ddl.Items.Clear()

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Pagamento_Id") & " - " & Dr("Descricao"), Dr("Pagamento_Id")))
        Next
    End Sub

    Private Sub CarregarContasBancarias(ByRef ddl As Object, ByVal Where As String)
        Dim objReceber As New [Lib].Negocio.ListBancosXContas(True, Where)

        ddl.Items.Clear()

        For Each row As [Lib].Negocio.BancosXContas In objReceber
            Dim agencia As String = IIf(row.DigitoAgencia.Length > 0, Funcoes.AlinharDireita(row.Agencia, 4, " ") & "-" & row.DigitoAgencia, Funcoes.AlinharDireita(row.Agencia, 6, "0"))
            Dim conta As String = IIf(row.DigitoConta.Length > 0, Funcoes.AlinharDireita(row.Conta, 10, " ") & "-" & row.DigitoConta, Funcoes.AlinharDireita(row.Conta, 10, " "))
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(row.NomeBanco, 20, ".") & " - AG: " & agencia & " - CTA: " & conta,
                                 row.CodigoBanco & ";" & row.Agencia & ";" & row.DigitoAgencia & ";" & row.Conta & ";" & row.DigitoConta & ";" & row.TipoConta))
        Next
    End Sub

    Private Sub CarregarPlanoDeContas(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "Select Conta_id as codigo, convert(varchar,conta_id) + ' - ' + Titulo as Descricao " & vbCrLf &
                               "  From PlanodeContas "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        If Not Where.Contains("order by") Then strSQL &= " order by conta_id"

        ds = Banco.ConsultaDataSet(strSQL, "PlanodeContas")
        ddl.Items.Clear()

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Descricao"), Dr("codigo")))
        Next
    End Sub

    Private Sub CarregarPlanoDeContasComEncargos(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "Select PC.Conta_id as Codigo, PC.Titulo as Descricao " & vbCrLf &
                               "  From PlanodeContas PC " & vbCrLf &
                               " Where Exists (Select 1" & vbCrLf &
                               "                 From EncargosPlanoDeContas EPC" & vbCrLf &
                               "                Where EPC.Conta_Id = pc.conta_Id)" & vbCrLf

        If Where.Length > 0 Then
            strSQL &= " AND " & Where
        End If

        strSQL &= " order by PC.conta_id"

        ds = Banco.ConsultaDataSet(strSQL, "PlanodeContas")
        ddl.Items.Clear()

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("codigo") & " - " & Dr("Descricao"), Dr("codigo")))
        Next
    End Sub

    Private Sub CarregarContasOrcamentarias(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Codigo_Id AS Codigo, Descricao " &
                               "  FROM ContasOrcamentarias "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY Codigo_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "ContasOrcamentarias").Tables(0).Rows
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr("Codigo"), 6, ".") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub CarregarControleDeQualidade(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Codigo_Id AS Codigo, Descricao " &
                               "  FROM ControleDeQualidade "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY Codigo_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "ControleDeQualidade").Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub CarregarCFOP(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT GrupoCfop_Id,  Cfop_Id, Descricao " &
                               "  FROM Cfop "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY GrupoCfop_Id, Cfop_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "ControleDeQualidade").Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("GrupoCfop_Id") & "-" & Dr("Cfop_Id") & " - " & Dr("Descricao"), Dr("Cfop_Id")))
        Next
    End Sub

    Private Sub CarregarCFOPGrupo(ByRef ddl As Object, ByVal Where As String)
        Dim sql As String
        sql = "Select GrupoCFOP_Id, Descricao " & vbCrLf &
              "  from CFOPTitulo"

        If Where.Length > 0 Then
            sql &= " Where " & Where
        End If

        sql &= " Order By GrupoCFOP_Id "

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "GrupoCFOP")

        ddl.Items.Clear()
        For Each row As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(row("GrupoCFOP_Id") & "-" & row("Descricao"), row("GrupoCFOP_Id")))
        Next
    End Sub

    Private Sub CarregarCNAE(ByRef ddl As Object, ByVal Where As String)
        Dim sql As String
        sql = "Select CNAE_Id, Descricao" & vbCrLf &
              "  from CNAE"

        If Where.Length > 0 Then
            sql &= " Where " & Where
        End If

        sql &= " Order By CNAE_Id "

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "CNAE")

        ddl.Items.Clear()
        For Each row As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(row("CNAE_Id") & "-" & row("Descricao"), row("CNAE_Id")))
        Next
    End Sub

    Private Sub CarregarDacon(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Dacon_Id AS Codigo, Descricao" &
                               "  FROM Dacon"
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= "  ORDER BY Dacon_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "PlanoDeContas").Tables(0).Rows
            ddl.Items.Add(New ListItem(Funcoes.AlinharDireita(Dr("Codigo"), 4, "0") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub CarregarPlanoDeCustos(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Codigo_Id as Codigo, Descricao " &
                               "FROM PlanoDeCustos "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        ds = Banco.ConsultaDataSet(strSQL, "PlanoDeCustos")

        ddl.Items.Clear()

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub CarregarReferencialBacen(ByRef ddl As Object, ByVal where As String)
        Dim strSQL As String = "Select Conta_Id, Descricao" &
                               "  from PlanoDeContasReferencialBacen"
        If where.Length > 0 Then
            strSQL &= " Where " & where
        End If

        strSQL &= " ORDER BY Conta_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "PlanoDeContasReferencialBacen").Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Conta_Id") & " - " & Dr("Descricao"), Dr("Conta_Id")))
        Next
    End Sub

    Private Sub CarregarRegiao(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "Select Regiao_Id, Descricao" &
                               "  from Regioes"
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY Regiao_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Regiao").Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Regiao_Id") & " - " & Dr("Descricao"), Dr("Regiao_Id")))
        Next
    End Sub

    Private Sub CarregarMicroRegiao(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "Select MicroRegiao_Id, Descricao" &
                               "  from MicroRegioes"
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY MicroRegiao_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "MicroRegiao").Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("MicroRegiao_Id") & " - " & Dr("Descricao"), Dr("MicroRegiao_Id")))
        Next
    End Sub

    Private Sub CarregarResponsabilidadeDaConta(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Codigo_Id AS Codigo, Descricao " &
                               "  FROM Responsabilidades "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= "ORDER BY Codigo_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "Responsabilidades").Tables(0).Rows
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr("Codigo"), 6, ".") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next

    End Sub

    Private Sub CarregarCentroDeCusto(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT CentroDeCusto_Id, CentroDeCusto_Id + ' - ' + Descricao as Descricao FROM CentrosDeCustos Where Ativo = 1"

        If Where.Length > 0 Then
            strSQL &= " And " & Where
        End If

        ds = Banco.ConsultaDataSet(strSQL, "CentroDeCusto")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "CentroDeCusto_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarCentroDeCustoDescricao(ByRef ddl As Object, ByVal Where As String)
        Dim sql As String = "SELECT CentroDeCusto_Id, Descricao " &
                            "  FROM CentrosDeCustos Where Ativo = 1"

        If Where.Length > 0 Then
            sql &= " AND " & Where
        End If

        sql &= " ORDER BY Descricao"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Pagamentos")
        If Not ds Is Nothing Then
            For Each drCentroCusto As DataRow In ds.Tables(0).Rows
                ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(drCentroCusto("Descricao").ToString(), 50, ".") & " - " &
                                                      drCentroCusto("CentroDeCusto_Id").ToString(),
                                                      drCentroCusto("CentroDeCusto_Id").ToString()))
            Next
        End If
    End Sub

    Private Sub CarregarUnidadeDeNegocio(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf &
                               "  FROM Clientes C " & vbCrLf &
                               " INNER JOIN ClientesXTipos CT " & vbCrLf &
                               "    ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf &
                               " WHERE CT.Tipo_Id = 050 " & vbCrLf &
                               " ORDER BY Nome" & vbCrLf

        ds = Banco.ConsultaDataSet(strSQL, "UnidadeDeNegocio")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarEmpresas(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = " SELECT DISTINCT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Case When isnull(cxe.matriz,'N') = 'S' Then 'Matriz' else 'Filial' end MF" & vbCrLf &
                               "   FROM GruposXEmpresas " & vbCrLf &
                               "  INNER JOIN Clientes" & vbCrLf &
                               "     ON GruposXEmpresas.Cliente_Id    = Clientes.Cliente_Id " & vbCrLf &
                               "    AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
                               "  Inner Join ClientesXEmpresas cxe" & vbCrLf &
                               "     on cxe.Empresa_Id    = Clientes.Cliente_Id" & vbCrLf &
                               "    and cxe.EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf
        If Where.Length > 0 Then
            strSQL &= " Where GruposXEmpresas.Empresa_Id = '" & Where & "'" & vbCrLf
        End If
        strSQL &= " order by Clientes.Reduzido"

        ds = Banco.ConsultaDataSet(strSQL, "Empresas")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".") & " - " & Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".") & " " & Dr("Estado") & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Dr("Codigo")), 18, ".") & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido") & "-" & Dr("MF"), Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))))
            Next
        End If
    End Sub

    Private Sub CarregarEmpresasMonitor(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = " SELECT DISTINCT Clientes.Cliente_Id as Codigo, Clientes.Reduzido As Reduzido, Clientes.Cidade As Cidade, Clientes.Estado As Estado" & vbCrLf &
                               "   FROM GruposXEmpresas " & vbCrLf &
                               "  INNER JOIN Clientes" & vbCrLf &
                               "     ON GruposXEmpresas.Cliente_Id    = Clientes.Cliente_Id " & vbCrLf &
                               "    AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
                               "  Inner Join ClientesXEmpresas cxe" & vbCrLf &
                               "     on cxe.Empresa_Id    = Clientes.Cliente_Id" & vbCrLf &
                               "    and cxe.EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf
        If Where.Length > 0 Then
            strSQL &= " Where GruposXEmpresas.Empresa_Id = '" & Where & "'" & vbCrLf
        End If
        strSQL &= " order by Clientes.Reduzido"

        ds = Banco.ConsultaDataSet(strSQL, "Empresas")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                ddl.Items.Add(New ListItem(Dr("Codigo") & "-" & Dr("Reduzido") & "-" & Dr("Cidade") & "-" & Dr("Estado")))
            Next
        End If
    End Sub

    Private Sub CarregarEmpresasConsolidadas(ByRef ddl As Object, ByVal Where As String)
        Dim sql As String
        sql = "select distinct substring(ce.Empresa_Id,1,8) as Empresa_Id," & vbCrLf &
              "       (Select top(1) Nome " & vbCrLf &
              "          from Clientes" & vbCrLf &
              "         Where substring(Cliente_Id,1,8) = substring(ce.Empresa_Id,1,8)" & vbCrLf &
              "        ) as nome" & vbCrLf &
              "  From clientesxempresas ce" & vbCrLf

        If Where.Length > 0 Then
            sql &= " Where " & Where
        End If

        ds = Banco.ConsultaDataSet(sql, "Empresas")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Empresa_Id") & " - " & Dr("Nome"), Dr("Empresa_Id")))
        Next
    End Sub

    Private Sub CarregarClientesXEmpresas(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT CE.Empresa_Id, CE.EndEmpresa_Id, C.Reduzido, " & vbCrLf &
                 "       C.Nome, C.Cidade, C.Estado " & vbCrLf &
                 "  FROM ClientesXEmpresas CE " & vbCrLf &
                 " INNER JOIN Clientes C " & vbCrLf &
                 "    ON CE.Empresa_Id    = C.Cliente_Id  " & vbCrLf &
                 "   AND CE.EndEmpresa_Id = C.Endereco_Id " & vbCrLf

        If Where.Length > 0 Then
            strSQL &= "  Where " & Where & vbCrLf
        End If

        strSQL &= "Order by CE.Empresa_Id "

        ds = Banco.ConsultaDataSet(strSQL, "Empresas")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".") & " - " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Dr("Empresa_Id")), 18, ".") & "-" & CStr(Dr("EndEmpresa_Id")) & "-" & Dr("Reduzido") & "-" & Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".") & " " & Dr("Estado"), Dr("Empresa_Id") & "-" & CStr(Dr("EndEmpresa_Id"))))
        Next
    End Sub

    Private Sub CarregarDepositos(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT C.Cliente_Id as Codigo, C.Endereco_Id, C.Reduzido, C.Nome, C.Cidade, C.Estado " &
                 "  FROM Clientes C " &
                 " INNER JOIN ClientesXTipos CT " &
                 "    ON C.Cliente_Id = CT.Cliente_Id " &
                 "   AND C.Endereco_Id = CT.Endereco_Id " &
                 " WHERE CT.Tipo_Id = 3 " &
                 " ORDER BY Nome, Reduzido"

        ds = Banco.ConsultaDataSet(strSQL, "Depositos")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr("Nome"), 60, ".") & " - " & Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".") & " " & Dr("Estado") & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Dr("Codigo")), 18, ".") & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido"), Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))))
        Next
    End Sub

    Private Sub CarregarDepositosPedido(ByRef ddl As Object, ByVal Where As String, ByVal Parametros As Hashtable)
        'Parametros tratados
        '**************************************
        '*********** Chave Primaria ***********
        '**************************************
        'Empresa
        'EndEmpresa
        'Pedido
        '*************************************************************************************************************
        '********** Tipo Deposito DE - Deposito; OD - OrigemDestino; TR - Transbordo;  LE - Local de embarque ********
        '*************************************************************************************************************
        'TipoDeposito
        '**********************************
        '***** Adicionar este deposito ****
        '**********************************
        'Deposito
        'EndDeposito
        If Not Parametros.ContainsKey("Empresa") _
        Or Not Parametros.ContainsKey("EndEmpresa") _
        Or Not Parametros.ContainsKey("Pedido") _
        Or Not Parametros.ContainsKey("TipoDeposito") _
        Or Not Parametros.ContainsKey("Deposito") _
        Or Not Parametros.ContainsKey("EndDeposito") Then
            Exit Sub
        End If


        Dim strSQL As String
        strSQL = "Select sb.Cliente_id," & vbCrLf &
                 "	     sb.Endereco_Id," & vbCrLf &
                 "	     sb.Nome," & vbCrLf &
                 "	     sb.Cidade," & vbCrLf &
                 "	     sb.Estado," & vbCrLf &
                 "	     sum(sb.Principal) as principal" & vbCrLf &
                 "  from(" & vbCrLf &
                 "		SELECT C.Cliente_id," & vbCrLf &
                 "			   C.Endereco_Id," & vbCrLf &
                 "			   C.Nome," & vbCrLf &
                 "			   C.Cidade," & vbCrLf &
                 "			   C.Estado," & vbCrLf &
                 "			   PxD.Principal" & vbCrLf &
                 "		  FROM PedidosXDepositos PxD" & vbCrLf &
                 "		 Inner join Clientes C" & vbCrLf &
                 "			on PxD.Deposito_Id    = c.Cliente_Id" & vbCrLf &
                 "		   and pxd.EndDeposito_Id = c.Endereco_Id" & vbCrLf &
                 "		 where pxd.Empresa_Id    ='" & Parametros("Empresa") & "'" & vbCrLf &
                 "		   and pxd.EndEmpresa_Id = " & Parametros("EndEmpresa") & vbCrLf &
                 "		   and pxd.Pedido_Id     = " & Parametros("Pedido") & vbCrLf &
                 "		   and pxd.Tipo          ='" & Parametros("TipoDeposito") & "'" & vbCrLf &
                 "		 Union" & vbCrLf &
                 "		SELECT Cliente_id," & vbCrLf &
                 "			   Endereco_Id," & vbCrLf &
                 "			   Nome," & vbCrLf &
                 "			   Cidade," & vbCrLf &
                 "			   Estado," & vbCrLf &
                 "			   0 as Principal" & vbCrLf &
                 "		  FROM Clientes" & vbCrLf &
                 "		 where Cliente_Id  ='" & Parametros("Deposito") & "'" & vbCrLf &
                 "		   and Endereco_Id = " & Parametros("EndDeposito") & vbCrLf &
                 "       ) sb" & vbCrLf &
                 "  group by sb.Cliente_id," & vbCrLf &
                 "		   sb.Endereco_Id," & vbCrLf &
                 "		   sb.Nome," & vbCrLf &
                 "		   sb.Cidade," & vbCrLf &
                 "		   sb.Estado" & vbCrLf &
                 "  order by sum(sb.Principal) desc" & vbCrLf

        ds = Banco.ConsultaDataSet(strSQL, "Depositos")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".") & " - " & Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".") & " " & Dr("Estado") & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Dr("Cliente_Id")), 18, ".") & "-" & CStr(Dr("Endereco_Id")), Dr("Cliente_Id") & "-" & CStr(Dr("Endereco_Id"))))
        Next
    End Sub

    Private Sub CarregarClasseToxicologica(ByRef ddl As Object, ByVal Where As String)
        Dim objListClasseToxicologica As New ListClasseToxicologica(True)

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As ClasseToxicologica In objListClasseToxicologica
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr.Descricao, 50, ".") & " - " & Dr.Codigo, Dr.Codigo))
        Next
    End Sub

    Private Sub CarregarClasseAmbiental(ByRef ddl As Object, ByVal Where As String)
        Dim objClasseAmbiental As New ListClasseAmbiental(True, "Descricao")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As ClasseAmbiental In objClasseAmbiental
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr.Descricao, 50, ".") & " - " & Dr.Codigo, Dr.Codigo))
        Next
    End Sub

    Private Sub CarregarEstadoFisicoIA(ByRef ddl As Object, ByVal Where As String)
        Dim objEstadoFisico As New ListEstadoFisicoIA(True, "EstadoFisicoIA_Id")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As EstadoFisicoIA In objEstadoFisico
            ddl.Items.Add(New ListItem(Funcoes.AlinharDireita(Dr.Codigo, 2, "0") & " - " & Dr.Descricao, Dr.Codigo))
        Next
    End Sub

    Private Sub CarregarGeneroDoProduto(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Codigo_Id AS Codigo, Descricao " &
                               "  FROM GeneroDoProduto "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= "ORDER BY Codigo_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "GeneroDoProduto").Tables(0).Rows
            ddl.Items.Add(New ListItem(Funcoes.AlinharDireita(Dr("Codigo"), 2, "0") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub CarregarGeneroDoProdutoXSub(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT SubCodigo_Id AS Codigo, Descricao " &
                               "  FROM GeneroDoProdutoXSub "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= "ORDER BY Codigo_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "GeneroDoProdutoXSub").Tables(0).Rows
            ddl.Items.Add(New ListItem(Funcoes.AlinharDireita(Dr("Codigo"), 2, "0") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub CarregarGruposDeAtivos(ByRef ddl As Object, ByVal Where As String)
        Dim Sql As String = "   SELECT distinct ga.Grupo_Id as Codigo, ga.Grupo_Id + ' - ' + ga.Descricao as Descricao " & vbCrLf &
                            "     FROM GruposDeAtivos ga                                                               " & vbCrLf
        If Not String.IsNullOrWhiteSpace(Where) Then
            Sql &= "    Inner Join Ativos a                                                                   " & vbCrLf &
                            "       on a.Grupo_Id = ga.grupo_Id                                                        " & vbCrLf &
                            IIf(String.IsNullOrWhiteSpace(Where), "", "Where a.Empresa_Id = '" & Left(Where, 8) & "'") & vbCrLf
        End If
        Sql &= "    order by ga.Grupo_Id                                                                  " & vbCrLf

        ddl.Items.Clear()
        ddl.DataValueField = "Codigo"
        ddl.DataTextField = "Descricao"

        ds = Banco.ConsultaDataSet(Sql, "GruposDeAtivos")
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarGrupoQuimico(ByRef ddl As Object, ByVal Where As String)
        Dim objGrupoQuimico As New ListGrupoQuimico(True, "Descricao")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As GrupoQuimico In objGrupoQuimico
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr.Descricao, 50, ".") & " - " & Dr.Codigo, Dr.Codigo))
        Next
    End Sub

    Private Sub CarregarFormulacaoDoFito(ByRef ddl As Object, ByVal Where As String)
        Dim objFormulacaoDoFito As New ListFormulacaoFito(True, "Descricao")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As FormulacaoFito In objFormulacaoDoFito
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr.Descricao, 50, ".") & " - " & Dr.Codigo, Dr.Codigo))
        Next
    End Sub

    Private Sub CarregarFiliaisPedido(ByRef ddl As Object, ByVal Where As String, ByVal Parametros As Hashtable)
        'Parametros tratados
        'Empresa_Id
        'EndEmpresa_Id
        'Cliente_NF
        'EndCliente_NF
        'Pedido

        If Not Parametros.ContainsKey("Empresa_Id") _
        Or Not Parametros.ContainsKey("EndEmpresa_Id") _
        Or Not Parametros.ContainsKey("Cliente_NF") _
        Or Not Parametros.ContainsKey("EndCliente_NF") _
        Or Not Parametros.ContainsKey("Pedido") Then
            Exit Sub
        End If

        Dim strSQL As String
        strSQL = " Select sb.Cliente_Id, sb.EndCliente_id, C.Nome, C.Cidade, C.Estado" & vbCrLf &
                 "   from (" & vbCrLf &
                 "         Select distinct NF.cliente_id, NF.EndCliente_Id, P.cliente, P.EndCliente " & vbCrLf &
                 "           from NotasFiscais NF " & vbCrLf &
                 "          Inner Join Pedidos P" & vbCrLf &
                 "             on P.Empresa_Id    = NF.Empresa_Id" & vbCrLf &
                 "            and P.EndEmpresa_Id = NF.EndEmpresa_Id" & vbCrLf &
                 "            and P.Pedido_Id     = NF.Pedido" & vbCrLf &
                 "          where NF.Situacao = 1" & vbCrLf &
                 "            and NF.Empresa_Id    ='" & Parametros("Empresa_Id") & "'" & vbCrLf &
                 "            and NF.EndEmpresa_Id = " & Parametros("EndEmpresa_Id") & vbCrLf &
                 "            and NF.Pedido        = " & Parametros("Pedido") & vbCrLf

        If Where.Length > 0 Then
            strSQL &= "            and " & Where & vbCrLf
        End If

        strSQL &= "          Union " & vbCrLf &
                  "         Select Cliente, EndCliente, Cliente, EndCliente " & vbCrLf &
                  "           from Pedidos" & vbCrLf &
                  "          where Empresa_Id    ='" & Parametros("Empresa_Id") & "'" & vbCrLf &
                  "            and EndEmpresa_Id = " & Parametros("EndEmpresa_Id") & vbCrLf &
                  "            and Pedido_Id     = " & Parametros("Pedido") & vbCrLf &
                  "          Union" & vbCrLf &
                  "         Select Cliente_Id, Endereco_id, Cliente_Id, Endereco_id" & vbCrLf &
                  "           from Clientes" & vbCrLf &
                  "          where Cliente_Id  ='" & Parametros("Cliente_NF") & "'" & vbCrLf &
                  "            and Endereco_Id = " & Parametros("EndCliente_NF") & vbCrLf &
                  "         ) sb" & vbCrLf &
                  "  inner join Clientes C" & vbCrLf &
                  "     on sb.Cliente_Id    = C.Cliente_Id" & vbCrLf &
                  "    and sb.EndCliente_Id = C.Endereco_Id" & vbCrLf &
                  "  order by case when sb.cliente_id = sb.cliente and sb.endcliente_id = sb.endcliente then 1 else 2 end, " & vbCrLf &
                  "        sb.cliente_Id," & vbCrLf &
                  "        sb.EndCliente_Id  "


        ds = Banco.ConsultaDataSet(strSQL, "FiliaisPedido")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".") & " - " & Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".") & " " & Dr("Estado") & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Dr("Cliente_id")), 18, ".") & "-" & CStr(Dr("EndCliente_Id")), Dr("Cliente_id") & "-" & CStr(Dr("EndCliente_Id"))))
        Next
    End Sub

    Private Sub CarregarHistoricos(ByRef ddl As Object, ByVal Where As String)
        Dim sql As String = "select Historico_Id, Descricao from Historicos order by Descricao"

        ddl.Items.Clear()
        ddl.DataValueField = "Historico_Id"
        ddl.DataTextField = "Descricao"

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Historicos").Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Descricao"), Dr("Historico_Id")))
        Next
    End Sub

    Private Sub CarregarHistoricosDataList(ByRef dataList As Object, ByVal Where As String)
        Dim sql As String = "select Historico_Id, Descricao from Historicos order by Descricao"

        dataList.InnerHtml = String.Empty

        Dim sb As StringBuilder = New StringBuilder()
        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Historicos").Tables(0).Rows
            sb.Append(String.Format("<option value='{0}'>{1}</option>", Dr("Historico_Id"), Dr("Descricao")))
        Next

        dataList.InnerHtml = sb.ToString()
    End Sub

    Private Sub CarregarIngredianteAtivo(ByRef ddl As Object, ByVal Where As String)
        Dim objIngredianteAtivo As New ListIA(True, "NomeComum")

        ddl.Items.Clear()
        ddl.DataValueField = "CodigoIA"
        ddl.DataTextField = "NomeComum"

        For Each Dr As IA In objIngredianteAtivo
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr.NomeComum, 50, ".") & " - " & Dr.CodigoIA, Dr.CodigoIA))
        Next
    End Sub

    Private Sub CarregarIndexador(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = " SELECT Indexador_id, convert(nvarchar,indexador_id) + ' - ' + Descricao as Descricao  " & _
                               "   FROM Indexadores "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= "ORDER BY Descricao"
        ds = Banco.ConsultaDataSet(strSQL, "Indexadores")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Indexador_id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarIndexGtinProduto(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Gtin, Codigo FROM ( " &
                               "SELECT isnull(Gtin8, 0) as G8, isnull(Gtin12, 0) as G2, isnull(Gtin13, 0) as G3, isnull(Gtin14, 0) as G4 " &
                               "FROM Produtos " &
                               "WHERE Produto_id = " & Where &
                               ") AS prd " &
                               "UNPIVOT " &
                               "(Codigo FOR Gtin IN (G8, G2, G3, G4)) AS upv "

        ds = Banco.ConsultaDataSet(strSQL, "Produtos")

        ddl.Items.Clear()

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Gtin"), Dr("Codigo")))
        Next
    End Sub


    Private Sub CarregarClasseDeRisco(ByRef ddl As Object, ByVal Where As String)
        Dim objClasseDeRisco As New ListClasseDeRisco(True, "Descricao")

        ddl.Items.Clear()
        ddl.DataValueField = "Codigo"
        ddl.DataTextField = "Descricao"

        For Each Dr As ClasseDeRisco In objClasseDeRisco
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr.Descricao, 50, ".") & " - " & Dr.Codigo, Dr.Codigo))
        Next
    End Sub

    Private Sub CarregarClasseDeOperacao(ByRef ddl As Object, ByVal Where As String)
        Dim objClasseDeOperacao As New ListClasseDeOperacao(True, Where)

        ddl.Items.Clear()
        ddl.DataValueField = "Codigo"
        ddl.DataTextField = "Descricao"

        ddl.DataSource = objClasseDeOperacao.ToArray
        ddl.DataBind()
    End Sub

    Private Sub CarregarCategoriaCliente(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "Select Categoria_Id, convert(nvarchar,Categoria_Id) + ' - ' + Descricao as Descricao" & _
                               "  from Categorias  "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= "ORDER BY Categoria_Id"
        ds = Banco.ConsultaDataSet(strSQL, "Categoria_Id")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Categoria_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarCarteirasXTributos(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Tributo_Id as Codigo, (Encargos.Descricao + ' - ' + Tributo_Id) as Descricao " & _
                               "  FROM CarteirasXTributos " & _
                               " INNER JOIN Encargos " & _
                               "         ON CarteirasXTributos.Tributo_ID = Encargos.Encargo_id "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= "ORDER BY Tributo_Id"

        ds = Banco.ConsultaDataSet(strSQL, "CarteirasXTributos")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"
        ddl.DataSource = ds
        If ds.Tables(0).Rows.Count > 0 Then
            ddl.DataBind()
        End If

    End Sub

    Private Sub CarregarCultura(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Cultura_Id, Descricao  " & _
                               "  FROM cultura "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= "ORDER BY Descricao"
        ds = Banco.ConsultaDataSet(strSQL, "cultura")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Cultura_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarMarca(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Marca_Id, Descricao  " & _
                               "  FROM Marca "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= "ORDER BY Descricao"
        ds = Banco.ConsultaDataSet(strSQL, "Marca")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Marca_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarMes(ByRef ddl As Object, ByVal Where As String)
        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("JANEIRO", 1))
        ddl.Items.Add(New ListItem("FEVEREIRO", 2))
        ddl.Items.Add(New ListItem("MARCO", 3))
        ddl.Items.Add(New ListItem("ABRIL", 4))
        ddl.Items.Add(New ListItem("MAIO", 5))
        ddl.Items.Add(New ListItem("JUNHO", 6))
        ddl.Items.Add(New ListItem("JULHO", 7))
        ddl.Items.Add(New ListItem("AGOSTO", 8))
        ddl.Items.Add(New ListItem("SETEMBRO", 9))
        ddl.Items.Add(New ListItem("OUTUBRO", 10))
        ddl.Items.Add(New ListItem("NOVEMBRO", 11))
        ddl.Items.Add(New ListItem("DEZEMBRO", 12))
    End Sub

    Private Sub CarregarMoeda(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Moeda_Id, Descricao  " & _
                               "  FROM Moedas "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= "ORDER BY Moeda_Id, Descricao"
        ds = Banco.ConsultaDataSet(strSQL, "Moedas")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Moeda_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarMotivo(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Motivo_Id, Descricao  " & _
                               "  FROM Motivo "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= "ORDER BY Motivo_Id, Descricao"
        ds = Banco.ConsultaDataSet(strSQL, "Moedas")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Motivo_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarModeloCertidaoNegativa(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT ModeloCertidao_Id, Descricao " & _
                               "  FROM  CertidaoNegativaModelo "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        ds = Banco.ConsultaDataSet(strSQL, "CN")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "ModeloCertidao_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarTipoDeEmbalagem(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT TipoDeEmbalagem_Id, TipoDeEmbalagem_Id + ' - ' + Descricao as Descricao  " & _
                               "  FROM TipoDeEmbalagem "


        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY TipoDeEmbalagem_Id, Descricao"

        ds = Banco.ConsultaDataSet(strSQL, "TipoDeEmbalagem")

        ddl.Items.Clear()

        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "TipoDeEmbalagem_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarTabelaDeClassificacoes(ByRef ddl As Object, ByVal Where As String)
        Dim objTabeladeClassificacao As New Classificacoes()

        ddl.Items.Clear()
        ddl.DataValueField = "Codigo"
        ddl.DataTextField = "Descricao"

        If objTabeladeClassificacao.Selecionar() Then
            For Each Dr As Classificacao In objTabeladeClassificacao
                ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr.Codigo, 3, ".") & " - " & Dr.Descricao, Dr.Codigo))
            Next
        End If
    End Sub

    Private Sub CarregarTabelaDePreco(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Codigo_Id AS Codigo, Descricao " & _
                               "  FROM TabelaDePrecos "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= "  ORDER BY Codigo_Id"

        ddl.Items.Clear()

        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "TabelaDePrecos").Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Codigo").ToString("00") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub CarregarTipoDeCliente(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Tipo_Id AS Codigo, Descricao " & _
                               "  FROM TiposDeClientes"

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= "  ORDER BY Tipo_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "PlanoDeContas").Tables(0).Rows
            ddl.Items.Add(New ListItem(Format(Dr("Codigo"), "000") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub CarregarTipoDaContaContabil(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT TipoDeConta_Id AS Codigo " & _
                               "  FROM TipoDeContaContabil"

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= "  ORDER BY TipoDeConta_Id"

        ddl.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "TipoDeContaContabil").Tables(0).Rows
            ddl.Items.Add(New ListItem(Dr("Codigo"), Dr("Codigo")))
        Next
    End Sub

    Private Sub CarregarTipoDeDocumento(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Codigo_Id, convert(nvarchar,Codigo_Id) + ' - ' + Descricao as Descricao  " & _
                               "  FROM TipoDeDocumento "


        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY Codigo_Id"

        ds = Banco.ConsultaDataSet(strSQL, "TipoDeDocumento")

        ddl.Items.Clear()

        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarTipoDeFaturamento(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT TipoDeFaturamento_Id, convert(nvarchar,TipoDeFaturamento_Id) + ' - ' + Descricao as Descricao  " & _
                               "  FROM TipoDeFaturamento "


        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY TipoDeFaturamento_Id"

        ds = Banco.ConsultaDataSet(strSQL, "TipoDeFaturamento")

        ddl.Items.Clear()

        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "TipoDeFaturamento_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarTipoDeCertidao(ByRef ddl As Object, Optional where As String = "")
        Dim Sql As String = "Select Tipo_Id, Descricao From TipoDeCertidao " & IIf(Not String.IsNullOrWhiteSpace(where), "where " & where, "")

        ddl.Items.Clear()
        For Each row As DataRow In Banco.ConsultaDataSet(Sql, "TipoDeCertidao").Tables(0).Rows
            ddl.Items.Add(New ListItem(row("Tipo_Id") & " - " & row("Descricao"), row("Tipo_Id")))
        Next
    End Sub

    Private Sub CarregarTipoDePagamento(ByRef ddl As Object, ByVal Where As String, ByVal Parametros As Hashtable)
        Dim strSQL As String

        strSQL = "SELECT TipoDePagamento_Id as Codigo, convert(varchar,REPLICATE('0', 2 - LEN(CAST(TipoDePagamento_Id AS varchar))) + CAST(TipoDePagamento_Id AS varchar)) + '  -  ' + Descricao as Descricao"

        If Parametros("listarTudo").ToString() = "S" Then strSQL &= ", isnull(EnviaAoBanco,'N') as EnviaAoBanco" & vbCrLf

        strSQL &= "  FROM TiposDePagamentos"

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY TipoDePagamento_Id, Descricao"

        ds = Banco.ConsultaDataSet(strSQL, "TipoDePagamento")

        ddl.Items.Clear()

        ddl.DataValueField = "Codigo"
        ddl.DataTextField = "Descricao"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarTipoDeVeiculo(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String

        strSQL = "SELECT Codigo_Id, convert(nvarchar,Codigo_id) + ' - ' + Descricao as Descricao " & _
                 " FROM TiposDeVeiculos "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY Codigo_Id, Descricao"

        ds = Banco.ConsultaDataSet(strSQL, "TiposDeVeiculos")

        ddl.Items.Clear()

        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarEmbalagem(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Embalagem_Id, ISNULL( " & vbCrLf &
                                                            " REPLICATE('.', LEN(EmbalagemIndea) - LEN(CAST(ISNULL(EmbalagemIndea, '0') AS varchar))) " & vbCrLf &
                                                            " + CAST(ISNULL(EmbalagemIndea, '0') AS varchar) " & vbCrLf &
                                                            " + ' - ' + Descricao, " & vbCrLf &
                                                            " ' - ' + Descricao " & vbCrLf &
                                                            " ) AS Descricao,  " & vbCrLf &
                                                            "case " & vbCrLf &
                                                            "	when Embalagem_Id = 1 " & vbCrLf &
                                                            "		then 0 " & vbCrLf &
                                                            "		else 1 " & vbCrLf &
                                                            "	end as Ordem " & vbCrLf &
                               "FROM Embalagens " & vbCrLf
        If Where.Length > 0 Then
            strSQL &= " WHERE " & Where & vbCrLf
        End If

        'strSQL &= "ORDER BY EmbalagemIndea, Descricao"
        strSQL &= "ORDER BY Ordem"

        ds = Banco.ConsultaDataSet(strSQL, "Embalagens")

        ddl.Items.Clear()

        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Embalagem_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarEmbalagemTipoQtdeDoProduto(ByRef ddl As Object, ByVal Produto As String)
        Dim strSQL As String
        strSQL = "SELECT Convert(nvarchar,ProdutoXEmbalagem.Embalagem_Id)+';'+ProdutoXEmbalagem.TipoDeEmbalagem_Id+';'+ convert(nvarchar,ProdutoXEmbalagem.CapacidadeEmbalagem_Id) as Codigo," & vbCrLf & _
                 "       Embalagens.EmbalagemIndea +' - ' +  Embalagens.Descricao + ' / ' + ProdutoXEmbalagem.TipoDeEmbalagem_Id +' - '+ TipoDeEmbalagem.Descricao +'.....:  '+ convert(nvarchar,ProdutoXEmbalagem.CapacidadeEmbalagem_Id) as Descricao" & vbCrLf & _
                 "  FROM ProdutoXEmbalagem " & vbCrLf & _
                 " INNER JOIN TipoDeEmbalagem " & vbCrLf & _
                 "    ON ProdutoXEmbalagem.TipoDeEmbalagem_Id = TipoDeEmbalagem.TipoDeEmbalagem_Id " & vbCrLf & _
                 " INNER JOIN Embalagens " & vbCrLf & _
                 "    ON ProdutoXEmbalagem.Embalagem_Id = Embalagens.Embalagem_Id" & vbCrLf

        If Produto.Length > 0 Then
            strSQL &= " Where ProdutoXEmbalagem.Produto_Id = '" & Produto & "'"
        End If

        strSQL &= "Order By Embalagens.EmbalagemIndea, ProdutoXEmbalagem.TipoDeEmbalagem_Id, ProdutoXEmbalagem.CapacidadeEmbalagem_Id"

        ds = Banco.ConsultaDataSet(strSQL, "Embalagens")

        ddl.Items.Clear()

        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarEPI(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Codigo_Id, Descricao " & _
                               "FROM EPI "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= "ORDER BY Codigo_Id"

        ds = Banco.ConsultaDataSet(strSQL, "EPI")

        ddl.Items.Clear()

        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarEtapas(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Etapa_Id, Descricao " & _
                               "FROM Etapas "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= "ORDER BY Descricao"

        ds = Banco.ConsultaDataSet(strSQL, "Etapas")

        ddl.Items.Clear()

        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Etapa_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarGrupoProduto(ByRef ddl As Object, ByVal Where As String)
        Dim strGrupoProduto As String
        Dim strSQL As String = "SELECT Grupo_Id, Descricao " & _
                               "  FROM GruposDeEstoques GE " & _
                               " WHERE EXISTS (SELECT NULL " & _
                               "                 FROM Produtos P " & _
                               "                WHERE P.Grupo   = GE.Grupo_Id "

        If Where.Length > 0 Then
            strSQL &= " And P.Agrupar = '" & Where & "' OR GE.Grupo_Id = '10101' OR GE.Grupo_Id = '10102' OR GE.Grupo_Id = '10103' OR GE.Grupo_Id = '10105' OR GE.Grupo_Id = '10201' OR GE.Grupo_Id = '10401' OR GE.Grupo_Id = '10106' OR GE.Grupo_Id = '10401' OR GE.Grupo_Id = '30101' OR GE.Grupo_Id = '10110' OR GE.Grupo_Id = '10111' OR GE.Grupo_Id = '10112' OR GE.Grupo_Id = '19101'"
        End If
        strSQL &= ") "

        strSQL &= "ORDER BY Descricao"

        ds = Banco.ConsultaDataSet(strSQL, "GruposDeEstoques")

        ddl.Items.Clear()

        If Not ds Is Nothing Then

            For Each drGrupoProduto As DataRow In ds.Tables(0).Rows
                strGrupoProduto = Funcoes.AlinharEsquerda(drGrupoProduto("Grupo_Id").ToString(), 12, ".") & " - " & drGrupoProduto("Descricao")
                ddl.Items.Add(New ListItem(strGrupoProduto, drGrupoProduto("Grupo_Id").ToString()))
            Next

        End If

    End Sub

    Private Sub CarregarGrupoProdutoXConsumo(ByRef ddl As Object, ByVal Where As String)
        Dim strGrupoProduto As String
        Dim strSQL As String = "SELECT Grupo_Id, Descricao " & vbCrLf & _
                               "  FROM GruposDeEstoques GE " & vbCrLf & _
                               " WHERE EXISTS (SELECT NULL " & vbCrLf & _
                               "                 FROM Produtos P " & vbCrLf & _
                               "                     INNER JOIN ProdutoXConsumos pXc " & vbCrLf & _
                               "                             ON pXc.Produto_Id = P.Produto_Id" & vbCrLf & _
                               "                WHERE P.Grupo   = GE.Grupo_Id "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= ") "

        strSQL &= "ORDER BY Descricao"

        ds = Banco.ConsultaDataSet(strSQL, "ProdutoXConsumo")

        ddl.Items.Clear()

        For Each drGrupoProduto As DataRow In ds.Tables(0).Rows
            strGrupoProduto = Funcoes.AlinharEsquerda(drGrupoProduto("Grupo_Id").ToString(), 12, ".") & " - " & drGrupoProduto("Descricao")
            ddl.Items.Add(New ListItem(strGrupoProduto, drGrupoProduto("Grupo_Id").ToString()))
        Next
    End Sub

    Private Sub CarregarGrupoProdutoGeral(ByRef ddl As Object, ByVal Where As String)
        Dim strGrupoProduto As String
        Dim strSQL As String = "SELECT Grupo_Id, Descricao " & _
                               "  FROM GruposDeEstoques GE " & _
                               "  WHERE LEN(GE.Grupo_Id) >=5 " & _
                               "  ORDER BY Descricao "

        ds = Banco.ConsultaDataSet(strSQL, "GruposDeEstoques")

        ddl.Items.Clear()

        For Each drGrupoProduto As DataRow In ds.Tables(0).Rows
            strGrupoProduto = Funcoes.AlinharEsquerda(drGrupoProduto("Descricao").ToString(), 45, ".") & " - " & drGrupoProduto("Grupo_Id")
            ddl.Items.Add(New ListItem(strGrupoProduto, drGrupoProduto("Grupo_Id").ToString()))
        Next
    End Sub

    Private Sub CarregarNaturezaJuridica(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT NatJur_Id, Nome" & _
                               "  FROM NaturezaJuridica"

        If Where.Length > 0 Then
            strSQL &= " Where " & Where & ""
        End If

        ds = Banco.ConsultaDataSet(strSQL, "NatJur")

        ddl.Items.Clear()
        ddl.DataTextField = "Nome"
        ddl.DataValueField = "NatJur_Id"

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Funcoes.AlinharDireita(Dr("NatJur_Id"), 6, " ") & " - " & Dr("Nome"), Dr("NatJur_Id")))
        Next
    End Sub

    Private Sub CarregarNivelGrupoProduto(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Grupo_Id As Codigo, Descricao " & _
                               "  FROM GruposDeEstoques GE "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where & ""
        End If

        ds = Banco.ConsultaDataSet(strSQL, "GruposDeEstoques")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        For Each Dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(Funcoes.AlinharDireita(Dr("Codigo"), 6, " ") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub CarregarPais(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String

        strSQL = "SELECT Pais_Id, Descricao " & vbCrLf & _
                 "  FROM Pais "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY case when pais_id = 1058 then 0 else 1 end ,Descricao "

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "Pais")

        For Each drPais As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drPais("Descricao"), drPais("Pais_Id").ToString()))
        Next
    End Sub

    Private Sub CarregarProcedimentoParaProducao(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String

        strSQL = "SELECT Codigo_Id AS Codigo , Descricao " & vbCrLf & _
                 "  FROM ProcedimentoDeProducao "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY Codigo_Id "

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "Processos")

        For Each drPais As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drPais("Descricao"), drPais("Codigo").ToString()))
        Next
    End Sub

    Private Sub CarregarProcessos(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String

        strSQL = "SELECT Processo_Id " & vbCrLf & _
                 "  FROM Processos "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " ORDER BY Processo_Id "

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "Processos")

        For Each drPais As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drPais("Processo_Id").ToString().ToUpper(), drPais("Processo_Id").ToString().ToUpper()))
        Next
    End Sub

    Private Sub CarregarProduto(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        Dim strProdutos As String

        strSQL = "SELECT Produto_Id, " & vbCrLf

        Dim tEmpresas As ClienteXEmpresa = New ClienteXEmpresa(HttpContext.Current.Session("ssEmpresa"), 0)

        If tEmpresas.UsarRegistroMinAgr Then
            strSQL &= "case" & vbCrLf & _
                        "	when len(isnull(RegMinAgr,'')) > 0" & vbCrLf & _
                        "	then Nome + '-' + Descricao + '(' + isnull(RegMinAgr,'')  + ')'" & vbCrLf & _
                        "	else " & vbCrLf & _
                        "		case when Nome = Descricao" & vbCrLf & _
                        "			then Nome" & vbCrLf & _
                        "			else Nome + '-' + Descricao" & vbCrLf & _
                        "		end" & vbCrLf & _
                        "end Nome " & vbCrLf
        ElseIf tEmpresas.UsarDescricaoProduto Then
            strSQL &= "case when Nome = Descricao" & vbCrLf & _
                      "		then Nome" & vbCrLf & _
                      "		else Nome + '-' + Descricao" & vbCrLf & _
                      "end Nome " & vbCrLf
        Else
            strSQL &= "Nome" & vbCrLf
        End If

        strSQL &= "FROM Produtos " & vbCrLf & _
                 "Where Situacao = 1 " & vbCrLf

        If Where.Length > 0 Then
            strSQL &= " AND " & Where
        End If

        strSQL &= " ORDER BY Nome "

        ddl.Items.Clear()

        ds = Banco.ConsultaDataSet(strSQL, "Produtos")

        For Each drProduto As DataRow In ds.Tables(0).Rows
            strProdutos = Funcoes.AlinharEsquerda(drProduto("Produto_Id").ToString(), 12, ".") & " - " & drProduto("Nome")
            ddl.Items.Add(New ListItem(strProdutos, drProduto("Produto_Id").ToString()))
        Next
    End Sub

    Private Sub CarregarProdutoProducao(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        Dim strProdutos As String

        strSQL = "SELECT DISTINCT P.Produto_Id, P.Nome " & vbCrLf & _
                 "FROM Produtos P " & vbCrLf & _
                 "    INNER JOIN ProdutoXConsumos pXc " & vbCrLf & _
                 "            ON pXc.Produto_Id = P.Produto_Id" & vbCrLf & _
                 "Where P.Situacao = 1 " & vbCrLf

        If Where.Length > 0 Then
            strSQL &= " AND " & Where
        End If

        strSQL &= " ORDER BY P.Nome "

        ddl.Items.Clear()

        ds = Banco.ConsultaDataSet(strSQL, "Produtos")

        For Each drProduto As DataRow In ds.Tables(0).Rows
            strProdutos = Funcoes.AlinharEsquerda(drProduto("Produto_Id").ToString(), 12, ".") & " - " & drProduto("Nome")
            ddl.Items.Add(New ListItem(strProdutos, drProduto("Produto_Id").ToString()))
        Next
    End Sub

    Private Sub CarregarProdutoPorGrupo(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        Dim strProdutos As String

        strSQL = "SELECT Pr.Produto_Id, Pr.Nome " & vbCrLf & _
                 "  FROM Produtos Pr" & vbCrLf & _
                 "      Inner Join GruposDeEstoques GE" & vbCrLf & _
                 "          on GE.Grupo_Id = Pr.Grupo" & vbCrLf & _
                 "Where Situacao = 1 " & vbCrLf

        If Where.Length > 0 Then
            strSQL &= " AND Pr.Grupo = '" & Where & "'"
        End If

        strSQL &= " ORDER BY Pr.Nome "

        ddl.Items.Clear()

        ds = Banco.ConsultaDataSet(strSQL, "Produtos")

        For Each drProduto As DataRow In ds.Tables(0).Rows
            strProdutos = Funcoes.AlinharEsquerda(drProduto("Produto_Id").ToString(), 12, ".") & " - " & drProduto("Nome")
            ddl.Items.Add(New ListItem(strProdutos, drProduto("Produto_Id").ToString()))
        Next
    End Sub

    Private Sub CarregarEspecificacaoDoProduto(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String

        strSQL = "SELECT Codigo_Id AS Codigo, Descricao " & vbCrLf & _
                 "  FROM EspecificacaoDoProduto" & vbCrLf

        If Where.Length > 0 Then
            strSQL &= "Where " & Where
        End If

        strSQL &= " ORDER BY Descricao "

        ddl.Items.Clear()

        ds = Banco.ConsultaDataSet(strSQL, "EspecificacaoDoProduto")
        For Each drProduto As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drProduto("Descricao"), drProduto("Codigo")))
        Next

    End Sub


    Private Sub CarregarProvisoes(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT Provisao_Id as Codigo, convert(varchar,REPLICATE('0', 2 - LEN(CAST(Provisao_Id AS varchar))) + CAST(Provisao_Id AS varchar)) + '  -  ' + Descricao as Descricao" & _
                 "  FROM Provisoes"

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        strSQL &= " Order By Provisao_Id"

        ddl.Items.Clear()

        ddl.DataValueField = "Codigo"
        ddl.DataTextField = "Descricao"
        ddl.DataSource = Banco.ConsultaDataSet(strSQL, "Provisoes")
        ddl.DataBind()
    End Sub

    Private Sub CarregarLoteContabil(ByRef ddl As Object, ByVal Where As String)
        Dim sql As String

        sql = "SELECT     Lotes.Lote_Id as Codigo, convert(varchar,Lotes.Lote_Id) + '-' + Lotes.Descricao  as Descricao " & vbCrLf & _
              " FROM     Lotes RIGHT OUTER JOIN " & vbCrLf & _
              " Sistemas ON Lotes.Sistema_Id = Sistemas.Sistema_Id " & vbCrLf & _
              " Where Sistemas.Sistema_Id = 2" & vbCrLf
        If Where.Length > 0 Then
            sql &= " And " & Where
        End If

        ddl.Items.Clear()

        ds = Banco.ConsultaDataSet(sql, "Produtos")
        For Each drProduto As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drProduto("Descricao"), drProduto("Codigo")))
        Next
    End Sub

    Private Sub CarregarLoteProduto(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String

        strSQL = "SELECT Lote_Id " & _
                 "  FROM Lote  "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " ORDER BY Lote_Id"

        ddl.Items.Clear()

        ds = Banco.ConsultaDataSet(strSQL, "Produtos")
        For Each drProduto As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drProduto("Lote_Id"), drProduto("Lote_Id")))
        Next
    End Sub

    Private Sub CarregarLoteClassificacaoProduto(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT Classificacao_Id " & _
                 "FROM LotexClassificacao  "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " ORDER BY Classificacao_Id"

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "Classificacao")

        For Each drProduto As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drProduto("Classificacao_Id"), drProduto("Classificacao_Id")))
        Next
    End Sub

    Private Sub CarregarFinalidade(ByRef ddl As Object, ByVal Where As String)
        Dim objFinalidades As New Finalidades(True)

        For Each objFinalidade As Finalidade In objFinalidades
            ddl.Items.Add(New ListItem(objFinalidade.Codigo.ToString("00") & " - " & objFinalidade.Descricao, objFinalidade.Codigo))
        Next
    End Sub

    Private Sub CarregarSafra(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT Safra_Id " & _
                 "FROM Safras  "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " ORDER BY Safra_Id"

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "Safra")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each drProduto As DataRow In ds.Tables(0).Rows
                ddl.Items.Add(New ListItem(drProduto("Safra_Id"), drProduto("Safra_Id")))
            Next
        End If
    End Sub

    Private Sub CarregarSeguimento(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT Seguimento_Id, Descricao " &
                 "  FROM Seguimentos  "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " ORDER BY Seguimento_Id"

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "Seguimentos")

        For Each drSeguimento As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drSeguimento("Seguimento_Id") & " - " & drSeguimento("Descricao"), drSeguimento("Seguimento_Id")))
        Next
    End Sub

    Private Sub CarregarServidores(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "Select Cliente_id as Servidor, Cliente_id + ' - ' + Nome + ' - ' + Cidade + ' - ' + Estado as DescServidor " & vbCrLf &
                 "  from clientes c" & vbCrLf &
                 " Where exists(select 1" & vbCrLf &
                 "                from clientesxtipos cxt" & vbCrLf &
                 "               where c.cliente_id  = cxt.cliente_id" & vbCrLf &
                 "                 and c.endereco_id = cxt.endereco_id" & vbCrLf &
                 "                 and cxt.tipo_id   = 100)" & vbCrLf

        If Where.Length > 0 Then
            strSQL &= " And " & Where
        End If
        strSQL &= " ORDER BY Cliente_Id"

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "Servidor")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            For Each drProduto As DataRow In ds.Tables(0).Rows
                ddl.Items.Add(New ListItem(drProduto("DescServidor"), drProduto("Servidor")))
            Next
        End If
    End Sub

    Private Sub CarregarSeguimentos(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Seguimento_Id, Descricao  " &
                               "  FROM Seguimentos "
        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= "ORDER BY Descricao"
        ds = Banco.ConsultaDataSet(strSQL, "Seguimentos")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Seguimento_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarSituacao(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT Situacao_Id, Descricao " & _
                 "  FROM Situacoes  "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " ORDER BY Situacao_Id"

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "Situacao")

        For Each drSituacao As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drSituacao("Situacao_Id") & " - " & drSituacao("Descricao"), drSituacao("Situacao_Id")))
        Next
    End Sub

    Private Sub CarregarSituacaoICMS(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT SituacaoTributaria_Id, Descricao " & _
                 "  FROM SituacaoTributaria  "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " ORDER BY SituacaoTributaria_Id"

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "SituacaoTributaria")

        For Each drSituacao As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drSituacao("SituacaoTributaria_Id") & " - " & drSituacao("Descricao"), drSituacao("SituacaoTributaria_Id")))
        Next
    End Sub

    Private Sub CarregarSituacaoIPI(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT SituacaoTributariaIPI_Id, Descricao " & _
                 "  FROM SituacaoTributariaIPI  "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " ORDER BY SituacaoTributariaIPI_Id"

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "SituacaoTributariaIPI")

        For Each drSituacao As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drSituacao("SituacaoTributariaIPI_Id") & " - " & drSituacao("Descricao"), drSituacao("SituacaoTributariaIPI_Id")))
        Next
    End Sub

    Private Sub CarregarSituacaoPISCOFINS(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT SituacaoTributariaPISCOFINS_Id, Descricao " & _
                 "  FROM SituacaoTributariaPISCOFINS  "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " ORDER BY SituacaoTributariaPISCOFINS_Id"

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "SituacaoTributariaPISCOFINS")

        For Each drSituacao As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drSituacao("SituacaoTributariaPISCOFINS_Id") & " - " & drSituacao("Descricao"), drSituacao("SituacaoTributariaPISCOFINS_Id")))
        Next
    End Sub

    Private Sub CarregarSituacaoIBSCBS(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT SituacaoTributariaIBSCBS_Id, Descricao " &
                 "  FROM SituacaoTributariaIBSCBS  "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " ORDER BY SituacaoTributariaIBSCBS_Id"

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "SituacaoTributariaIBSCBS_Id")

        For Each drSituacao As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drSituacao("SituacaoTributariaIBSCBS_Id") & " - " & drSituacao("Descricao"), drSituacao("SituacaoTributariaIBSCBS_Id")))
        Next

    End Sub

    Private Sub CarregarSituacaoPISCOFINSObs(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT Observacao_Id, Descricao " & _
                 "  FROM SituacaoTributariaPISCOFINSObs  "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " ORDER BY Observacao_Id"

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "SituacaoTributariaPISCOFINSObs")

        For Each drSituacao As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drSituacao("Observacao_Id") & " - " & drSituacao("Descricao"), drSituacao("Observacao_Id")))
        Next
    End Sub



    Private Sub CarregarOperacao(ByRef ddl As Object, ByVal Where As String)
        ddl.Items.Clear()

        Dim objOperacoes As New ListOperacao(True)

        For Each objOperacao As Operacao In objOperacoes
            ddl.Items.Add(New ListItem(objOperacao.Codigo.ToString("00") & "-" & objOperacao.Descricao, _
                                                  objOperacao.Codigo))
        Next
    End Sub

    Private Sub CarregaObservacaoTributariaGeral(ByRef ddl As Object, ByVal Where As String)
        ddl.Items.Clear()

        Dim objObs As New ListObservacaoTributariaGeral(Where)

        For Each obj In objObs
            ddl.Items.Add(New ListItem(obj.Codigo.ToString("000") & "-" & obj.Descricao, obj.Codigo))
        Next
    End Sub



    Private Sub CarregarOperacaoSubOperacao(ByRef ddl As Object, ByVal Where As String)
        ddl.Items.Clear()

        Dim objSubOperacoes As New ListSubOperacao(Where)

        For Each objSubOperacao As SubOperacao In objSubOperacoes
            ddl.Items.Add(New ListItem(objSubOperacao.CodigoOperacao.ToString("00") & "-" & objSubOperacao.Codigo.ToString("00") & " " & objSubOperacao.Descricao, _
                                                  objSubOperacao.CodigoOperacao & "-" & objSubOperacao.Codigo))
        Next
    End Sub

    Private Sub CarregarSubOperacaoSemOperacao(ByRef ddl As Object, ByVal Where As String)
        ddl.Items.Clear()

        Dim objSubOperacoes As New ListSubOperacao(Where)

        For Each objSubOperacao As SubOperacao In objSubOperacoes
            ddl.Items.Add(New ListItem(objSubOperacao.Codigo.ToString("00") & " " & objSubOperacao.Descricao, objSubOperacao.Codigo))
        Next
    End Sub

    Private Sub CarregarOperacaoSubOperacaoPermitidasNaNota(ByRef ddl As Object, ByVal Where As String, ByVal Parametros As Hashtable)
        'Parametros lembre-se q sao sensitivos
        'Empresa     string
        'EndEmpresa  int
        'Pedido      int
        'Operacao    int
        'SubOperacao int
        ddl.Items.Clear()
        If Not Parametros.ContainsKey("Operacao") Or Not Parametros.ContainsKey("SubOperacao") Or Not Parametros.ContainsKey("Pedido") Or Not Parametros.ContainsKey("Empresa") Or Not Parametros.ContainsKey("EndEmpresa") Then
            Exit Sub
        End If

        Dim sop As New SubOperacao(CInt(Parametros("Operacao")), CInt(Parametros("SubOperacao")))

        Dim sql As String
        sql = "Select SO.Operacao_id, SO.Suboperacoes_id, SO.Descricao" & vbCrLf &
              "  From Operacoes OP" & vbCrLf &
              " inner Join Suboperacoes SO" & vbCrLf &
              "    on Op.Operacao_id = SO.Operacao_Id" & vbCrLf &
              " Where 1=1 --SO.PrecoFixo='" & IIf(sop.PrecoFixo, "S", "N") & "'" & vbCrLf &
              "   and SO.Situacao= 1 " & vbCrLf &
              "   and isnull(SO.Consignacao,0)= " & IIf(sop.Consignacao, "1", "0") & vbCrLf &
              "   and isnull(SO.Liminar,0)= " & IIf(sop.Liminar, "1", "0") & vbCrLf &
              "   and SO.classe <> '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'" & vbCrLf &
              "   and ISNULL(OP.Classe,'')                = '" & sop.Operacao.CodigoClasse & "'" & vbCrLf

        'Usada na importação do xml para que as operações sejam filtradas pelos cfops presentes na tabela XMLCFOPSaidaXCFOPEntrada'
        If Parametros.ContainsKey("CFOPSaida") Then
            sql &= "  AND Exists(SELECT 1" & vbCrLf & _
                   "               FROM OperacaoXEstado OxE" & vbCrLf & _
                   "               JOIN XmlCFOPSaidaXCFOPEntrada CF" & vbCrLf & _
                   "                 ON OxE.CodigoFiscal = CF.CFOPEntrada_Id" & vbCrLf & _
                   "              WHERE (OxE.Empresa = '99999999' OR OxE.Empresa ='" & Left(Parametros("Empresa"), 8) & "')" & vbCrLf & _
                   "                AND OxE.EstadoOrigem ='" & Parametros("EstadoOrigem") & "'" & vbCrLf & _
                   "                AND (OxE.EstadoDestino ='" & Parametros("EstadoDestino") & "'  OR OxE.EstadoDestino  ='" & Parametros("RegiaoDestino") & "')" & vbCrLf & _
                   "                AND (OxE.GrupoProduto  ='" & Parametros("GrupoDeProduto") & "' OR OxE.Produto ='" & Parametros("Produto") & "')" & vbCrLf & _
                   "                AND CF.CFOPSaida_Id  = " & Parametros("CFOPSaida") & vbCrLf & _
                   "              )" & vbCrLf
        ElseIf Parametros.ContainsKey("EstadoOrigem") Then
            sql &= "  AND Exists(SELECT 1" & vbCrLf & _
                   "               FROM OperacaoXEstado OxE" & vbCrLf & _
                   "              WHERE (OxE.Empresa = '99999999' OR OxE.Empresa ='" & Left(Parametros("Empresa"), 8) & "')" & vbCrLf & _
                   "                AND OxE.EstadoOrigem ='" & Parametros("EstadoOrigem") & "'" & vbCrLf & _
                   "                AND (OxE.EstadoDestino ='" & Parametros("EstadoDestino") & "'  OR OxE.EstadoDestino  ='" & Parametros("RegiaoDestino") & "')" & vbCrLf & _
                   "                AND (OxE.GrupoProduto  ='" & Parametros("GrupoDeProduto") & "' OR OxE.Produto ='" & Parametros("Produto") & "')" & vbCrLf & _
                   "              )" & vbCrLf
        End If
        sql &= " order by Case" & vbCrLf & _
               "            When SO.Operacao_id = " & sop.CodigoOperacao & vbCrLf & _
               "              then 0" & vbCrLf & _
               "              else 1" & vbCrLf & _
               "          end, SO.Operacao_id, So.Suboperacoes_id"

        ds = Banco.ConsultaDataSet(sql, "SO")

        For Each dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(dr("Operacao_id").ToString().PadLeft(2, "0") & "-" & dr("Suboperacoes_id").ToString().PadLeft(2, "0") & " " & dr("Descricao"), dr("Operacao_id") & "-" & dr("Suboperacoes_id")))
        Next
    End Sub

    Private Sub CarregarClientes(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        Dim ds As New DataSet
        Dim Codigo As String
        Dim Cnpj As String
        Dim Nome As String
        Dim Cidade As String
        Dim Descricao As String

        strSQL = "SELECT Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Estado, Clientes.Nome, Clientes.Cidade  " & vbCrLf & _
                 "  FROM Clientes  " & vbCrLf & _
                "   Where Len(Clientes.Nome) > 5 " & vbCrLf

        If Not String.IsNullOrWhiteSpace(Where) Then
            strSQL &= "And " & Where & vbCrLf
        End If
        strSQL &= "   Order By Nome"

        ds = Banco.ConsultaDataSet(strSQL, "Clientes")

        For Each dr As DataRow In ds.Tables(0).Rows
            Codigo = dr("Cliente_Id") & "-" & CStr(dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(dr("Cliente_Id"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & dr("Estado") & " " & Cnpj & "-" & CStr(dr("Endereco_Id"))

            ddl.Items.Add(New ListItem(Descricao, Codigo))
        Next
    End Sub

    Private Sub CarregarClientesTipo(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        Dim ds As New DataSet
        Dim Codigo As String
        Dim Cnpj As String
        Dim Nome As String
        Dim Cidade As String
        Dim Descricao As String

        strSQL = "SELECT Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Estado, Clientes.Nome, Clientes.Cidade  " & vbCrLf & _
                 "  FROM Clientes  " & vbCrLf & _
                 " INNER JOIN ClientesXTipos " & vbCrLf & _
                 "    ON Clientes.Cliente_Id  = ClientesXTipos.Cliente_Id" & vbCrLf & _
                 "   AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id " & vbCrLf & _
                 " WHERE ClientesXTipos.Tipo_Id = " & Where & ""

        ds = Banco.ConsultaDataSet(strSQL, "ClientesXTipos")

        For Each dr As DataRow In ds.Tables(0).Rows
            Codigo = dr("Cliente_Id") & "-" & CStr(dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(dr("Cliente_Id"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & dr("Estado") & " " & Cnpj & "-" & CStr(dr("Endereco_Id"))

            ddl.Items.Add(New ListItem(Descricao, Codigo))
        Next
    End Sub

    Private Sub CarregarCarteiraFinanceira(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        Dim ds As New DataSet

        strSQL = "SELECT  Produto_Id AS Codigo, Produto_Id + ' - ' + Descricao AS Descricao "
        strSQL &= "FROM ComprasXProdutos "
        If Where.Length > 0 Then
            strSQL &= "Where " & Where
        End If
        strSQL &= "Order By Produto_Id"

        ds = Banco.ConsultaDataSet(strSQL, "Carteiras")

        ddl.Items.Clear()

        For Each drCarteira As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drCarteira("Descricao"), drCarteira("Codigo")))
        Next

        ddl.DataValueField = "Codigo"
        ddl.DataTextField = "Descricao"
        ddl.DataSource = Banco.ConsultaDataSet(strSQL, "Carteiras")
        ddl.DataBind()
    End Sub

    Private Sub CarregarCarteiraFinanceiraConta(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        Dim ds As New DataSet

        strSQL = "SELECT Cart.Produto_Id AS Codigo," & vbCrLf &
              "       case" & vbCrLf &
              "			when isnull(Cart.ContaClientes,'') = ''" & vbCrLf &
              "			   then Cart.Produto_Id + '  -  ' + Cart.Descricao" & vbCrLf &
              "			   else Cart.Produto_Id + '  -  ' + Cart.Descricao + ' (' + Cart.ContaClientes + '-' + pl.Titulo + ')'" & vbCrLf &
              "		  end AS Descricao" & vbCrLf &
              "" & vbCrLf &
              "  FROM ComprasXProdutos Cart" & vbCrLf &
              "  LEFT JOIN PlanoDeContas pl " & vbCrLf &
              "         on pl.Conta_Id = Cart.ContaClientes " & vbCrLf

        If Where.Length > 0 Then
            strSQL &= "Where " & Where & vbCrLf
        End If

        strSQL &= "Order By Cart.Produto_Id"

        ds = Banco.ConsultaDataSet(strSQL, "Carteiras")

        ddl.Items.Clear()

        For Each drCarteira As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drCarteira("Descricao"), drCarteira("Codigo")))
        Next

        ddl.DataValueField = "Codigo"
        ddl.DataTextField = "Descricao"
        ddl.DataSource = Banco.ConsultaDataSet(strSQL, "Carteiras")
        ddl.DataBind()
    End Sub

    Private Sub CarregarCarteiraDoTitulo(ByRef ddl As Object, ByVal Where As String)
        Dim sql As String

        sql = "SELECT C.Carteira_Id AS Codigo, C.Descricao, isnull(C.Banco,0) AS Banco, isnull(C.FluxoDeCaixa,0) AS FluxoDeCaixa, isnull(C.EmiteDuplicata,0) AS EmiteDuplicata," & vbCrLf & _
              "       case isnull(C.Banco,0) " & vbCrLf & _
              "         when 0 " & vbCrLf & _
              "           then '' " & vbCrLf & _
              "           else (convert(varchar,C.Banco) + '-' + B.Descricao) " & vbCrLf & _
              "       end AS DescricaoBanco " & vbCrLf & _
              "  FROM Carteira AS C " & vbCrLf & _
              "  LEFT JOIN Bancos AS B" & vbCrLf & _
              "    ON B.Banco_id = C.Banco "

        If Where.Length > 0 Then
            sql &= "Where " & Where
        End If
        sql &= "Order By C.Carteira_Id"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "CarteirasDoTitulo")

        ddl.Items.Clear()
        For Each row As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(String.Format("{0} - {1}", row("Codigo").ToString().PadLeft(3, "0"), row("Descricao")), row("Codigo")))
        Next
    End Sub

    Private Sub CarregarEstados(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Estado_Id, Descricao, Regiao FROM Estados"

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        ds = Banco.ConsultaDataSet(strSQL, "Estados")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Estado_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarEstadoERegiao(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Estado_Id, Descricao, Regiao, Estado_id + ' - ' + Descricao + ' / ' + Regiao as DescReg  FROM Estados"

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " order by case when estado_id in ('SX','DR','EX') then 2 else 1 end, descricao"

        ds = Banco.ConsultaDataSet(strSQL, "Estados")
        ddl.Items.Clear()
        ddl.DataTextField = "DescReg"
        ddl.DataValueField = "Estado_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarEstadosUF(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Estado_Id, Descricao, Regiao FROM Estados"

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        ds = Banco.ConsultaDataSet(strSQL, "Estados")
        ddl.Items.Clear()
        ddl.DataTextField = "Estado_Id"
        ddl.DataValueField = "Estado_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarCidades(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "SELECT Estado_Id, Municipio_Id, Codigo_Id, EstadoIbge FROM Municipios"

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        ds = Banco.ConsultaDataSet(strSQL, "Municipios")
        ddl.Items.Clear()
        ddl.DataTextField = "Municipio_Id"
        ddl.DataValueField = "Codigo_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub



    Private Sub CarregarAjustesDaApuracaoIcms(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT Codigo_Id as Codigo, Descricao " &
                 "FROM AjustesDaApuracaoIcms  "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "AjustesDaApuracaoIcms")

        For Each dr As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(dr("Codigo") & " - " & dr("Descricao"), dr("Codigo")))
        Next
    End Sub

    Private Sub CarregarAno(ByRef ddl As Object, ByVal Where As String)
        ' no where passar 3 parametros separados por ponto e virgula
        ' Primeiro Ano inicial: passar ano ou a letra C para ano corrente
        ' Segundo numero de anos apartir do inicial
        ' Terceiro Se a carga comeca pelo ano I - inicial ou C - corrente data atual 01/11/2012 
        '      ex 1: "2009;5;C"  ordem carregamento opcao corrente  2012,2013,2009,2010,2011
        '      ex 2: "2009;5;I"  ordem carregamento opcao Inicial   2009,2010,2011,2012,2013

        Dim param As String() = Where.Split(";")

        'Primeiro parametro
        Dim AnoInicial As Integer
        If param(0) = "C" Then
            AnoInicial = Date.Now.Year
        Else
            AnoInicial = param(0)
        End If

        'Segundo parametro
        Dim NumeroDeAnos As Integer = param(1)

        ddl.Items.Clear()
        If param(0) = "C" Or param(2) = "I" Then
            For i As Integer = AnoInicial To AnoInicial + NumeroDeAnos
                ddl.Items.Add(New ListItem(i, i))
            Next
        Else
            For i As Integer = AnoInicial To AnoInicial + (Date.Now.Year - AnoInicial - 1)
                ddl.Items.Add(New ListItem(i, i))
            Next
            For i As Integer = Date.Now.Year To Date.Now.Year + (NumeroDeAnos - (Date.Now.Year - AnoInicial))
                ddl.Items.Add(New ListItem(i, i))
            Next
        End If
    End Sub

    Private Sub CarregarBancos(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String = "   SELECT   Banco_Id,                                                                                       " & vbCrLf & _
                                "           Case When LEN(CAST(Banco_Id AS varchar)) < 5 Then                                                     " & vbCrLf & _
                                "               REPLICATE('0', 5 - LEN(CAST(Banco_Id AS varchar))) + CAST(Banco_Id AS varchar) + ' - ' + Descricao    " & vbCrLf & _
                                "           Else CAST(Banco_Id AS varchar) + ' - ' + Descricao                                                    " & vbCrLf & _
                                "           End AS Descricao, Ativo                                                                                   " & vbCrLf & _
                                "   FROM Bancos order by Banco_Id                                                                         " & vbCrLf

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        ds = Banco.ConsultaDataSet(strSQL, "Bancos")
        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Banco_Id"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarBancosXContaContabil(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "Select BC.ContaContabil, BC.Titulo" & vbCrLf & _
                 "  from (" & vbCrLf & _
                 "		  Select distinct BC.Empresa_Id, BC.EndEmpresa_Id, BC.ContaContabil, BC.ContaContabil + ' - ' + PC.Titulo as Titulo" & vbCrLf & _
                 "		    from BancosXContas BC" & vbCrLf & _
                 "		   inner join PlanodeContas PC" & vbCrLf & _
                 "		  	  on BC.ContaContabil = PC.Conta_Id" & vbCrLf & _
                 "		   union" & vbCrLf & _
                 "		  Select BC.Empresa_Id, BC.EndEmpresa_Id, BC.ContaGrupoBanco, BC.ContaGrupoBanco + ' - ' + PC.Titulo as Titulo" & vbCrLf & _
                 "		    from clientesXEmpresas BC" & vbCrLf & _
                 "		   inner join PlanodeContas PC" & vbCrLf & _
                 "		 	  on BC.ContaGrupoBanco = PC.Conta_Id" & vbCrLf & _
                 "        ) BC" & vbCrLf

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If

        ds = Banco.ConsultaDataSet(strSQL, "BancosXContas")
        ddl.Items.Clear()
        ddl.DataTextField = "Titulo"
        ddl.DataValueField = "ContaContabil"
        ddl.DataSource = ds
        ddl.DataBind()
    End Sub

    Private Sub CarregarUnidadeDeMedida(ByRef ddl As Object, ByVal Where As String)
        ddl.Items.Clear()

        Dim objUnidadeDeMedida As New ListUnidadeDeMedida()

        For Each row As UnidadeDeMedida In objUnidadeDeMedida
            ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(row.Unidade, 3, ".") & " - " & row.Descricao, row.Unidade))
        Next
    End Sub

    Private Sub CarregarUsuarios(ByRef ddl As Object, ByVal Where As String)
        Dim strSQL As String
        strSQL = "SELECT Usuario_id, NomeCompleto " & _
                 "FROM Usuarios  "

        If Where.Length > 0 Then
            strSQL &= " Where " & Where
        End If
        strSQL &= " ORDER BY Usuario_id"

        ddl.Items.Clear()
        ds = Banco.ConsultaDataSet(strSQL, "Usuarios")

        For Each drProduto As DataRow In ds.Tables(0).Rows
            ddl.Items.Add(New ListItem(drProduto("NomeCompleto"), drProduto("Usuario_id")))
        Next
    End Sub
#End Region

End Class