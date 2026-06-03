Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis
Imports System.Web

<Serializable()> _
Public Class ClienteXEmpresa
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal Empresa As String, ByVal EndEmpresa As Integer)
        Dim Sql As String
        Dim db As New AcessaBanco
        Dim ds As DataSet

        Sql = " SELECT Empresa_Id, EndEmpresa_Id, isnull(Servidor,'') as Servidor, AtividadeEconomica, AtividadeEstadual, DatadeFundacao, RegistroNaJunta, EstadodaJunta, NomeDoContador, CPFContador, CNPJContador, " & vbCrLf &
              "        CRCContador, NomeDoTitular, CPFTitular, CondicaoTitular, Matriz, InscricaoMunicipal, RegistroEp, RegistroEi, RegistroEc, isnull(RegistroNire,'') as RegistroNire, isnull(DataRegistroNire,'') as DataRegistroNire, " & vbCrLf &
              "        isnull(QualificacaoContador,'') as QualificacaoContador, isnull(CodigoQualificacaoContador,900) as CodigoQualificacaoContador, isnull(QualificacaoTitular,'') as QualificacaoTitular, isnull(CodigoQualificacaoTitular,203) as CodigoQualificacaoTitular," & vbCrLf &
              "        isnull(CodigoDeRelacionamento,10) as CodigoDeRelacionamento, isnull(EscrituracaoContabil,'') as EscrituracaoContabil, isnull(Marca,'') as Marca, isnull(NotaFiscalEletronica,'N') as NotaFiscalEletronica, isnull(CertidaoNegativa,0) as CertidaoNegativa, isnull(FluxoDeCaixa,'FALSE') as FluxoDeCaixa, isnull(SugereDeposito,0) as SugereDeposito, isnull(NossaEmissao,0) as NossaEmissao, " & vbCrLf &
              "        isnull(ContaVariacaoMonetariaPassiva,'') as ContaVariacaoMonetariaPassiva, isnull(ContaVariacaoMonetariaAtiva,'') as ContaVariacaoMonetariaAtiva, isnull(ContaVariacaoMonetariaCliente,'') as ContaVariacaoMonetariaCliente, isnull(ContaVariacaoMonetariaFornecedor,'') as ContaVariacaoMonetariaFornecedor, isnull(ContaGrupoBanco,'') as ContaGrupoBanco, isnull(RegistroDeExportacao,1) as RegistroDeExportacao, " & vbCrLf &
              "        isnull(FretePedido,0) as FretePedido, isnull(PedidoBloqueado,0) as PedidoBloqueado, isnull(LiberaCarregamento,0) as LiberaCarregamento, " & vbCrLf &
              "        isnull(ContaJuroPago,'') as ContaJuroPago, isnull(ContaJuroRecebido,'') as ContaJuroRecebido,  isnull(ContaGrupoDuplicatasDescontada,'') as ContaGrupoDuplicatasDescontada, " & vbCrLf &
              "        isnull(PrazoDeCancelamentoNFE,0) as PrazoDeCancelamentoNFE, isnull(ContaGrupoComissoes,'') as ContaGrupoComissoes, isnull(ContaAdiantamentoDeFrete, '') as ContaAdiantamentoDeFrete, isnull(ContaPedagioDeFrete, '') as ContaPedagioDeFrete, " & vbCrLf &
              "        isnull(CodigoProdutoDeFrete, '') as CodigoProdutoDeFrete, isnull(CodigoProdutoDeEstadia, '') as CodigoProdutoDeEstadia, isnull(CodigoProdutoDeMDFe, '') as CodigoProdutoDeMDFe, isnull(CodigoUnidadeDeNegocio, '') as CodigoUnidadeDeNegocio, isnull(ViasNFE, 1) as ViasNFE, " & vbCrLf &
              "        isnull(ControlaEmissaoCheque, 0) as ControlaEmissaoCheque, isnull(ControlaDataMovimentoNFG, 0) as ControlaDataMovimentoNFG, isnull(ObservacaoSefazNFE,'') AS ObservacaoSefazNFE, isnull(Crt,3) AS Crt, ISNULL(ArquivoNFE, 0) AS ArquivoNFE, ISNULL(ConferenciaNFE, 0) AS ConferenciaNFE, PathDownloadNFe, " & vbCrLf &
              "        ISNULL(ContaGrupoTEDDOC,'') AS ContaGrupoTEDDOC, ISNULL(ContaDescontoConcedido,'') AS ContaDescontoConcedido, ISNULL(ContaDescontoObtido,'') AS ContaDescontoObtido,  ISNULL(ContaFornecedorFrete,'') AS ContaFornecedorFrete, ISNULL(ContaCaixaCompensacao,'') AS ContaCaixaCompensacao,  " & vbCrLf &
              "        ISNULL(DiasNFRetroativa,0) DiasNFRetroativa, isnull(ObrigaEncargo,0) as ObrigaEncargo, isnull(Tolerancia,0) as Tolerancia," & vbCrLf &
              "        ISNULL(EstadoExpCRC, '') EstadoExpCRC, isnull(TelefoneContador,'') as TelefoneContador, isnull(EmailContador,'') as EmailContador," & vbCrLf &
              "        isnull(ContadeCustoInicial,'') as ContadeCustoInicial, isnull(ContadeCustoFinal,'') as ContadeCustoFinal, isnull(ContaGrupoEstoque,'') as ContaGrupoEstoque," & vbCrLf &
              "        isnull(ContaEstoqueEmNossoPoder,'') as ContaEstoqueEmNossoPoder, isnull(ContaEstoqueEmPoderDeTerceiros,'') as ContaEstoqueEmPoderDeTerceiros," & vbCrLf &
              "        isnull(ContaPatrimonio,'') as ContaPatrimonio, isnull(Cnae,'') as Cnae, isnull(NatJur,0) as NatJur," & vbCrLf &
              "        isnull(DataFinanceiro,'2001-01-01') as DataFinanceiro, isnull(BaixaFinanceiroPorLote,0) as BaixaFinanceiroPorLote, " & vbCrLf &
              "        isnull(ObrigaNfProdutor,0) as ObrigaNfProdutor, isnull(ObrigaChaveNf,0) as ObrigaChaveNf, isnull(ObrigaChaveNfg,0) as ObrigaChaveNfg, " & vbCrLf &
              "        isnull(UsarDescricaoProduto,0) as UsarDescricaoProduto, isnull(UsarRegistroMinAgr,0) as UsarRegistroMinAgr, isnull(ObrigaNavio,0) as ObrigaNavio, " & vbCrLf &
              "        isnull(UsuarioInclusao,'') as UsuarioInclusao," & vbCrLf &
              "        isnull(UsuarioInclusaoData,CURRENT_TIMESTAMP) as UsuarioInclusaoData," & vbCrLf &
              "        isnull(UsuarioAlteracao,'') as UsuarioAlteracao," & vbCrLf &
              "        isnull(UsuarioAlteracaoData,CURRENT_TIMESTAMP) as UsuarioAlteracaoData," & vbCrLf &
              "        isnull(EmitirCTe,0) as EmitirCTe " & vbCrLf &
              "   FROM ClientesXEmpresas" & vbCrLf &
              "  WHERE Empresa_Id    ='" & Empresa & "'" & vbCrLf &
              "    AND EndEmpresa_Id = " & EndEmpresa

        ds = db.ConsultaDataSet(Sql, "Empresa")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Empresa_id = row("Empresa_Id")
                EndEmpresa_Id = row("EndEmpresa_Id")
                CodigoUnidadeNegocio = row("CodigoUnidadeDeNegocio")
                Servidor = row("Servidor")
                AtividadeEconomica = row("AtividadeEconomica")
                AtividadeEstadual = row("AtividadeEstadual")
                DataDeFundacao = row("DataDeFundacao")
                RegistroNaJunta = row("RegistroNaJunta")
                EstadoDaJunta = row("EstadoDaJunta")
                NomeDoContador = row("NomeDoContador")
                CPFContador = row("CPFContador")
                CNPJContador = row("CNPJContador")
                CRCContador = row("CRCContador")
                EstadoExpCRC = row("EstadoExpCRC")
                TelefoneContador = row("TelefoneContador")
                EmailContador = row("EmailContador")
                NomeDoTitular = row("NomeDoTitular")
                CPFTitular = row("CPFTitular")
                CondicaoTitular = row("CondicaoTitular")
                Matriz = row("Matriz")
                InscricaoMunicipal = row("InscricaoMunicipal")
                RegistroEp = row("RegistroEp")
                RegistroEi = row("RegistroEi")
                RegistroEc = row("RegistroEc")
                RegistroNire = row("RegistroNire")
                DataRegistroNire = row("DataRegistroNire")
                QualificacaoContador = row("QualificacaoContador")
                CodigoQualificacaoContador = row("CodigoQualificacaoContador")
                QualificacaoTitular = row("QualificacaoTitular")
                CodigoQualificacaoTitular = row("CodigoQualificacaoTitular")
                CodigoDeRelacionamento = row("CodigoDeRelacionamento")
                EscrituracaoContabil = row("EscrituracaoContabil")
                Marca = row("Marca")
                NotaFiscalEletronica = row("NotaFiscalEletronica").Equals("S")
                CertidaoNegativa = row("CertidaoNegativa")
                FluxoDeCaixa = row("FluxoDeCaixa").Equals("TRUE")
                SugereDeposito = row("SugereDeposito")
                NossaEmissao = row("NossaEmissao")

                CodigoContadeCustoInicial = row("ContadeCustoInicial")
                CodigoContadeCustoFinal = row("ContadeCustoFinal")

                CodigoContaPatrimonio = row("ContaPatrimonio")

                CodigoContaGrupoEstoque = row("ContaGrupoEstoque")

                CodigoContaEstoqueEmNossoPoder = row("ContaEstoqueEmNossoPoder")
                CodigoContaEstoqueEmPoderDeTerceiros = row("ContaEstoqueEmPoderDeTerceiros")

                CodigoContaVariacaoMonetariaAtiva = row("ContaVariacaoMonetariaAtiva")
                CodigoContaVariacaoMonetariaPassiva = row("ContaVariacaoMonetariaPassiva")

                CodigoContaVariacaoMonetariaCliente = row("ContaVariacaoMonetariaCliente")
                CodigoContaVariacaoMonetariaFornecedor = row("ContaVariacaoMonetariaFornecedor")

                CodigoContaGrupoBanco = row("ContaGrupoBanco")
                CodigoContaJuroPago = row("ContaJuroPago")
                CodigoContaJuroRecebido = row("ContaJuroRecebido")

                RegistroDeExportacao = row("RegistroDeExportacao")
                FretePedido = row("FretePedido")
                PedidoBloqueado = row("PedidoBloqueado")
                LiberaCarregamento = row("LiberaCarregamento")
                CodigoContaJuroPago = row("ContaJuroPago")
                CodigoContaJuroRecebido = row("ContaJuroRecebido")
                CodigoContaGrupoDuplicatasDescontada = row("ContaGrupoDuplicatasDescontada")
                PrazoCancelamentoNFE = row("PrazoDeCancelamentoNFE")
                CodigoContaGrupoComissoes = row("ContaGrupoComissoes")
                CodigoContaAdiantamentoDeFrete = row("ContaAdiantamentoDeFrete")
                CodigoContaPedagioDeFrete = row("ContaPedagioDeFrete")
                CodigoProdutoDeFrete = row("CodigoProdutoDeFrete")
                CodigoProdutoDeEstadia = row("CodigoProdutoDeEstadia")
                CodigoProdutoDeMDFe = row("CodigoProdutoDeMDFe")
                ViasNFE = row("ViasNFE")
                ControlaEmissaoCheque = row("ControlaEmissaoCheque")
                ControlaDataMovimentoNFG = row("ControlaDataMovimentoNFG")
                ObservacaoSefazNFE = row("ObservacaoSefazNFE")
                Crt = row("Crt")
                ArquivoNFE = row("ArquivoNFE")
                ConferenciaNFE = row("ConferenciaNFE")
                PathDownloadNFe = row("PathDownloadNFe").ToString
                CodigoContaGrupoTedDoc = row("ContaGrupoTEDDOC")
                CodigoContaDescontoConcedido = row("ContaDescontoConcedido")
                CodigoContaDescontoObtido = row("ContaDescontoObtido")
                CodigoContaFornecedorFrete = row("ContaFornecedorFrete")
                CodigoContaCaixaCompensacao = row("ContaCaixaCompensacao")
                DiasNFRetroativa = row("DiasNFRetroativa")
                ApenasExcedenteTolerancia = row("Tolerancia")
                ObrigaEncargo = row("ObrigaEncargo")
                DataFinanceiro = row("DataFinanceiro")
                CodigoCnae = row("Cnae")
                CodigoNaturezaJuridica = row("NatJur")
                BaixaFinanceiroPorLote = row("BaixaFinanceiroPorLote")
                ObrigaNfProdutor = row("ObrigaNfProdutor")
                ObrigaChaveNf = row("ObrigaChaveNf")
                ObrigaChaveNfg = row("ObrigaChaveNfg")
                UsarDescricaoProduto = row("UsarDescricaoProduto")
                UsarRegistroMinAgr = row("UsarRegistroMinAgr")
                ObrigaNavio = row("ObrigaNavio")

                UsuarioInclusao = row("UsuarioInclusao")
                DataInclusao = row("UsuarioInclusaoData")
                UsuarioAlteracao = row("UsuarioAlteracao")
                DataAlteracao = row("UsuarioAlteracaoData")

                EmitirCTe = row("EmitirCTe")

            Next
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Empresa_id As String
    Private _EndEmpresa_Id As Integer
    Private _Servidor As String = ""
    Private _AtividadeEconomica As String
    Private _AtividadeEstadual As String
    Private _DataDeFundacao As Date
    Private _RegistroNaJunta As String
    Private _EstadoDaJunta As String

    Private _NomeDoContador As String
    Private _CPFContador As String
    Private _CNPJContador As String
    Private _CRCContador As String
    Private _EstadoExpCRC As String
    Private _TelefoneContador As String
    Private _EmailContador As String

    Private _NomeDoTitular As String
    Private _CPFTitular As String
    Private _CondicaoTitular As String
    Private _Matriz As String
    Private _InscricaoMunicipal As String
    Private _RegistroEp As String
    Private _RegistroEi As String
    Private _RegistroEc As String
    Private _RegistroNire As String
    Private _DataRegistroNire As Date
    Private _QualificacaoContador As String
    Private _CodigoQualificacaoContador As Integer
    Private _QualificacaoTitular As String
    Private _CodigoQualificacaoTitular As Integer
    Private _CodigoDeRelacionamento As Integer
    Private _EscrituracaoContabil As String
    Private _Marca As String
    Private _NotaFiscalEletronica As Boolean = False
    Private _EmitirCTe As Boolean = False
    Private _CertidaoNegativa As Boolean = False
    Private _FluxoDeCaixa As Boolean = False
    Private _SugereDeposito As Boolean = False
    Private _NossaEmissao As Boolean = False
    Private _RegistroDeExportacao As Boolean = False
    Private _FretePedido As Boolean = False
    Private _PedidoBloqueado As Boolean = False
    Private _LiberaCarregamento As Boolean = False

    Private _CodigoContadeCustoInicial As String = ""
    Private _ContadeCustoInicial As PlanoDeConta

    Private _CodigoContadeCustoFinal As String = ""
    Private _ContadeCustoFinal As PlanoDeConta

    Private _CodigoContaPatrimonio As String = ""
    Private _ContaPatrimonio As PlanoDeConta

    Private _CodigoContaGrupoEstoque As String = ""
    Private _ContaGrupoEstoque As PlanoDeConta

    Private _CodigoContaEstoqueEmNossoPoder As String = ""
    Private _ContaEstoqueEmNossoPoder As PlanoDeConta

    Private _CodigoContaEstoqueEmPoderDeTerceiros As String = ""
    Private _ContaEstoqueEmPoderDeTerceiros As PlanoDeConta

    Private _CodigoContaVariacaoMonetariaPassiva As String = ""
    Private _ContaVariacaoMonetariaPassiva As PlanoDeConta

    Private _CodigoContaVariacaoMonetariaAtiva As String = ""
    Private _ContaVariacaoMonetariaAtiva As PlanoDeConta

    Private _CodigoContaVariacaoMonetariaCliente As String = ""
    Private _ContaVariacaoMonetariaCliente As PlanoDeConta

    Private _CodigoContaVariacaoMonetariaFornecedor As String = ""
    Private _ContaVariacaoMonetariaFornecedor As PlanoDeConta

    Private _CodigoContaGrupoBanco As String = ""
    Private _ContaGrupoBanco As PlanoDeConta

    Private _CodigoContaJuroPago As String = ""
    Private _ContaJuroPago As PlanoDeConta

    Private _CodigoContaJuroRecebido As String = ""
    Private _ContaJuroRecebido As PlanoDeConta

    Private _CodigoContaGrupoDuplicatasDescontada As String = ""
    Private _ContaGrupoDuplicatasDescontada As PlanoDeConta

    Private _CodigoContaGrupoComissoes As String = ""
    Private _ContaGrupoComissoes As PlanoDeConta

    Private _CodigoContaAdiantamentoDeFrete As String = ""
    Private _ContaAdiantamentoDeFrete As PlanoDeConta

    Private _CodigoContaPedagioDeFrete As String = ""
    Private _ContaPedagioDeFrete As PlanoDeConta

    Private _CodigoContaGrupoTedDoc As String = ""
    Private _ContaGrupoTedDoc As PlanoDeConta

    Private _CodigoContaFornecedorFrete As String = ""
    Private _ContaFornecedorFrete As PlanoDeConta

    Private _CodigoContaCaixaCompensacao As String = ""
    Private _ContaCaixaCompensacao As PlanoDeConta


    Private _CodigoContaDescontoObtido As String = ""
    Private _ContaDescontoObtido As PlanoDeConta

    Private _CodigoContaDescontoConcedido As String = ""
    Private _ContaDescontoConcedido As PlanoDeConta

    Private _CodigoProdutoDeFrete As String
    Private _CodigoProdutoDeEstadia As String
    Private _CodigoProdutoDeMDFe As String
    Private _CodigoUnidadeNegocio As String

    Private _ViasNFE As Integer
    Private _PrazoConcelamentoNFE As Integer

    Private _ControlaEmissaoCheque As Boolean = False
    Private _ControlaDataMovimentoNFG As Boolean = False

    Private _ObservacaoSefazNFE As String = ""
    Private _ArquivoNFE As Boolean = False
    Private _ConferenciaNFE As Boolean = False

    Private _Crt As Integer
    Private _PathDownloadNFe As String

    Private _DiasNFRetroativa As Integer = 0

    Private _ObrigaEncargo As Boolean
    Private _ApenasExcedenteTolerancia As Boolean = False

    Private _DataFinanceiro As Date

    Private _CodigoCnae As String
    Private _CodigoNaturezaJuridica As Integer

    Private _BaixaFinanceiroPorLote As Boolean = False

    Private _ObrigaNfProdutor As Boolean = False

    Private _ObrigaChaveNf As Boolean = False
    Private _ObrigaChaveNfg As Boolean = False

    Private _UsarDescricaoProduto As Boolean = False
    Private _UsarRegistroMinAgr As Boolean = False

    Private _ObrigaNavio As Boolean = False

    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As DateTime

    Private _ParametrosECF As ClienteXEmpresaXECF
#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property

    Public Property Empresa_id() As String
        Get
            Return _Empresa_id
        End Get
        Set(ByVal value As String)
            _Empresa_id = value
        End Set
    End Property

    Public Property EndEmpresa_Id() As Integer
        Get
            Return _EndEmpresa_Id
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa_Id = value
        End Set
    End Property

    Public Property Servidor As String
        Get
            Return _Servidor
        End Get
        Set(value As String)
            _Servidor = value
        End Set
    End Property

    Public Property AtividadeEconomica() As String
        Get
            Return _AtividadeEconomica
        End Get
        Set(ByVal value As String)
            _AtividadeEconomica = value
        End Set
    End Property

    Public Property AtividadeEstadual As String
        Get
            Return _AtividadeEstadual
        End Get
        Set(value As String)
            _AtividadeEstadual = value
        End Set
    End Property

    Public Property DataDeFundacao() As Date
        Get
            Return _DataDeFundacao
        End Get
        Set(ByVal value As Date)
            _DataDeFundacao = value
        End Set
    End Property

    Public Property RegistroNaJunta() As String
        Get
            Return _RegistroNaJunta
        End Get
        Set(ByVal value As String)
            _RegistroNaJunta = value
        End Set
    End Property

    Public Property EstadoDaJunta() As String
        Get
            Return _EstadoDaJunta
        End Get
        Set(ByVal value As String)
            _EstadoDaJunta = value
        End Set
    End Property

    Public Property NomeDoContador() As String
        Get
            Return _NomeDoContador
        End Get
        Set(ByVal value As String)
            _NomeDoContador = value
        End Set
    End Property

    Public Property CPFContador() As String
        Get
            Return _CPFContador
        End Get
        Set(ByVal value As String)
            _CPFContador = value
        End Set
    End Property

    Public Property CNPJContador() As String
        Get
            Return _CNPJContador
        End Get
        Set(ByVal value As String)
            _CNPJContador = value
        End Set
    End Property

    Public Property CRCContador() As String
        Get
            Return _CRCContador
        End Get
        Set(ByVal value As String)
            _CRCContador = value
        End Set
    End Property

    Public Property EstadoExpCRC As String
        Get
            Return _EstadoExpCRC
        End Get
        Set(value As String)
            _EstadoExpCRC = value
        End Set
    End Property
    Public Property TelefoneContador As String
        Get
            Return _TelefoneContador
        End Get
        Set(value As String)
            _TelefoneContador = value
        End Set
    End Property
    Public Property EmailContador As String
        Get
            Return _EmailContador
        End Get
        Set(value As String)
            _EmailContador = value
        End Set
    End Property

    Public Property NomeDoTitular() As String
        Get
            Return _NomeDoTitular
        End Get
        Set(ByVal value As String)
            _NomeDoTitular = value
        End Set
    End Property

    Public Property CPFTitular() As String
        Get
            Return _CPFTitular
        End Get
        Set(ByVal value As String)
            _CPFTitular = value
        End Set
    End Property

    Public Property CondicaoTitular() As String
        Get
            Return _CondicaoTitular
        End Get
        Set(ByVal value As String)
            _CondicaoTitular = value
        End Set
    End Property

    Public Property Matriz() As String
        Get
            Return _Matriz
        End Get
        Set(ByVal value As String)
            _Matriz = value
        End Set
    End Property

    Public Property InscricaoMunicipal() As String
        Get
            Return _InscricaoMunicipal
        End Get
        Set(ByVal value As String)
            _InscricaoMunicipal = value
        End Set
    End Property

    Public Property RegistroEp() As String
        Get
            Return _RegistroEp
        End Get
        Set(ByVal value As String)
            _RegistroEp = value
        End Set
    End Property

    Public Property RegistroEi() As String
        Get
            Return _RegistroEi
        End Get
        Set(ByVal value As String)
            _RegistroEi = value
        End Set
    End Property

    Public Property RegistroEc() As String
        Get
            Return _RegistroEc
        End Get
        Set(ByVal value As String)
            _RegistroEc = value
        End Set
    End Property

    Public Property RegistroNire As String
        Get
            Return _RegistroNire
        End Get
        Set(value As String)
            _RegistroNire = value
        End Set
    End Property

    Public Property DataRegistroNire() As Date
        Get
            Return _DataRegistroNire
        End Get
        Set(ByVal value As Date)
            _DataRegistroNire = value
        End Set
    End Property

    Public Property QualificacaoContador() As String
        Get
            Return _QualificacaoContador
        End Get
        Set(ByVal value As String)
            _QualificacaoContador = value
        End Set
    End Property

    Public Property CodigoQualificacaoContador As Integer
        Get
            Return _CodigoQualificacaoContador
        End Get
        Set(value As Integer)
            _CodigoQualificacaoContador = value
        End Set
    End Property

    Public Property QualificacaoTitular() As String
        Get
            Return _QualificacaoTitular
        End Get
        Set(ByVal value As String)
            _QualificacaoTitular = value
        End Set
    End Property

    Public Property CodigoQualificacaoTitular() As Integer
        Get
            Return _CodigoQualificacaoTitular
        End Get
        Set(ByVal value As Integer)
            _CodigoQualificacaoTitular = value
        End Set
    End Property

    Public Property CodigoDeRelacionamento() As Integer
        Get
            Return _CodigoDeRelacionamento
        End Get
        Set(ByVal value As Integer)
            _CodigoDeRelacionamento = value
        End Set
    End Property

    Public Property EscrituracaoContabil() As String
        Get
            Return _EscrituracaoContabil
        End Get
        Set(ByVal value As String)
            _EscrituracaoContabil = value
        End Set
    End Property

    Public Property Marca() As String
        Get
            Return _Marca
        End Get
        Set(ByVal value As String)
            _Marca = value
        End Set
    End Property

    Public Property NotaFiscalEletronica() As Boolean
        Get
            Return _NotaFiscalEletronica
        End Get
        Set(ByVal value As Boolean)
            _NotaFiscalEletronica = value
        End Set
    End Property

    Public Property CertidaoNegativa() As Boolean
        Get
            Return _CertidaoNegativa
        End Get
        Set(ByVal value As Boolean)
            _CertidaoNegativa = value
        End Set
    End Property

    Public Property FluxoDeCaixa() As Boolean
        Get
            Return _FluxoDeCaixa
        End Get
        Set(ByVal value As Boolean)
            _FluxoDeCaixa = value
        End Set
    End Property

    Public Property SugereDeposito() As Boolean
        Get
            Return _SugereDeposito
        End Get
        Set(ByVal value As Boolean)
            _SugereDeposito = value
        End Set
    End Property

    Public Property NossaEmissao() As Boolean
        Get
            Return _NossaEmissao
        End Get
        Set(ByVal value As Boolean)
            _NossaEmissao = value
        End Set
    End Property

    Public Property RegistroDeExportacao() As Boolean
        Get
            Return _RegistroDeExportacao
        End Get
        Set(ByVal value As Boolean)
            _RegistroDeExportacao = value
        End Set
    End Property

    Public Property FretePedido() As Boolean
        Get
            Return _FretePedido
        End Get
        Set(ByVal value As Boolean)
            _FretePedido = value
        End Set
    End Property

    Public Property PedidoBloqueado() As Boolean
        Get
            Return _PedidoBloqueado
        End Get
        Set(ByVal value As Boolean)
            _PedidoBloqueado = value
        End Set
    End Property

    Public Property LiberaCarregamento() As Boolean
        Get
            Return _LiberaCarregamento
        End Get
        Set(ByVal value As Boolean)
            _LiberaCarregamento = value
        End Set
    End Property

    Public Property CodigoContadeCustoInicial As String
        Get
            Return _CodigoContadeCustoInicial
        End Get
        Set(value As String)
            _CodigoContadeCustoInicial = value
            _ContadeCustoInicial = Nothing
        End Set
    End Property
    Public ReadOnly Property ContadeCustoInicial As PlanoDeConta
        Get
            If _ContadeCustoInicial Is Nothing And _CodigoContadeCustoInicial.Length > 0 Then _ContadeCustoInicial = New PlanoDeConta("", 0, _CodigoContadeCustoInicial)
            Return _ContadeCustoInicial
        End Get
    End Property

    Public Property CodigoContadeCustoFinal As String
        Get
            Return _CodigoContadeCustoFinal
        End Get
        Set(value As String)
            _CodigoContadeCustoFinal = value
            _ContadeCustoFinal = Nothing
        End Set
    End Property
    Public ReadOnly Property ContadeCustoFinal As PlanoDeConta
        Get
            If _ContadeCustoFinal Is Nothing And _CodigoContadeCustoFinal.Length > 0 Then _ContadeCustoFinal = New PlanoDeConta("", 0, _CodigoContadeCustoFinal)
            Return _ContadeCustoFinal
        End Get
    End Property


    Public Property CodigoContaPatrimonio As String
        Get
            Return _CodigoContaPatrimonio
        End Get
        Set(value As String)
            _CodigoContaPatrimonio = value
            _ContaPatrimonio = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaPatrimonio As PlanoDeConta
        Get
            If _ContaPatrimonio Is Nothing And Me.CodigoContaPatrimonio.Length > 0 Then _ContaPatrimonio = New PlanoDeConta("", 0, Me.CodigoContaPatrimonio)
            Return _ContaPatrimonio
        End Get
    End Property


    Public Property CodigoContaGrupoEstoque As String
        Get
            Return _CodigoContaGrupoEstoque
        End Get
        Set(value As String)
            _CodigoContaGrupoEstoque = value
            _ContaGrupoEstoque = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaGrupoEstoque As PlanoDeConta
        Get
            If _ContaGrupoEstoque Is Nothing And _CodigoContaGrupoEstoque.Length > 0 Then _ContaGrupoEstoque = New PlanoDeConta("", 0, _CodigoContaGrupoEstoque)
            Return _ContaGrupoEstoque
        End Get
    End Property

    Public Property CodigoContaEstoqueEmNossoPoder As String
        Get
            Return _CodigoContaEstoqueEmNossoPoder
        End Get
        Set(value As String)
            _CodigoContaEstoqueEmNossoPoder = value
            _ContaEstoqueEmNossoPoder = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaEstoqueEmNossoPoder As PlanoDeConta
        Get
            If _ContaEstoqueEmNossoPoder Is Nothing And _CodigoContaEstoqueEmNossoPoder.Length > 0 Then _ContaEstoqueEmNossoPoder = New PlanoDeConta("", 0, _CodigoContaEstoqueEmNossoPoder)
            Return _ContaEstoqueEmNossoPoder
        End Get
    End Property

    Public Property CodigoContaEstoqueEmPoderDeTerceiros As String
        Get
            Return _CodigoContaEstoqueEmPoderDeTerceiros
        End Get
        Set(value As String)
            _CodigoContaEstoqueEmPoderDeTerceiros = value
            _ContaEstoqueEmPoderDeTerceiros = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaEstoqueEmPoderDeTerceiros As PlanoDeConta
        Get
            If _ContaEstoqueEmPoderDeTerceiros Is Nothing And _CodigoContaEstoqueEmPoderDeTerceiros.Length > 0 Then _ContaEstoqueEmPoderDeTerceiros = New PlanoDeConta("", 0, _CodigoContaEstoqueEmPoderDeTerceiros)
            Return _ContaEstoqueEmPoderDeTerceiros
        End Get
    End Property

    Public Property CodigoContaVariacaoMonetariaPassiva() As String
        Get
            Return _CodigoContaVariacaoMonetariaPassiva
        End Get
        Set(ByVal value As String)
            _CodigoContaVariacaoMonetariaPassiva = value
            _ContaVariacaoMonetariaPassiva = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaVariacaoMonetariaPassiva() As PlanoDeConta
        Get
            If _ContaVariacaoMonetariaPassiva Is Nothing And _CodigoContaVariacaoMonetariaPassiva.Length > 0 Then _ContaVariacaoMonetariaPassiva = New PlanoDeConta("", 0, _CodigoContaVariacaoMonetariaPassiva)
            Return _ContaVariacaoMonetariaPassiva
        End Get
    End Property

    Public Property CodigoContaVariacaoMonetariaAtiva() As String
        Get
            Return _CodigoContaVariacaoMonetariaAtiva
        End Get
        Set(ByVal value As String)
            _CodigoContaVariacaoMonetariaAtiva = value
            _ContaVariacaoMonetariaAtiva = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaVariacaoMonetariaAtiva() As PlanoDeConta
        Get
            If _ContaVariacaoMonetariaAtiva Is Nothing And _CodigoContaVariacaoMonetariaAtiva.Length > 0 Then _ContaVariacaoMonetariaAtiva = New PlanoDeConta("", 0, _CodigoContaVariacaoMonetariaAtiva)
            Return _ContaVariacaoMonetariaAtiva
        End Get
    End Property

    Public Property CodigoContaVariacaoMonetariaCliente() As String
        Get
            Return _CodigoContaVariacaoMonetariaCliente
        End Get
        Set(ByVal value As String)
            _CodigoContaVariacaoMonetariaCliente = value
            _ContaVariacaoMonetariaCliente = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaVariacaoMonetariaCliente() As PlanoDeConta
        Get
            If _ContaVariacaoMonetariaCliente Is Nothing And _CodigoContaVariacaoMonetariaCliente.Length > 0 Then _ContaVariacaoMonetariaCliente = New PlanoDeConta("", 0, _CodigoContaVariacaoMonetariaCliente)
            Return _ContaVariacaoMonetariaCliente
        End Get
    End Property

    Public Property CodigoContaVariacaoMonetariaFornecedor() As String
        Get
            Return _CodigoContaVariacaoMonetariaFornecedor
        End Get
        Set(ByVal value As String)
            _CodigoContaVariacaoMonetariaFornecedor = value
            _ContaVariacaoMonetariaFornecedor = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaVariacaoMonetariafornecedor() As PlanoDeConta
        Get
            If _ContaVariacaoMonetariaFornecedor Is Nothing And _CodigoContaVariacaoMonetariaFornecedor.Length > 0 Then _ContaVariacaoMonetariaFornecedor = New PlanoDeConta("", 0, _CodigoContaVariacaoMonetariaFornecedor)
            Return _ContaVariacaoMonetariaFornecedor
        End Get
    End Property

    Public Property CodigoContaGrupoBanco() As String
        Get
            Return _CodigoContaGrupoBanco
        End Get
        Set(ByVal value As String)
            _CodigoContaGrupoBanco = value
            _ContaGrupoBanco = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaGrupoBanco() As PlanoDeConta
        Get
            If _ContaGrupoBanco Is Nothing And _CodigoContaGrupoBanco.Length > 0 Then _ContaGrupoBanco = New PlanoDeConta("", 0, _CodigoContaGrupoBanco)
            Return _ContaGrupoBanco
        End Get
    End Property

    Public Property CodigoContaJuroPago As String
        Get
            Return _CodigoContaJuroPago
        End Get
        Set(ByVal value As String)
            _CodigoContaJuroPago = value
            _ContaJuroPago = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaJuroPago() As PlanoDeConta
        Get
            If _ContaJuroPago Is Nothing And _CodigoContaJuroPago.Length > 0 Then _ContaJuroPago = New PlanoDeConta("", 0, _CodigoContaJuroPago)
            Return _ContaJuroPago
        End Get
    End Property

    Public Property CodigoContaJuroRecebido As String
        Get
            Return _CodigoContaJuroRecebido
        End Get
        Set(ByVal value As String)
            _CodigoContaJuroRecebido = value
            _ContaJuroRecebido = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaJuroRecebido() As PlanoDeConta
        Get
            If _ContaJuroRecebido Is Nothing And _CodigoContaJuroRecebido.Length > 0 Then _ContaJuroRecebido = New PlanoDeConta("", 0, _CodigoContaJuroRecebido)
            Return _ContaJuroRecebido
        End Get
    End Property

    Public Property CodigoContaGrupoDuplicatasDescontada As String
        Get
            Return _CodigoContaGrupoDuplicatasDescontada
        End Get
        Set(ByVal value As String)
            _CodigoContaGrupoDuplicatasDescontada = value
            _ContaGrupoDuplicatasDescontada = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaGrupoDuplicatasDescontada() As PlanoDeConta
        Get
            If _ContaGrupoDuplicatasDescontada Is Nothing And _CodigoContaGrupoDuplicatasDescontada.Length > 0 Then _ContaGrupoDuplicatasDescontada = New PlanoDeConta("", 0, _CodigoContaGrupoDuplicatasDescontada)
            Return _ContaGrupoDuplicatasDescontada
        End Get
    End Property

    Public Property CodigoContaGrupoComissoes As String
        Get
            Return _CodigoContaGrupoComissoes
        End Get
        Set(value As String)
            _CodigoContaGrupoComissoes = value
            _ContaGrupoComissoes = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaGrupoComissoes() As PlanoDeConta
        Get
            If _ContaGrupoComissoes Is Nothing And _CodigoContaGrupoComissoes.Length > 0 Then _ContaGrupoComissoes = New PlanoDeConta("", 0, _CodigoContaGrupoComissoes)
            Return _ContaGrupoComissoes
        End Get
    End Property

    Public Property CodigoContaAdiantamentoDeFrete As String
        Get
            Return _CodigoContaAdiantamentoDeFrete
        End Get
        Set(value As String)
            _CodigoContaAdiantamentoDeFrete = value
            _ContaAdiantamentoDeFrete = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaAdiantamentoDeFrete() As PlanoDeConta
        Get
            If _ContaAdiantamentoDeFrete Is Nothing And _CodigoContaAdiantamentoDeFrete.Length > 0 Then _ContaAdiantamentoDeFrete = New PlanoDeConta("", 0, _CodigoContaAdiantamentoDeFrete)
            Return _ContaAdiantamentoDeFrete
        End Get
    End Property

    Public Property CodigoContaPedagioDeFrete As String
        Get
            Return _CodigoContaPedagioDeFrete
        End Get
        Set(value As String)
            _CodigoContaPedagioDeFrete = value
            _ContaPedagioDeFrete = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaPedagioDeFrete() As PlanoDeConta
        Get
            If _ContaPedagioDeFrete Is Nothing And _CodigoContaPedagioDeFrete.Length > 0 Then _ContaPedagioDeFrete = New PlanoDeConta("", 0, _CodigoContaPedagioDeFrete)
            Return _ContaPedagioDeFrete
        End Get
    End Property


    Public Property CodigoContaGrupoTedDoc As String
        Get
            Return _CodigoContaGrupoTedDoc
        End Get
        Set(value As String)
            _CodigoContaGrupoTedDoc = value
            _ContaGrupoTedDoc = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaGrupoTedDoc() As PlanoDeConta
        Get
            If _ContaGrupoTedDoc Is Nothing And _CodigoContaGrupoTedDoc.Length > 0 Then _ContaGrupoTedDoc = New PlanoDeConta("", 0, _CodigoContaGrupoTedDoc)
            Return _ContaGrupoTedDoc
        End Get
    End Property

    Public Property CodigoContaFornecedorFrete As String
        Get
            Return _CodigoContaFornecedorFrete
        End Get
        Set(value As String)
            _CodigoContaFornecedorFrete = value
            _ContaFornecedorFrete = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaFornecedorFrete() As PlanoDeConta
        Get
            If _ContaFornecedorFrete Is Nothing And _CodigoContaFornecedorFrete.Length > 0 Then _ContaFornecedorFrete = New PlanoDeConta("", 0, _CodigoContaFornecedorFrete)
            Return _ContaFornecedorFrete
        End Get
    End Property

    Public Property CodigoContaCaixaCompensacao As String
        Get
            Return _CodigoContaCaixaCompensacao
        End Get
        Set(value As String)
            _CodigoContaCaixaCompensacao = value
            _ContaCaixaCompensacao = Nothing
        End Set
    End Property
    Public ReadOnly Property ContaCaixaCompensacao() As PlanoDeConta
        Get
            Return _ContaCaixaCompensacao
        End Get
    End Property

    Public Property CodigoContaDescontoObtido As String
        Get
            Return _CodigoContaDescontoObtido
        End Get
        Set(ByVal value As String)
            _CodigoContaDescontoObtido = value
            _ContaDescontoObtido = Nothing
        End Set
    End Property
    Public Property ContaDescontoObtido() As PlanoDeConta
        Get
            If _ContaDescontoObtido Is Nothing And _CodigoContaDescontoObtido.Length > 0 Then _ContaDescontoObtido = New PlanoDeConta("", 0, _CodigoContaDescontoObtido)
            Return _ContaDescontoObtido
        End Get
        Set(ByVal value As PlanoDeConta)
            _ContaDescontoObtido = value
        End Set
    End Property

    Public Property CodigoContaDescontoConcedido As String
        Get
            Return _CodigoContaDescontoConcedido
        End Get
        Set(ByVal value As String)
            _CodigoContaDescontoConcedido = value
            _ContaDescontoConcedido = Nothing
        End Set
    End Property
    Public Property ContaDescontoConcedido() As PlanoDeConta
        Get
            If _ContaDescontoConcedido Is Nothing And _CodigoContaDescontoConcedido.Length > 0 Then _ContaDescontoConcedido = New PlanoDeConta("", 0, _CodigoContaDescontoConcedido)
            Return _ContaDescontoConcedido
        End Get
        Set(ByVal value As PlanoDeConta)
            _ContaDescontoConcedido = value
        End Set
    End Property


    Public Property CodigoProdutoDeFrete() As String
        Get
            Return _CodigoProdutoDeFrete
        End Get
        Set(ByVal value As String)
            _CodigoProdutoDeFrete = value
        End Set
    End Property

    Public Property CodigoProdutoDeEstadia() As String
        Get
            Return _CodigoProdutoDeEstadia
        End Get
        Set(ByVal value As String)
            _CodigoProdutoDeEstadia = value
        End Set
    End Property

    Public Property CodigoProdutoDeMDFe() As String
        Get
            Return _CodigoProdutoDeMDFe
        End Get
        Set(ByVal value As String)
            _CodigoProdutoDeMDFe = value
        End Set
    End Property

    Public Property CodigoUnidadeNegocio() As String
        Get
            Return _CodigoUnidadeNegocio
        End Get
        Set(ByVal value As String)
            _CodigoUnidadeNegocio = value
        End Set
    End Property

    Public Property ViasNFE() As Integer
        Get
            Return _ViasNFE
        End Get
        Set(ByVal value As Integer)
            _ViasNFE = value
        End Set
    End Property

    Public Property PrazoCancelamentoNFE() As Integer
        Get
            Return _PrazoConcelamentoNFE
        End Get
        Set(ByVal value As Integer)
            _PrazoConcelamentoNFE = value
        End Set
    End Property

    Public Property ControlaEmissaoCheque() As Boolean
        Get
            Return _ControlaEmissaoCheque
        End Get
        Set(ByVal value As Boolean)
            _ControlaEmissaoCheque = value
        End Set
    End Property

    Public Property ControlaDataMovimentoNFG() As Boolean
        Get
            Return _ControlaDataMovimentoNFG
        End Get
        Set(ByVal value As Boolean)
            _ControlaDataMovimentoNFG = value
        End Set
    End Property

    Public Property ObservacaoSefazNFE() As String
        Get
            Return _ObservacaoSefazNFE
        End Get
        Set(ByVal value As String)
            _ObservacaoSefazNFE = value
        End Set
    End Property

    Public Property Crt() As Integer
        Get
            Return _Crt
        End Get
        Set(ByVal value As Integer)
            _Crt = value
        End Set
    End Property

    Public Property ArquivoNFE() As Boolean
        Get
            Return _ArquivoNFE
        End Get
        Set(ByVal value As Boolean)
            _ArquivoNFE = value
        End Set
    End Property

    Public Property ConferenciaNFE() As Boolean
        Get
            Return _ConferenciaNFE
        End Get
        Set(ByVal value As Boolean)
            _ConferenciaNFE = value
        End Set
    End Property

    Public Property PathDownloadNFe() As String
        Get
            Return _PathDownloadNFe
        End Get
        Set(value As String)
            _PathDownloadNFe = value
        End Set
    End Property


    Public Property DiasNFRetroativa() As Integer
        Get
            Return _DiasNFRetroativa
        End Get
        Set(ByVal value As Integer)
            _DiasNFRetroativa = value
        End Set
    End Property

    Public Property ObrigaEncargo As Boolean
        Get
            Return _ObrigaEncargo
        End Get
        Set(value As Boolean)
            _ObrigaEncargo = value
        End Set
    End Property

    Public Property ApenasExcedenteTolerancia() As Boolean
        Get
            Return _ApenasExcedenteTolerancia
        End Get
        Set(ByVal value As Boolean)
            _ApenasExcedenteTolerancia = value
        End Set
    End Property

    Public Property DataFinanceiro As Date
        Get
            Return _DataFinanceiro
        End Get
        Set(value As Date)
            _DataFinanceiro = value
        End Set
    End Property

    Public Property CodigoCnae As String
        Get
            Return _CodigoCnae
        End Get
        Set(value As String)
            _CodigoCnae = value
        End Set
    End Property

    Public Property CodigoNaturezaJuridica As Integer
        Get
            Return _CodigoNaturezaJuridica
        End Get
        Set(value As Integer)
            _CodigoNaturezaJuridica = value
        End Set
    End Property

    Public Property BaixaFinanceiroPorLote() As Boolean
        Get
            Return _BaixaFinanceiroPorLote
        End Get
        Set(ByVal value As Boolean)
            _BaixaFinanceiroPorLote = value
        End Set
    End Property

    Public Property ParametrosECF As ClienteXEmpresaXECF
        Get
            If _ParametrosECF Is Nothing Then _ParametrosECF = New ClienteXEmpresaXECF(Me)
            Return _ParametrosECF
        End Get
        Set(value As ClienteXEmpresaXECF)
            _ParametrosECF = value
        End Set
    End Property

    Public Property ObrigaNfProdutor() As Boolean
        Get
            Return _ObrigaNfProdutor
        End Get
        Set(ByVal value As Boolean)
            _ObrigaNfProdutor = value
        End Set
    End Property

    Public Property ObrigaChaveNf() As Boolean
        Get
            Return _ObrigaChaveNf
        End Get
        Set(ByVal value As Boolean)
            _ObrigaChaveNf = value
        End Set
    End Property

    Public Property ObrigaChaveNfg() As Boolean
        Get
            Return _ObrigaChaveNfg
        End Get
        Set(ByVal value As Boolean)
            _ObrigaChaveNfg = value
        End Set
    End Property

    Public Property UsarDescricaoProduto() As Boolean
        Get
            Return _UsarDescricaoProduto
        End Get
        Set(ByVal value As Boolean)
            _UsarDescricaoProduto = value
        End Set
    End Property

    Public Property UsarRegistroMinAgr() As Boolean
        Get
            Return _UsarRegistroMinAgr
        End Get
        Set(ByVal value As Boolean)
            _UsarRegistroMinAgr = value
        End Set
    End Property

    Public Property ObrigaNavio() As Boolean
        Get
            Return _ObrigaNavio
        End Get
        Set(ByVal value As Boolean)
            _ObrigaNavio = value
        End Set
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property DataInclusao() As DateTime
        Get
            Return _DataInclusao
        End Get
        Set(ByVal value As DateTime)
            _DataInclusao = value
        End Set
    End Property

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property DataAlteracao() As DateTime
        Get
            Return _DataAlteracao
        End Get
        Set(ByVal value As DateTime)
            _DataAlteracao = value
        End Set
    End Property

    Public Property EmitirCTe As Boolean
        Get
            Return _EmitirCTe
        End Get
        Set(value As Boolean)
            _EmitirCTe = value
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
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = "MERGE ClientesXEmpresas AS Target" & vbCrLf &
                      "USING (SELECT '" & Empresa_id & "' as Empresa_id, " & EndEmpresa_Id & " as EndEmpresa_id) AS Source" & vbCrLf &
                      "   ON Target.Empresa_id    = Source.Empresa_id" & vbCrLf &
                      "  AND Target.EndEmpresa_id = Source.EndEmpresa_id" & vbCrLf &
                      "WHEN MATCHED THEN" & vbCrLf &
                      " UPDATE SET" & vbCrLf &
                      "    AtividadeEconomica         ='" & Me.AtividadeEconomica & "'" & vbCrLf &
                      "  , AtividadeEstadual          ='" & Me.AtividadeEstadual & "'" & vbCrLf &
                      "  , DatadeFundacao             ='" & Me.DataDeFundacao.ToString("yyyy/MM/dd") & "'" & vbCrLf &
                      "  , RegistroNaJunta            ='" & Me.RegistroNaJunta & "'" & vbCrLf &
                      "  , EstadodaJunta              ='" & Me.EstadoDaJunta & "'" & vbCrLf &
                      "  , NomeDoContador             ='" & Me.NomeDoContador & "'" & vbCrLf &
                      "  , CPFContador                ='" & Me.CPFContador.Replace(".", "").Replace("/", "").Replace("-", "") & "'" & vbCrLf &
                      "  , CNPJContador                ='" & Me.CNPJContador.Replace(".", "").Replace("/", "").Replace("-", "") & "'" & vbCrLf &
                      "  , CRCContador                ='" & Me.CRCContador & "'" & vbCrLf &
                      "  , EstadoExpCRC               ='" & Me.EstadoExpCRC & "'" & vbCrLf &
                      "  , TelefoneContador           ='" & Me.TelefoneContador & "'" & vbCrLf &
                      "  , EmailContador              ='" & Me.EmailContador & "'" & vbCrLf &
                      "  , NomeDoTitular              ='" & Me.NomeDoTitular & "'" & vbCrLf &
                      "  , CPFTitular                 ='" & Me.CPFTitular.Replace(".", "").Replace("/", "").Replace("-", "") & "'" & vbCrLf &
                      "  , CondicaoTitular            ='" & Me.QualificacaoTitular & "'" & vbCrLf &
                      "  , Matriz                     ='" & Me.Matriz & "'" & vbCrLf &
                      "  , InscricaoMunicipal         ='" & Me.InscricaoMunicipal & "'" & vbCrLf &
                      "  , RegistroEp                 ='" & Me.RegistroEp & "'" & vbCrLf &
                      "  , RegistroEi                 ='" & Me.RegistroEi & "'" & vbCrLf &
                      "  , RegistroEc                 ='" & Me.RegistroEc & "'" & vbCrLf &
                      "  , Marca                      ='" & Me.Marca & "'" & vbCrLf &
                      "  , RegistroNire               ='" & Me.RegistroNire & "'" & vbCrLf &
                      "  , DataRegistroNire           ='" & Me.DataRegistroNire.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      "  , QualificacaoContador       ='" & Me.QualificacaoContador & "'" & vbCrLf &
                      "  , CodigoQualificacaoContador = " & Me.CodigoQualificacaoContador & vbCrLf &
                      "  , QualificacaoTitular        ='" & Me.QualificacaoTitular & "'" & vbCrLf &
                      "  , CodigoQualificacaoTitular  = " & Me.CodigoQualificacaoTitular & vbCrLf &
                      "  , CodigoDeRelacionamento     = " & Me.CodigoDeRelacionamento & vbCrLf &
                      "  , EscrituracaoContabil       ='" & Me.EscrituracaoContabil & "'" & vbCrLf &
                      "  , NotaFiscalEletronica       ='" & IIf(Me.NotaFiscalEletronica, "S", "N") & "'" & vbCrLf &
                      "  , EmitirCTe                  ='" & Me.EmitirCTe.ToString & "'" & vbCrLf &
                      "  , ContadeCustoInicial        ='" & Funcoes.AlinharEsquerda(Me.CodigoContadeCustoInicial.ToString, 9, "0") & "'" & vbCrLf &
                      "  , ContadeCustoFinal          ='" & Funcoes.AlinharEsquerda(Me.CodigoContadeCustoFinal.ToString, 9, "9") & "'" & vbCrLf &
                      "  , ContaPatrimonio            ='" & Me.CodigoContaPatrimonio & "'" & vbCrLf &
                      "  , Crt                        = " & Me.Crt & vbCrLf &
                      "  , CertidaoNegativa           ='" & Me.CertidaoNegativa.ToString & "'" & vbCrLf &
                      "  , FluxoDeCaixa               ='" & Me.FluxoDeCaixa.ToString & "'" & vbCrLf &
                      "  , SugereDeposito             ='" & Me.SugereDeposito.ToString & "'" & vbCrLf &
                      "  , NossaEmissao               ='" & Me.NossaEmissao.ToString & "'" & vbCrLf &
                      "  , RegistroDeExportacao       ='" & Me.RegistroDeExportacao.ToString & "'" & vbCrLf &
                      "  , ContaVariacaoMonetariaAtiva      ='" & Me.CodigoContaVariacaoMonetariaAtiva & "'" & vbCrLf &
                      "  , ContaVariacaoMonetariaPassiva    ='" & Me.CodigoContaVariacaoMonetariaPassiva & "'" & vbCrLf &
                      "  , ContaVariacaoMonetariaCliente    ='" & Me.CodigoContaVariacaoMonetariaCliente & "'" & vbCrLf &
                      "  , ContaVariacaoMonetariaFornecedor ='" & Me.CodigoContaVariacaoMonetariaFornecedor & "'" & vbCrLf &
                      "  , ContaGrupoBanco                  ='" & Me.CodigoContaGrupoBanco & "'" & vbCrLf &
                      "  , FretePedido                      ='" & Me.FretePedido.ToString & "'" & vbCrLf &
                      "  , PedidoBloqueado                  ='" & Me.PedidoBloqueado.ToString & "'" & vbCrLf &
                      "  , LiberaCarregamento               ='" & Me.LiberaCarregamento.ToString & "'" & vbCrLf &
                      "  , ContaJuroPago                    ='" & Me.CodigoContaJuroPago & "'" & vbCrLf &
                      "  , ContaJuroRecebido                ='" & Me.CodigoContaJuroRecebido & "'" & vbCrLf &
                      "  , ContaGrupoDuplicatasDescontada   ='" & Me.CodigoContaGrupoDuplicatasDescontada & "'" & vbCrLf &
                      "  , ContaGrupoEstoque                ='" & Me.CodigoContaGrupoEstoque & "'" & vbCrLf &
                      "  , ContaGrupoComissoes              ='" & Me.CodigoContaGrupoComissoes & "'" & vbCrLf &
                      "  , ContaAdiantamentoDeFrete         ='" & Me.CodigoContaAdiantamentoDeFrete & "'" & vbCrLf &
                      "  , ContaPedagioDeFrete              ='" & Me.CodigoContaPedagioDeFrete & "'" & vbCrLf &
                      "  , CodigoProdutoDeFrete             ='" & Me.CodigoProdutoDeFrete & "'" & vbCrLf &
                      "  , CodigoProdutoDeEstadia           ='" & Me.CodigoProdutoDeEstadia & "'" & vbCrLf &
                      "  , CodigoProdutoDeMDFe              ='" & Me.CodigoProdutoDeMDFe & "'" & vbCrLf &
                      "  , Servidor                         ='" & Me.Servidor & "'" & vbCrLf &
                      "  , ViasNFE                          ='" & Me.ViasNFE & "'" & vbCrLf &
                      "  , PrazoDeCancelamentoNFE           = " & Me.PrazoCancelamentoNFE.ToSqlNULL & vbCrLf &
                      "  , ControlaEmissaoCheque            ='" & Me.ControlaEmissaoCheque.ToString & "'" & vbCrLf &
                      "  , ControlaDataMovimentoNFG         ='" & Me.ControlaDataMovimentoNFG.ToString & "'" & vbCrLf &
                      "  , ObservacaoSefazNFE               ='" & Me.ObservacaoSefazNFE & "'" & vbCrLf &
                      "  , ArquivoNFE                       ='" & Me.ArquivoNFE.ToString & "'" & vbCrLf &
                      "  , PathDownloadNFe                  ='" & Me.PathDownloadNFe & "'" & vbCrLf &
                      "  , DiasNFRetroativa                 = " & Me.DiasNFRetroativa & vbCrLf &
                      "  , ConferenciaNFE                   ='" & Me.ConferenciaNFE.ToString & "'" & vbCrLf &
                      "  , ContaGrupoTEDDOC                 ='" & Me.CodigoContaGrupoTedDoc & "'" & vbCrLf &
                      "  , ContaDescontoConcedido           ='" & Me.CodigoContaDescontoConcedido & "'" & vbCrLf &
                      "  , ContaDescontoObtido              ='" & Me.CodigoContaDescontoObtido & "'" & vbCrLf &
                      "  , ContaFornecedorFrete             ='" & Me.CodigoContaFornecedorFrete & "'" & vbCrLf &
                      "  , ContaCaixaCompensacao            ='" & Me.CodigoContaCaixaCompensacao & "'" & vbCrLf &
                      "  , ContaEstoqueEmNossoPoder         ='" & Me.CodigoContaEstoqueEmNossoPoder & "'" & vbCrLf &
                      "  , ContaEstoqueEmPoderDeTerceiros   ='" & Me.CodigoContaEstoqueEmPoderDeTerceiros & "'" & vbCrLf &
                      "  , DataFinanceiro                   = " & Me.DataFinanceiro.ToSqlNULL & vbCrLf &
                      "  , ObrigaEncargo                    ='" & Me.ObrigaEncargo.ToString & "'" & vbCrLf &
                      "  , CNAE                             = " & Me.CodigoCnae.ToSqlNULL & vbCrLf &
                      "  , NatJur                           = " & Me.CodigoNaturezaJuridica.ToSqlNULL & vbCrLf &
                      "  , BaixaFinanceiroPorLote           ='" & Me.BaixaFinanceiroPorLote.ToString & "'" & vbCrLf &
                      "  , ObrigaNfProdutor                 ='" & Me.ObrigaNfProdutor.ToString & "'" & vbCrLf &
                      "  , ObrigaChaveNf                    ='" & Me.ObrigaChaveNf.ToString & "'" & vbCrLf &
                      "  , ObrigaChaveNfg                   ='" & Me.ObrigaChaveNfg.ToString & "'" & vbCrLf &
                      "  , UsarDescricaoProduto             ='" & Me.UsarDescricaoProduto.ToString & "'" & vbCrLf &
                      "  , UsarRegistroMinAgr               ='" & Me.UsarRegistroMinAgr.ToString & "'" & vbCrLf &
                      "  , ObrigaNavio                      ='" & Me.ObrigaNavio.ToString & "'" & vbCrLf &
                      "	 , UsuarioAlteracao                 ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                      "	 , UsuarioAlteracaoData             = getdate() " & vbCrLf &
                      "WHEN NOT MATCHED BY TARGET THEN" & vbCrLf &
                      " INSERT(Empresa_Id, EndEmpresa_Id" & vbCrLf &
                      ", AtividadeEconomica, AtividadeEstadual" & vbCrLf &
                      ", DatadeFundacao, RegistroNaJunta" & vbCrLf &
                      ", EstadodaJunta" & vbCrLf &
                      ", NomeDoContador" & vbCrLf &
                      ", CPFContador, CNPJContador, CRCContador, EstadoExpCRC,TelefoneContador,EmailContador" & vbCrLf &
                      ", NomeDoTitular, CPFTitular" & vbCrLf &
                      ", CondicaoTitular, Matriz" & vbCrLf &
                      ", InscricaoMunicipal" & vbCrLf &
                      ", RegistroEp, RegistroEi, RegistroEc" & vbCrLf &
                      ", Marca" & vbCrLf &
                      ", RegistroNire, DataRegistroNire" & vbCrLf &
                      ", QualificacaoContador, CodigoQualificacaoContador, QualificacaoTitular, CodigoQualificacaoTitular" & vbCrLf &
                      ", CodigoDeRelacionamento" & vbCrLf &
                      ", EscrituracaoContabil" & vbCrLf &
                      ", NotaFiscalEletronica" & vbCrLf &
                      ", EmitirCTe" & vbCrLf &
                      ", ContadeCustoInicial, ContadeCustoFinal, ContaPatrimonio" & vbCrLf &
                      ", Crt" & vbCrLf &
                      ", CertidaoNegativa" & vbCrLf &
                      ", FluxoDeCaixa" & vbCrLf &
                      ", SugereDeposito" & vbCrLf &
                      ", NossaEmissao" & vbCrLf &
                      ", RegistroDeExportacao " & vbCrLf &
                      ", ContaVariacaoMonetariaAtiva, ContaVariacaoMonetariaPassiva, ContaVariacaoMonetariaCliente, ContaVariacaoMonetariaFornecedor, ContaGrupoBanco" & vbCrLf &
                      ", FretePedido" & vbCrLf &
                      ", PedidoBloqueado" & vbCrLf &
                      ", LiberaCarregamento" & vbCrLf &
                      ", ContaJuroPago, ContaJuroRecebido, ContaGrupoDuplicatasDescontada, ContaGrupoEstoque, ContaGrupoComissoes" & vbCrLf &
                      ", ContaAdiantamentoDeFrete, ContaPedagioDeFrete, CodigoProdutoDeFrete, CodigoProdutoDeEstadia, CodigoProdutoDeMDFe" & vbCrLf &
                      ", Servidor" & vbCrLf &
                      ", ViasNFE, PrazoDeCancelamentoNFE" & vbCrLf &
                      ", ControlaEmissaoCheque, ControlaDataMovimentoNFG" & vbCrLf &
                      ", ObservacaoSefazNFE" & vbCrLf &
                      ", ArquivoNFE, PathDownloadNFe, DiasNFRetroativa" & vbCrLf &
                      ", ConferenciaNFE" & vbCrLf &
                      ", ContaGrupoTEDDOC, ContaDescontoConcedido, ContaDescontoObtido, ContaFornecedorFrete, ContaCaixaCompensacao, ContaEstoqueEmNossoPoder, ContaEstoqueEmPoderDeTerceiros " & vbCrLf &
                      ", DataFinanceiro, ObrigaEncargo, CNAE, NatJur, BaixaFinanceiroPorLote, ObrigaNfProdutor, ObrigaChaveNf, ObrigaChaveNfg, UsarDescricaoProduto, UsarRegistroMinAgr " & vbCrLf &
                      ", ObrigaNavio, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf &
                      " VALUES " & vbCrLf &
                      "('" & Me.Empresa_id & "'" & vbCrLf &
                      "," & Me.EndEmpresa_Id & vbCrLf &
                      ", '" & Me.AtividadeEconomica & "'" & vbCrLf &
                      ", '" & Me.AtividadeEstadual & "'" & vbCrLf &
                      ", '" & Me.DataDeFundacao.ToString("yyyy/MM/dd") & "'" & vbCrLf &
                      ", '" & Me.RegistroNaJunta & "'" & vbCrLf &
                      ", '" & Me.EstadoDaJunta & "'" & vbCrLf &
                      ", '" & Me.NomeDoContador & "'" & vbCrLf &
                      ", '" & Me.CPFContador.Replace(".", "").Replace("/", "").Replace("-", "") & "'" & vbCrLf &
                      ", '" & Me.CNPJContador.Replace(".", "").Replace("/", "").Replace("-", "") & "'" & vbCrLf &
                      ", '" & Me.CRCContador & "'" & vbCrLf &
                      ", '" & Me.EstadoExpCRC & "'" & vbCrLf &
                      ", '" & Me.TelefoneContador & "'" & vbCrLf &
                      ", '" & Me.EmailContador & "'" & vbCrLf &
                      ", '" & Me.NomeDoTitular & "'" & vbCrLf &
                      ", '" & Me.CPFTitular.Replace(".", "").Replace("/", "").Replace("-", "") & "'" & vbCrLf &
                      ", '" & Me.QualificacaoTitular & "'" & vbCrLf &
                      ", '" & Me.Matriz & "'" & vbCrLf &
                      ", '" & Me.InscricaoMunicipal & "'" & vbCrLf &
                      ", '" & Me.RegistroEp & "'" & vbCrLf &
                      ", '" & Me.RegistroEi & "'" & vbCrLf &
                      ", '" & Me.RegistroEc & "'" & vbCrLf &
                      ", '" & Me.Marca & "'" & vbCrLf &
                      ", '" & Me.RegistroNire & "'" & vbCrLf &
                      ", '" & Me.DataRegistroNire.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      ", '" & Me.QualificacaoContador & "'" & vbCrLf &
                      ",  " & Me.CodigoQualificacaoContador & vbCrLf &
                      ", '" & Me.QualificacaoTitular & "'" & vbCrLf &
                      ",  " & Me.CodigoQualificacaoTitular & vbCrLf &
                      ",  " & Me.CodigoDeRelacionamento & vbCrLf &
                      ", '" & Me.EscrituracaoContabil & "'" & vbCrLf &
                      ", '" & IIf(Me.NotaFiscalEletronica, "S", "N") & "'" & vbCrLf &
                      ", '" & Me.EmitirCTe.ToString & "'" & vbCrLf &
                      ", '" & Funcoes.AlinharEsquerda(Me.CodigoContadeCustoInicial.ToString, 9, "0") & "'" & vbCrLf &
                      ", '" & Funcoes.AlinharEsquerda(Me.CodigoContadeCustoFinal.ToString, 9, "9") & "'" & vbCrLf &
                      ", '" & Me.CodigoContaPatrimonio & "'" & vbCrLf &
                      ",  " & Me.Crt & vbCrLf &
                      ", '" & Me.CertidaoNegativa.ToString & "'" & vbCrLf &
                      ", '" & Me.FluxoDeCaixa.ToString & "'" & vbCrLf &
                      ", '" & Me.SugereDeposito.ToString & "'" & vbCrLf &
                      ", '" & Me.NossaEmissao.ToString & "'" & vbCrLf &
                      ", '" & Me.RegistroDeExportacao.ToString & "'" & vbCrLf &
                      ", '" & Me.CodigoContaVariacaoMonetariaAtiva & "'" & vbCrLf &
                      ", '" & Me.CodigoContaVariacaoMonetariaPassiva & "'" & vbCrLf &
                      ", '" & Me.CodigoContaVariacaoMonetariaCliente & "'" & vbCrLf &
                      ", '" & Me.CodigoContaVariacaoMonetariaFornecedor & "'" & vbCrLf &
                      ", '" & Me.CodigoContaGrupoBanco & "'" & vbCrLf &
                      ", '" & Me.FretePedido.ToString & "'" & vbCrLf &
                      ", '" & Me.PedidoBloqueado.ToString & "'" & vbCrLf &
                      ", '" & Me.LiberaCarregamento.ToString & "'" & vbCrLf &
                      ", '" & Me.CodigoContaJuroPago & "'" & vbCrLf &
                      ", '" & Me.CodigoContaJuroRecebido & "'" & vbCrLf &
                      ", '" & Me.CodigoContaGrupoDuplicatasDescontada & "'" & vbCrLf &
                      ", '" & Me.CodigoContaGrupoEstoque & "'" & vbCrLf &
                      ", '" & Me.CodigoContaGrupoComissoes & "'" & vbCrLf &
                      ", '" & Me.CodigoContaAdiantamentoDeFrete & "'" & vbCrLf &
                      ", '" & Me.CodigoContaPedagioDeFrete & "'" & vbCrLf &
                      ", '" & Me.CodigoProdutoDeFrete & "'" & vbCrLf &
                      ", '" & Me.CodigoProdutoDeEstadia & "'" & vbCrLf &
                      ", '" & Me.CodigoProdutoDeMDFe & "'" & vbCrLf &
                      ", '" & Me.Servidor & "'" & vbCrLf &
                      ", '" & Me.ViasNFE & "'" & vbCrLf &
                      ",  " & Me.PrazoCancelamentoNFE.ToSqlNULL & vbCrLf &
                      ", '" & Me.ControlaEmissaoCheque.ToString & "'" & vbCrLf &
                      ", '" & Me.ControlaDataMovimentoNFG.ToString & "'" & vbCrLf &
                      ", '" & Me.ObservacaoSefazNFE & "'" & vbCrLf &
                      ", '" & Me.ArquivoNFE.ToString & "'" & vbCrLf &
                      ", '" & Me.PathDownloadNFe & "'" & vbCrLf &
                      ",  " & Me.DiasNFRetroativa & vbCrLf &
                      ", '" & Me.ConferenciaNFE.ToString & "'" & vbCrLf &
                      ", '" & Me.CodigoContaGrupoTedDoc & "'" & vbCrLf &
                      ", '" & Me.CodigoContaDescontoConcedido & "'" & vbCrLf &
                      ", '" & Me.CodigoContaDescontoObtido & "'" & vbCrLf &
                      ", '" & Me.CodigoContaFornecedorFrete & "'" & vbCrLf &
                      ", '" & Me.CodigoContaCaixaCompensacao & "'" & vbCrLf &
                      ", '" & Me.CodigoContaEstoqueEmNossoPoder & "'" & vbCrLf &
                      ", '" & Me.CodigoContaEstoqueEmPoderDeTerceiros & "'" & vbCrLf &
                      ",  " & Me.DataFinanceiro.ToSqlNULL & vbCrLf &
                      ",  " & Me.ObrigaEncargo.ToSql & vbCrLf &
                      ",  " & Me.CodigoCnae.ToSqlNULL & vbCrLf &
                      ",  " & Me.CodigoNaturezaJuridica.ToSqlNULL & vbCrLf &
                      ", '" & Me.BaixaFinanceiroPorLote.ToString & "'" & vbCrLf &
                      ", '" & Me.ObrigaNfProdutor.ToString & "'" & vbCrLf &
                      ", '" & Me.ObrigaChaveNf.ToString & "'" & vbCrLf &
                      ", '" & Me.ObrigaChaveNfg.ToString & "'" &
                      ", '" & Me.UsarDescricaoProduto.ToString & "'" &
                      ", '" & Me.UsarRegistroMinAgr.ToString & "'" &
                      ", '" & Me.ObrigaNavio.ToString & "'" &
                      ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                      ", getdate());"

                Sqls.Add(Sql)

                ParametrosECF.IUD = Me.IUD
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                Me.ParametrosECF.IUD = "D"
                SalvarTabelasRelacionadasSql(Sqls)

                Sql = "DELETE ClientesXEmpresas" & vbCrLf & _
                      " WHERE Empresa_id    ='" & Me.Empresa_id & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & Me.EndEmpresa_Id & vbCrLf
                Sqls.Add(Sql)
        End Select


    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        Me.ParametrosECF.SalvarSql(Sqls)
    End Sub

#End Region
End Class