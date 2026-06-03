Imports System.ComponentModel
Public Enum eSituacaoBancaria
    <Description("AGUARDANDO RETORNO BANCÁRIO")>
    AguardandoArquivoDeRetorno = 0
    <Description("ENTRADA CONFIRMADA")>
    EntradaConfirmada = 2
    <Description("ENTRADA REJEITADA")>
    EntradaRejeitada = 3
    <Description("ALTERAÇÃO DE DADOS")>
    AlteracaoDeDados = 4
    <Description("LIQUIDAÇÃO NORMAL")>
    LiquidacaoNormal = 6
    <Description("LIQUIDAÇÃO PARCIAL COBRANÇA INTELIGENTE (B2B)")>
    LiquidacaoParcial = 7
    <Description("LIQUIDAÇÃO EM CARTÓRIO")>
    LiquidacaoEmCartorio = 8
    <Description("BAIXA SIMPLES")>
    BaixaSimples = 9
    <Description("BAIXA POR TER SIDO LIQUIDADO")>
    BaixaPorLiquidado = 10
    <Description("VENCIMENTO ALTERADO")>
    VencimentoAlterado = 14
    <Description("BAIXAS REJEITADAS")>
    BaixasRejeitadas = 15
    <Description("INSTRUÇÕES REJEITADAS")>
    InstrucaoRejeitada = 16
    <Description("INSTRUÇÃO DE PROTESTO")>
    Protestar = 19
    <Description("TARIFA DE SUSTENTAÇÃO DE PROTESTO")>
    TarifaDeSustentacaoProtesto = 20
    <Description("NÃO PROTESTAR")>
    NaoProtestar = 21
    <Description("TARIFA DE TÍTULO ENVIADO A CARTÓRIO")>
    TarifaDeEnvioACartorio = 23
    <Description("INSTRUÇÃO DE PROTESTO (REJEITADA/SUSTADA/PENDENTE)")>
    TarifaDeSituacaoNoCartorio = 24
    <Description("ALEGAÇÕES DO PAGADOR")>
    AlegacoesDoPagador = 25
    <Description("TARIFA DE AVISO DE COBRANÇA")>
    TarifaAvisoDeCobranca = 26
    <Description("TARIFA DE EXTRATO POSIÇÃO")>
    TarifaDeExtrato = 27
    <Description("TARIFA DE RELAÇÃO DAS LIQUIDAÇÕES")>
    TarifaDeLiquidacao = 28
    <Description("TARIFA DE MANUTENÇÃO")>
    TarifaDeManutencao = 29
    <Description("BAIXA POR TER SIDO PROTESTADO")>
    BaixaPorProtesto = 32
    <Description("CUSTAS DE SUSTAÇÃO")>
    CustasDeSustacao = 34
    <Description("CUSTAS DE CARTÓRIO DISTRIBUIDOR")>
    CustasDeCartorio = 35
    <Description("TARIFA DE INSTRUÇÃO")>
    TarifaDeInstrucao = 38
    <Description("TARIFA DE OCORRÊNCIAS")>
    TarifaDeOcorrencia = 39
    <Description("TÍTULO DDA RECUSADO PELA CIP")>
    DdaRecusadoPelaCip = 53
    <Description("INSTRUÇÃO CANCELADA")>
    InstrucaoCancelada = 57
    <Description("TÍTULO EM CARTEIRA (EM SER)")>
    TituloEmCarteiraEmSer = 11
    <Description("BAIXA COM TRANSFERÊNCIA PARA DESCONTO")>
    BaixaTPDesconto = 47
End Enum
