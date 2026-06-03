Imports System.ComponentModel
Imports NGS.Lib.Uteis

Public Enum eTipoDeDocumentoFrete

    <Description("Circulação")> _
    Circulacao = 1

    <Description("Comprovação")> _
    Comprovacao = 2

    '<Description("Prestação de Serviço Próprio")> _
    'PrestacaoServicoProprio = 3

    <Description("Prestação de Serviço de Terceiros c/ Financeiro")> _
    PrestacaoServicoTerceirosComFinanceiro = 4

    <Description("Prestação de Serviço de Terceiros s/ Financeiro")> _
    PrestacaoServicoTerceirosSemFinanceiro = 5

    <Description("Complemento de Frete")> _
    Complemento = 6

    <Description("Estadia")> _
    Estadia = 7

    <Description("Anulação")> _
    Anulacao = 8

    <Description("Recibo de Pagamento a Autônomo(RPA)")> _
    RPA = 9
End Enum
