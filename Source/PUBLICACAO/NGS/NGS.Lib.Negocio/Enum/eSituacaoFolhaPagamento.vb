Imports System.ComponentModel
Public Enum eSituacaoFolhaPagamento
    <Description("WB - Aguardando retorno bancário")>
    AguardandoArquivoDeRetorno = 101
    <Description("BD - Inclusão Efetuada com Sucesso")>
    EntradaConfirmada = 102
    <Description("00 - Crédito ou Débito Efetivado - O pagamento foi confirmado")>
    Liquidacao = 103
    <Description("EE - Entrada Inválida, Verifique a documentação do bradesco")>
    Erro = 104
    <Description("EXCLUSÃO")>
    Exclusao = 105
    <Description("5T00 – Pagamento realizado em contrato na condição de TESTE")>
    LiquidacaoEmHomologacao = 106
End Enum

