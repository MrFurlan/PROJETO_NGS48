Imports System.ComponentModel

Public Enum eRetornoRemessaItau
    <Description("ENTRADA CONFIRMADA / DESCONTO ACEITO")>
    DescontoAceito = 2
    <Description("ENTRADA REJEITADA / DESCONTO RECUSADO")>
    DescontoRecusado = 3
    <Description("ALTERAÇÃO DE DADOS")>
    AlteracaoDeDados = 4
    <Description("LIQUIDAÇÃO NORMAL")>
    LiquidacaoNormal = 6
    <Description("CONFIRMA RECEBIMENTO DA INSTRUÇÃO DE ‘PROTESTO’")>
    ConfirmaRecebimento = 19
End Enum
