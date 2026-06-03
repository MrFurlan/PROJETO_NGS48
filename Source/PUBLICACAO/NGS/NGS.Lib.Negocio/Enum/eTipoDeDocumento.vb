Imports System.ComponentModel
Imports NGS.Lib.Uteis

Public Enum eTipoDeDocumento

    <Description("Nota Fiscal")> _
    Nota = 1

    <Description("Conhecimento de Transporte")> _
    CTRC = 2

    <Description("Recibo de Frete")> _
    ReciboDeFrete = 3

    <Description("Contrato de Frete")> _
    ContratoDeFrete = 4

    <Description("Cupom Fiscal")> _
    CupomFiscal = 5

    <Description("Recibo")> _
    Recibo = 6

    <Description("Fatura")> _
    Fatura = 7

    <Description("Conhecimento de Transporte s/ Nota Fiscal")> _
    CTRC_SEM_NF = 8

    <Description("Nota Fiscal Municipal")> _
    NFMunicipal = 9

    <Description("Estadia")> _
    Estadia = 10

    <Description("Anulação")> _
    Anulacao = 11

    <Description("MDF-e")> _
    ManifestoEletronico = 12

    <Description("Recibo De Pagamento a Autônomo")> _
    RPA = 13

    <Description("Complemento de Frete")> _
    ComplementoDeFrete = 14

    <Description("Nota de Produtor")> _
    NotaDeProdutor = 15

    <Description("Energia Eletrica")> _
    EnergiaEletrica = 16

    <Description("Nota de Débito")> _
    NotaDeDebito = 20

    <Description("Serviço de Comunicação")> _
    ServicoComunicacao = 21

    <Description("Serviço de Telecomunicação")> _
    ServicoTelecomunicacao = 22

    <Description("Conhecimento de Transporte Eletrônico")> _
    CT_E = 57

    <Description("Conhecimento de Transporte Eletrônico - Tomador de Serviço")> _
    CT_E_TOM = 58

    <Description("Cupom Fiscal Eletrônico")> _
    CupomFiscalEletronico = 59

    <Description("NOTA FISCAL ELETRONICA AO CONSUMIDOR FINAL NFC-e")> _
    NFC_e = 65

    <Description("NOTA FISCAL DE ENERGIA ELETRICA ELETRONICA NF3-e")>
    NF3_e = 66

    <Description("Conhecimento de Transporte Eletrônico Outros Serviços")> _
    CT_E_OUT = 67

End Enum