Imports NGS.Lib.Uteis
Imports System.ComponentModel

Public Enum eTipoCampo

    ''' <summary>
    ''' Campo do tipo Data.
    ''' </summary>
    ''' <remarks></remarks>
    <Description("Data")> _
    Data

    ''' <summary>
    ''' Campo do tipo Numérico inteiro.
    ''' </summary>
    ''' <remarks></remarks>
    <Description("Numérico")> _
    Numerico

    ''' <summary>
    ''' Campo do tipo Numérico com casas decimais(Decimal, Double, Float).
    ''' </summary>
    ''' <remarks></remarks>
    <Description("ValorComTotal")> _
    ValorComTotalizador

    ''' <summary>
    ''' Campo do tipo Numérico com casas decimais(Decimal, Double, Float). Não gera totalizadores no excel
    ''' </summary>
    ''' <remarks></remarks>
    <Description("ValorSemTotal")> _
    ValorSemTotalizador

    ''' <summary>
    ''' Campo do tipo Numérico com casas decimais(Decimal, Double, Float). Não gera totalizadores no excel
    ''' </summary>
    ''' <remarks></remarks>
    <Description("ValorSemTotal3Decimais")> _
    ValorSemTotalizadorCom3CasasDecimais

    ''' <summary>
    ''' Campo do tipo Numérico com casas decimais(Decimal, Double, Float). Não gera totalizadores no excel
    ''' </summary>
    ''' <remarks></remarks>
    <Description("ValorSemTotal4Decimais")> _
    ValorSemTotalizadorCom4CasasDecimais

    ''' <summary>
    ''' Campo do tipo String com a fórmula Excel.
    ''' </summary>
    ''' <remarks></remarks>
    <Description("Fórmula")> _
   Formula

End Enum
