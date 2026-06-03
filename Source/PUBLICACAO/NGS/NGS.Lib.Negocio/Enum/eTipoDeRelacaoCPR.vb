Imports System.ComponentModel
Imports NGS.Lib.Uteis

Public Enum eTipoDeRelacaoCPR

    <Description("Devedor")> _
    Devedor = 1

    <Description("Fiador / Devedor Solidário")> _
    FiadorDevedorSolidario = 2

    <Description("Avalista")> _
    Avalista = 3

    <Description("Anuente")> _
    Anuente = 4
End Enum
