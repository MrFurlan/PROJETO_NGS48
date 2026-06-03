Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Interface IBasePage

    Sub Carregar(ByVal str As String)
    Sub Carregar(ByVal str As String, ByVal dec As Decimal)
    Sub Carregar(ByVal obj As IBaseEntity)

End Interface
