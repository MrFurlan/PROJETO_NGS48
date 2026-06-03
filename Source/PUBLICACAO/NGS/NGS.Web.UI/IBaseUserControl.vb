Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Interface IBaseUserControl

    Sub Carregar(ByVal str As String)
    Sub Carregar(ByVal str As String, ByVal dec As Decimal)
    Sub Carregar(ByVal obj As IBaseEntity)
    Sub Carregar(ByVal parameters As Dictionary(Of String, Object))

End Interface
