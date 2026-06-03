Imports NGS.Lib.Uteis

<Serializable()> _
Public Class BaseEntity(Of t As Class)
    Implements IBaseEntity

    Public Overrides Function ToString() As String
        Return [GetType]().Name
    End Function

End Class