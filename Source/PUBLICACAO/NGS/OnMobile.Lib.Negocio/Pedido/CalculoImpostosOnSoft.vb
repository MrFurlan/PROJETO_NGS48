<Serializable()>
Public Class ListCalculoImpostosOnSoft
    Inherits List(Of CalculoImpostosOnSoft)

    Public Sub New(ByVal codigoEstado As String, ByVal grupoMaisDe10Kg As Boolean)

        Dim sql As String =
        "SELECT [Estado_Id] " & vbCrLf &
        ",[AliquotaMVAEstado]" & vbCrLf &
        ",[FundoCombateAPobreza]" & vbCrLf &
        ",[EstadoDestino]" & vbCrLf &
        ",[AliquotaICMSEstadoDestino]" & vbCrLf &
        ",[EstadoOrigem]" & vbCrLf &
        ",[AliquotaICMSEstadoOrigem]" & vbCrLf &
        ",[AliquotaIPI]" & vbCrLf &
        ",[AjustarBaseICMSST]" & vbCrLf &
        ",[PossuiIPI]" & vbCrLf &
        ",[Indice]" & vbCrLf &
        ",[GrupoMaisDe10Kg]" & vbCrLf &
        "FROM [dbo].[CalculoImpostosOnSoft] " & vbCrLf &
        "WHERE Estado_Id = '" & codigoEstado & "' AND GrupoMaisDe10Kg = " & IIf(grupoMaisDe10Kg, 1, 0) & vbCrLf

        Dim Banco As New AcessaBancoOnMobile
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "TabPrecoProduto")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New CalculoImpostosOnSoft
            ws.Estado_Id = row("Codigo_Id")
            ws.AliquotaMVAEstado = row("AliquotaMVAEstado")
            'ws.TabCondicaoCod = row("TipoDePagamento_Id")
            ws.FundoCombateAPobreza = row("FundoCombateAPobreza")
            ws.EstadoDestino = row("EstadoDestino")
            ws.AliquotaICMSEstadoDestino = row("AliquotaICMSEstadoDestino")
            ws.EstadoOrigem = row("EstadoOrigem")
            ws.AliquotaICMSEstadoOrigem = row("AliquotaICMSEstadoOrigem")
            ws.AliquotaIPI = row("AliquotaIPI")
            ws.AjustarBaseICMSST = row("AjustarBaseICMSST")
            ws.PossuiIPI = row("PossuiIPI")
            ws.Indice = row("Indice")
            ws.GrupoMaisDe10Kg = row("GrupoMaisDe10Kg")
            Me.Add(ws)
        Next
    End Sub
End Class

<Serializable()>
Public Class CalculoImpostosOnSoft
    Public Sub New()

    End Sub
    Public Sub New(ByVal codigoEstado As String, ByVal grupoMaisDe10Kg As Boolean, ByVal comIPI As Boolean)

        Dim sql As String =
        "SELECT [Estado_Id] " & vbCrLf &
        ",[AliquotaMVAEstado]" & vbCrLf &
        ",[FundoCombateAPobreza]" & vbCrLf &
        ",[EstadoDestino]" & vbCrLf &
        ",[AliquotaICMSEstadoDestino]" & vbCrLf &
        ",[EstadoOrigem]" & vbCrLf &
        ",[AliquotaICMSEstadoOrigem]" & vbCrLf &
        ",[AliquotaIPI]" & vbCrLf &
        ",[AjustarBaseICMSST]" & vbCrLf &
        ",[PossuiIPI]" & vbCrLf &
        ",[Indice]" & vbCrLf &
        ",[GrupoMaisDe10Kg]" & vbCrLf &
        "FROM [dbo].[CalculoImpostosOnSoft] " & vbCrLf &
        "WHERE Estado_Id       = '" & codigoEstado & "'" & vbCrLf &
        "  AND GrupoMaisDe10Kg = " & IIf(grupoMaisDe10Kg, 1, 0) & vbCrLf
        If comIPI Then
            sql &= "  AND PossuiIPI       = 1" & vbCrLf
        Else
            sql &= "  AND PossuiIPI       = 0" & vbCrLf
        End If

        Dim Banco As New AcessaBancoOnMobile
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "CalculoImpostosOnSoft")

        For Each row As DataRow In ds.Tables(0).Rows
            Me.Estado_Id = row("Estado_Id")
            Me.AliquotaMVAEstado = row("AliquotaMVAEstado")
            Me.FundoCombateAPobreza = row("FundoCombateAPobreza")
            Me.EstadoDestino = row("EstadoDestino")
            Me.AliquotaICMSEstadoDestino = row("AliquotaICMSEstadoDestino")
            Me.EstadoOrigem = row("EstadoOrigem")
            Me.AliquotaICMSEstadoOrigem = row("AliquotaICMSEstadoOrigem")
            Me.AliquotaIPI = row("AliquotaIPI")
            Me.AjustarBaseICMSST = row("AjustarBaseICMSST")
            Me.PossuiIPI = row("PossuiIPI")
            Me.Indice = row("Indice")
            Me.GrupoMaisDe10Kg = row("GrupoMaisDe10Kg")
        Next
    End Sub
    Public Property Estado_Id As String
    Public Property AliquotaMVAEstado As Decimal
    Public Property FundoCombateAPobreza As Decimal
    Public Property EstadoDestino As String
    Public Property AliquotaICMSEstadoDestino As Decimal
    Public Property EstadoOrigem As String
    Public Property AliquotaICMSEstadoOrigem As Decimal
    Public Property AliquotaIPI As Decimal
    Public Property AjustarBaseICMSST As Boolean
    Public Property PossuiIPI As Boolean
    Public Property Indice As Double
    Public Property GrupoMaisDe10Kg As Boolean
End Class
