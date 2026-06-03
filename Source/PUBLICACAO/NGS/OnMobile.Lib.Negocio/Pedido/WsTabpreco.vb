Imports System.Configuration
Imports Newtonsoft.Json

<Serializable()>
Public Class ListWsTabPreco
    Inherits List(Of WsTabPreco)
    Public Sub New()
        Dim sql As String = String.Concat("SELECT Codigo_Id, Descricao FROM TabelaDePrecos WHERE Ativo = 1")

        Dim Banco As New AcessaBancoOnMobile
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "Contrato")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New WsTabPreco
            ws.TabPrecoCod = row("Codigo_Id")
            ws.TabPrecoDescricao = Left(row("Descricao"), 40)
            ws.TabPrecoValidade = DateTime.Now.AddMonths(6)
            Me.Add(ws)
        Next
    End Sub
End Class

<Serializable()>
Public Class WsTabPreco
    <JsonProperty("tabprecoCod")>
    Public Property TabPrecoCod As String

    <JsonProperty("tabprecoDescricao")>
    Public Property TabPrecoDescricao As String

    <JsonProperty("tabprecoValidade")>
    Public Property TabPrecoValidade As DateTime
End Class
