Imports Newtonsoft.Json

<Serializable()>
Public Class ListWsPedidoObs
    Inherits List(Of WsPedidoObs)
    Public Sub New(Optional ByVal CarregarDados As Boolean = False)
        Dim sql As String = "Select * from Clientes"
        Dim Banco As New AcessaBancoOnMobile
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "Contrato")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New WsPedidoObs

            Me.Add(ws)
        Next
    End Sub
End Class

<Serializable()>
Public Class WsPedidoObs
    <JsonProperty("vendedorCod    ")>
    Public Property VendedorCod As String

    <JsonProperty("pedidoNum  ")>
    Public Property PedidoNum As String

    <JsonProperty("obsPadraoCod   ")>
    Public Property ObsPadraoCod As String

    <JsonProperty("pedidoObsCompl")>
    Public Property PedidoObsCompl As String
End Class
