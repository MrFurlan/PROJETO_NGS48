Imports Newtonsoft.Json

<Serializable()>
Public Class ListWsPedidoIntegracao
    Inherits List(Of WsPedidoIntegracao)
    Public Sub New()
        Dim sql As String = "Select * from Pedidos"
        Dim Banco As New AcessaBancoOnMobile
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "Contrato")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New WsPedidoIntegracao

            Me.Add(ws)
        Next
    End Sub
End Class

<Serializable()>
Public Class WsPedidoIntegracao
    <JsonProperty("vendedorCod")>
    Public Property VendedorCod As String

    <JsonProperty("pedidoNum")>
    Public Property PedidoNum As String

    <JsonProperty("pedidoNumpedcli")>
    Public Property PedidoNumPedCli As String

    <JsonProperty("pedidoDtenvio")>
    Public Property PedidoDtEnvio As DateTime
End Class
