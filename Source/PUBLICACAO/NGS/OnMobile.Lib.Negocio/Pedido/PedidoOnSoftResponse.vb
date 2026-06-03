Imports Newtonsoft.Json

<Serializable()>
Public Class PedidoOnSoftResponse
    <JsonProperty("statusCode")>
    Public Property statusCode As Integer

    <JsonProperty("pageSize")>
    Public Property pageSize As Integer

    <JsonProperty("currentPage")>
    Public Property currentPage As Integer

    <JsonProperty("startRow")>
    Public Property startRow As Integer

    <JsonProperty("totalItems")>
    Public Property totalItems As Integer

    <JsonProperty("totalPages")>
    Public Property totalPages As Integer

    <JsonProperty("items")>
    Public Property Items As List(Of WsPedido)
End Class

