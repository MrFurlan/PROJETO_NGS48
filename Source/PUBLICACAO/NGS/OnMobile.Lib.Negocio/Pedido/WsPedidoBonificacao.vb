Imports Newtonsoft.Json

<Serializable()>
Public Class WsPedidoBonificacao
    <JsonProperty("vendedorCod    ")>
    Public Property VendedorCod As String

    <JsonProperty("pedidoNum  ")>
    Public Property PedidoNum As String

    <JsonProperty("bonificacaoNumitem")>
    Public Property BonificacaoNumitem As String

    <JsonProperty("produtoCod")>
    Public Property ProdutoCod As String

    <JsonProperty("bonificacaoQtde")>
    Public Property BonificacaoQtde As Integer

    <JsonProperty("bonificacaoVrunit")>
    Public Property BonificacaoVrunit As Double
End Class
