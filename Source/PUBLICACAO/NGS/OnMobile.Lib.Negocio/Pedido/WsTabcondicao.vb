Imports System.Configuration
Imports Newtonsoft.Json

<Serializable()>
Public Class ListWsTabCondicao
    Inherits List(Of WsTabCondicao)
    Public Sub New(Optional ByVal CarregarDados As Boolean = False)
        Dim sql As String = String.Concat("
            SELECT 
	            tp.TipoDePagamento_Id, 
	            Descricao 
            FROM [dbo].[TiposDePagamentos] tp
            INNER JOIN TiposDePagamentosXPagamentos tpp ON tp.TipoDePagamento_Id = tpp.TipoDePagamento_Id
            WHERE (TPagSefaz <> 0 OR TPagSefaz <> null) 
            group by tp.TipoDePagamento_Id, Descricao
            ORDER BY tp.TipoDePagamento_Id asc")

        Me.Add(New WsTabCondicao With {.TabCondicaoCod = "001", .TabCondicaoDescricao = "PADRAO"})
    End Sub
End Class

<Serializable()>
Public Class WsTabCondicao
    <JsonProperty("tabcondicaoCod")>
    Public Property TabCondicaoCod As String

    <JsonProperty("tabcondicaoDescricao")>
    Public Property TabCondicaoDescricao As String
End Class
