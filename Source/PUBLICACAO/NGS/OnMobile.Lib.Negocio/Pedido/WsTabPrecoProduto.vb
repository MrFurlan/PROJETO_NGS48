Imports System.Configuration
Imports Newtonsoft.Json

<Serializable()>
Public Class ListWsTabPrecoProduto
    Inherits List(Of WsTabPrecoProduto)
    Public Sub New()
        Dim sql As String = String.Concat("
            SELECT 
	            tp.Codigo_Id,
	            p.Produto_Id,
	            pp.Valor,
				case
					when ISNULL(pp.MargemMaior, 0) = 0
						then pp.Valor
						else pp.MargemMaior
					end AS MargemMaior,
				case
					when ISNULL(pp.MargemMenor, 0) = 0
						then pp.Valor
						else pp.MargemMenor
					end AS MargemMenor
            FROM TabelaDePrecos tp
            INNER JOIN ProdutosXPrecos pp on tp.Codigo_Id = pp.Tabela_Id
            INNER JOIN Produtos p on pp.Produto_Id = p.Produto_Id
            inner join (
                select Produto_Id, max(Data_Id) as MaxDate
                from ProdutosXPrecos
                group by Produto_Id, Tabela_Id, Cliente_Id
            ) tm on pp.Produto_Id = tm.Produto_Id and pp.Data_Id = tm.MaxDate
            WHERE left(p.Grupo, 1) = '8'
            GROUP BY p.Produto_Id,
	            tp.Codigo_Id,
	            p.Produto_Id,
	            pp.Valor,
	            MargemMaior,
	            MargemMenor
            order by tp.Codigo_Id asc")

        Dim Banco As New AcessaBancoOnMobile
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "TabPrecoProduto")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New WsTabPrecoProduto
            ws.TabPrecoCod = row("Codigo_Id")
            ws.ProdutoCod = row("Produto_Id")
            'ws.TabCondicaoCod = row("TipoDePagamento_Id")
            ws.ProdutoStatus = "N"
            ws.TabPrecoVr = row("Valor")
            ws.TabPrecoVrmin = row("MargemMenor")
            ws.TabPrecoVrmax = row("MargemMaior")
            Me.Add(ws)
        Next
    End Sub
End Class

<Serializable()>
Public Class WsTabPrecoProduto
    <JsonProperty("tabprecoCod")>
    Public Property TabPrecoCod As String

    <JsonProperty("produtoCod")>
    Public Property ProdutoCod As String

    <JsonProperty("tabcondicaoCod")>
    Public Property TabCondicaoCod As String

    <JsonProperty("produtoStatus")>
    Public Property ProdutoStatus As String

    <JsonProperty("tabprecoVr")>
    Public Property TabPrecoVr As Double

    <JsonProperty("tabprecoVr2")>
    Public Property TabPrecoVr2 As Double

    <JsonProperty("tabprecoVrmin")>
    Public Property TabPrecoVrmin As Double

    <JsonProperty("tabprecoVrmax")>
    Public Property TabPrecoVrmax As Double

    <JsonProperty("tabprecoFrete")>
    Public Property TabPrecoFrete As Double

    <JsonProperty("tabprecoPackVr")>
    Public Property TabPrecoPackVr As Double

    <JsonProperty("tabprecoPackVr2")>
    Public Property TabPrecoPackVr2 As Double

    <JsonProperty("tabprecoPackVrmin")>
    Public Property TabPrecoPackVrmin As Double

    <JsonProperty("tabprecoPackVrmax")>
    Public Property TabPrecoPackVrmax As Double

    <JsonProperty("tabprecoPackFrete")>
    Public Property TabPrecoPackFrete As Double

    <JsonProperty("tabprecoDesconto")>
    Public Property TabPrecoDesconto As Double

    <JsonProperty("tabprecoPackDesconto")>
    Public Property TabPrecoPackDesconto As Double
End Class
