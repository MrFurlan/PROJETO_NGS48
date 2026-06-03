Imports System.Configuration
Imports Newtonsoft.Json
Imports NGS.Lib.Negocio

<Serializable()>
Public Class ListWsCondicao
    Inherits List(Of WsCondicao)

    Public Sub New()
        Dim sql As String = String.Concat("
                SELECT 
	                tp.TipoDePagamento_Id, 
	                tp.Descricao AS TipoDePagamento,
	                p.Pagamento_Id,
	                e.Estado_Id,
	                p.Descricao
                FROM [dbo].[TiposDePagamentos] tp
                INNER JOIN TiposDePagamentosXPagamentos tpxp 
	                ON tp.TipoDePagamento_Id = tpxp.TipoDePagamento_Id
                INNER JOIN Pagamentos p 
	                ON tpxp.Pagamento_Id = p.Pagamento_Id 
                INNER JOIN [ProdutosXPrecos] pxp 
	                ON tpxp.Cliente_Id = pxp.Cliente_Id
	                AND tpxp.EndCliente_Id = pxp.EndCliente_Id
                INNER JOIN Clientes c 
	                ON pxp.Cliente_Id = c.Cliente_Id
	                AND pxp.EndCliente_Id = c.Endereco_Id
                INNER JOIN Estados e 
	                ON c.Estado = e.Estado_Id
                WHERE (TPagSefaz <> 0 OR TPagSefaz <> null)
                GROUP BY 
	                tp.TipoDePagamento_Id, 
	                p.Pagamento_Id,
	                e.Estado_Id,
	                p.Descricao,
	                tp.Descricao
                ORDER BY tp.TipoDePagamento_Id asc")

        Dim Banco As New AcessaBancoOnMobile
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "WsCondicao")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New WsCondicao
            ws.TabCondicaoCod = "001" 'TODO: verificar o que é? e como fazer a ligação com as outras tabelas.
            ws.CondicaoCod = String.Format("{0}|{1}", row("TipoDePagamento_Id"), row("Pagamento_Id"))
            ws.TabPrecoCod = row("Estado_Id")
            Dim nome = String.Format("{0} {1}", row("TipoDePagamento"), row("Descricao"))
            ws.CondicaoDescricao = Left(nome, 40)
            ws.CondicaoFator = 1
            'ws.CondicaoRnMinUe = row("")
            'ws.CondicaoRnMinVr = row("")
            'ws.CondicaoRnMinModo = row("")
            ws.CondicaoRecNo = (Me.Count + 1)
            Me.Add(ws)
        Next
    End Sub
End Class

<Serializable()>
Public Class WsCondicao
    <JsonProperty("tabcondicaoCod")>
    Public Property TabCondicaoCod As String

    <JsonProperty("condicaoCod")>
    Public Property CondicaoCod As String

    <JsonProperty("tabprecoCod")>
    Public Property TabPrecoCod As String

    <JsonProperty("condicaoDescricao")>
    Public Property CondicaoDescricao As String

    <JsonProperty("condicaoFator")>
    Public Property CondicaoFator As Double

    <JsonProperty("condicaoRnMinUe")>
    Public Property CondicaoRnMinUe As Double

    <JsonProperty("condicaoRnMinVr")>
    Public Property CondicaoRnMinVr As Double

    <JsonProperty("condicaoRnMinModo")>
    Public Property CondicaoRnMinModo As String

    <JsonProperty("condicaoRecNo")>
    Public Property CondicaoRecNo As Integer

    Public Function PegarCondicao(ByVal Codigo As Integer) As CondicaoPagamento
        Dim c As New CondicaoPagamento
        Dim Banco As New AcessaBancoOnMobile
        Try
            Dim sql As String
            sql = "SELECT Pagamento_Id, Descricao, Parcelas, isnull(AVista,0) as Avista, ISNULL(Antecipado,0) AS Antecipado " &
                  "  FROM Pagamentos " &
                  " WHERE Pagamento_Id = " & Codigo.ToString() & " " &
                  " ORDER BY Pagamento_Id"

            Dim dsCondicoes As DataSet = Banco.ConsultaDataSet(sql, "Produtos")

            If dsCondicoes.Tables(0).Rows.Count > 0 Then
                Dim row As DataRow = dsCondicoes.Tables(0).Rows(0)
                c.Codigo = row("Pagamento_Id")
                c.Descricao = row("Descricao")
                c.Parcelas = row("Parcelas")
                c.AVista = row("AVista")
                c.Antecipado = row("Antecipado")
            End If

            Return c
        Catch ex As Exception

        Finally
            Banco = Nothing
        End Try
    End Function

End Class
