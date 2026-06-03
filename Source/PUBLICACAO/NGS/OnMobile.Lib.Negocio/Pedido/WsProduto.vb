Imports System.Configuration
Imports Newtonsoft.Json

<Serializable()>
Public Class ListWsProduto
    Inherits List(Of WsProduto)
    Public Sub New(Optional ByVal CarregarDados As Boolean = False)
        Dim sql As String = String.Concat(
            "SELECT 
	            p.Produto_Id,
	            p.Grupo,
	            p.Nome,
	            p.Unidade,
	            p.Descricao
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
	            p.Grupo,
	            p.Nome,
	            p.Unidade,
	            p.Descricao")

        Dim Banco As New AcessaBancoOnMobile
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "produtos")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New WsProduto

            ws.ProdutoCod = row("Produto_Id")
            ws.GrupoProdutoCod = row("Grupo")
            'ws.SubgrupoProdutoCod = row("")
            ws.ProdutoDescricao = row("Nome")
            ws.ProdutoUnidade = row("Unidade")
            'ws.ProdutoUe = row("")
            ws.ProdutoObs = row("Descricao")
            'ws.ProdutoPackUnidade = row("")
            'ws.ProdutoPackUe = row("")
            'ws.ProdutoPackQtde = row("")

            Me.Add(ws)
        Next
    End Sub
End Class
Public Class WsProduto
    <JsonProperty("produtoCod")>
    Public Property ProdutoCod As String

    <JsonProperty("grupoprodutoCod")>
    Public Property GrupoProdutoCod As String

    <JsonProperty("subgrupoProdutoCod")>
    Public Property SubGrupoProdutoCod As String

    <JsonProperty("produtoDescricao")>
    Public Property ProdutoDescricao As String

    <JsonProperty("produtoUnidade")>
    Public Property ProdutoUnidade As String

    <JsonProperty("produtoUe")>
    Public Property ProdutoUe As Double

    <JsonProperty("produtoObs")>
    Public Property ProdutoObs As String

    <JsonProperty("produtoPackUnidade")>
    Public Property ProdutoPackUnidade As String

    <JsonProperty("produtoPackUe")>
    Public Property ProdutoPackUe As Double

    <JsonProperty("produtoPackQtde")>
    Public Property ProdutoPackQtde As Integer

    Public Sub New()

    End Sub
    Public Sub New(ByVal CodigoProduto As String)
        Dim objBanco As New AcessaBancoOnMobile()


        Dim strSQL As String = "SELECT Produto_Id, Grupo, Unidade, Etapa, Situacao, Embalagem, NCM, Nome, isnull(CodBarras,'') AS CodBarras, Descricao, " & vbCrLf &
                               "       DescricaoMapa, PesoQuantidade, EstoqueMinimo, UPPER(isnull(Agrupar,'N')) AS Agrupar, QuantidadeNaCaixa, " & vbCrLf &
                               "       Qualidade, PesoDoItem, IPI, CarteiraDeCompras, " & vbCrLf &
                               "       CarteiraDeVendas, ICMS, TipoDoItem, isnull(CodigoDoGenero,0) AS CodigoDoGenero, CodigoEX, CodigoDoServico, " & vbCrLf &
                               "       UPPER(isnull(controlarEstoque,'N')) as ControlarEstoque, UPPER(isnull(ControlarLote,'')) as ControlarLote, " & vbCrLf &
                               "       UPPER(isnull(ControlarEmbalagem,'N')) as ControlarEmbalagem, UPPER(isnull(Fitossanitario,'N')) as Fitossanitario, " & vbCrLf &
                               "       UPPER(isnull(ProdutoIndea,'')) as ProdutoIndea, isnull(Marca,0) as Marca, isnull(ControlarPecas,0) as ControlarPecas, " & vbCrLf &
                               "       isnull(ControlarPrecoDePauta,0) as ControlarPrecoDePauta, isnull(SubCodigoDoGenero,0) AS SubCodigoDoGenero, " & vbCrLf &
                               "       isnull(ControlarRomaneio,0) as ControlarRomaneio, isnull(ControlarPesagem,0) as ControlarPesagem, " & vbCrLf &
                               "       isnull(ControlarDecimais,0) as ControlarDecimais, ISNULL(Cnae,'') Cnae, isnull(Almoxarifado,0) as Almoxarifado, isnull(PrecoDoProduto,0) as PrecoDoProduto, " & vbCrLf &
                               "       isnull(UsuarioInclusao,'') as UsuarioInclusao, isnull(UsuarioInclusaoData,GetDate()) as UsuarioInclusaoData," & vbCrLf &
                               "       isnull(UsuarioAlteracao,'') as UsuarioAlteracao, isnull(UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
                               "       isnull(ControlarNumeroDoLote,0) as ControlarNumeroDoLote, isnull(CodigoEstadoFisico,0) as CodigoEstadoFisico, " & vbCrLf &
                               "       isnull(InfaDProd,'') as InfaDProd, isnull(AutorizacaoDeRetirada,0) AS AutorizacaoDeRetirada, isnull(RegMinAgr,'') AS RegMinAgr, " & vbCrLf &
                               "       isnull(CodigoProdutoTerceiro,'') as CodigoProdutoTerceiro" & vbCrLf &
                               "  FROM Produtos " & vbCrLf &
                               " WHERE Produto_Id = '" & CodigoProduto & "' "

        Dim dsProdutos As DataSet = objBanco.ConsultaDataSet(strSQL, "Produtos")

        If dsProdutos.Tables(0).Rows.Count > 0 Then
            Dim drProduto As DataRow = dsProdutos.Tables(0).Rows(0)

            Me.ProdutoCod = drProduto("Produto_Id").ToString()
            Me.ProdutoUnidade = drProduto("Unidade").ToString()
            Me.ProdutoDescricao = drProduto("Nome").ToString()
        End If

    End Sub
End Class
