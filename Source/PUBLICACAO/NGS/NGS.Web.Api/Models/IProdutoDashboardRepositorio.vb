Imports NGS.Lib.Negocio
Imports NGS.Web.Api.Models

Public Interface IProdutoDashboardRepositorio

    'Consultar Banco de dados
    Function ConsultarBancoDeDados() As List(Of SelectListItem)

    'Empresas 
    Function ConsultarEmpresas() As List(Of SelectListItem)

    'Seguimentos 
    Function ConsultarSeguimentos() As List(Of SelectListItem)

    'Produtos
    Function TotalProdutoPorProduto(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard
    Function FaturamentoPorProduto(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
    Function DistribuicaoDeFaturamentoPorProduto(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
    Function DetalhamentoProdutos(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

    'Clientes
    Function TotalProdutoPorClientes(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard
    Function FaturamentoPorClientes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
    Function DetalhamentoClientes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

    'Representantes
    Function TotalProdutoPorRepresentantes(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard
    Function FaturamentoPorRepresentantes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
    Function DetalhamentoRepresentantes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

    'Cidades
    Function TotalProdutoPorCidades(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard
    Function FaturamentoPorCidades(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
    Function DetalhamentoCidades(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

    'Estados
    Function TotalProdutoPorEstados(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard
    Function FaturamentoPorEstados(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
    Function DetalhamentoEstados(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

    'Pedidos
    Function TotalProdutoPorPedidos(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard
    Function FaturamentoPorPedidos(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
    Function DetalhamentoPedidos(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

    'Vendas Diárias
    Function TotalVendasDiarias(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard
    Function EvolucaoVendasDiarias(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
    Function DetalhamentoVendasDiarias(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

    'Vendas Mensais
    Function TotalVendasMensais(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard
    Function EvolucaoVendasMensais(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
    Function DetalhamentoVendasMensais(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

    'Consolidada
    Function TotalConsolidado(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard
    Function EvolucaoConsolidado(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
    Function DetalhamentoConsolidado(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
    Function ComparacaoVolumeXFaturamentoConsolidado(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

End Interface
