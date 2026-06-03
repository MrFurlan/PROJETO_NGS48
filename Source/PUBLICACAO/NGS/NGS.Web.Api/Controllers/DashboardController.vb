Imports System.Security.Claims
Imports System.Text
Imports System.Web.Http
Imports System.Web
Imports System.Net.Http
Imports System.Net
Imports NGS.Lib.Negocio

<Authorize>
Public Class DashboardController
    Inherits System.Web.Mvc.Controller

    Private Shared ReadOnly repository As IProdutoDashboardRepositorio = New ProdutoDashboardRepositorio()
    Private Shared ReadOnly _repo As New ProdutoDashboardRepositorio()

    'Home
    Function Index() As ActionResult

        ViewData("Title") = "Dashboard - Home"

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard As New ProdutoDashboard

        dashboard.Filtro = filtro

        Return View(dashboard)
    End Function

    'Produtos
    Function Produtos() As ActionResult

        ViewData("Title") = "Dashboard - Produtos"

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.TotalProdutoPorProduto(filtro)

        dashboard.Filtro = filtro

        Return View(dashboard)
    End Function

    <HttpGet>
    Public Function FaturamentoPorProduto() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.FaturamentoPorProduto(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_FaturamentoPorProduto.vbhtml", dashboard)
    End Function

    <HttpGet>
    Public Function DetalhamentoProdutos() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.DetalhamentoProdutos(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_DetalhamentoProdutos.vbhtml", dashboard)
    End Function

    <HttpGet>
    Public Function DistribuicaoDeFaturamentoPorProduto() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.DistribuicaoDeFaturamentoPorProduto(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_DistribuicaoDeFaturamentoPorProduto.vbhtml", dashboard)
    End Function

    'Clientes
    Function Clientes() As ActionResult

        ViewData("Title") = "Dashboard - Clientes"

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.TotalProdutoPorClientes(filtro)

        dashboard.Filtro = filtro

        Return View(dashboard)
    End Function

    <HttpGet>
    Public Function FaturamentoPorClientes() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.FaturamentoPorClientes(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_FaturamentoPorClientes.vbhtml", dashboard)
    End Function

    <HttpGet>
    Public Function DetalhamentoClientes() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.DetalhamentoClientes(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_DetalhamentoClientes.vbhtml", dashboard)
    End Function

    'Representantes
    Function Representantes() As ActionResult
        ViewData("Title") = "Dashboard - Representantes"

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.TotalProdutoPorRepresentantes(filtro)

        dashboard.Filtro = filtro

        Return View(dashboard)
    End Function

    <HttpGet>
    Public Function FaturamentoPorRepresentantes() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.FaturamentoPorRepresentantes(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_FaturamentoPorRepresentantes.vbhtml", dashboard)
    End Function

    <HttpGet>
    Public Function DetalhamentoRepresentantes() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.DetalhamentoRepresentantes(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_DetalhamentoRepresentantes.vbhtml", dashboard)
    End Function

    'Cidades
    Function Cidades() As ActionResult

        ViewData("Title") = "Dashboard - Cidades"

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.TotalProdutoPorCidades(filtro)

        dashboard.Filtro = filtro

        Return View(dashboard)
    End Function

    <HttpGet>
    Public Function FaturamentoPorCidades() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.FaturamentoPorCidades(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_FaturamentoPorCidades.vbhtml", dashboard)
    End Function

    <HttpGet>
    Public Function DetalhamentoCidades() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.DetalhamentoCidades(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_DetalhamentoCidades.vbhtml", dashboard)
    End Function

    'Estados
    Function Estados() As ActionResult

        ViewData("Title") = "Dashboard - Estados"

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.TotalProdutoPorEstados(filtro)

        dashboard.Filtro = filtro

        Return View(dashboard)
    End Function

    <HttpGet>
    Public Function FaturamentoPorEstados() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.FaturamentoPorEstados(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_FaturamentoPorEstados.vbhtml", dashboard)
    End Function

    <HttpGet>
    Public Function DetalhamentoEstados() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.DetalhamentoEstados(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_DetalhamentoEstados.vbhtml", dashboard)
    End Function

    'Pedidos
    Function Pedidos() As ActionResult

        ViewData("Title") = "Dashboard - Pedidos"

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.TotalProdutoPorPedidos(filtro)

        dashboard.Filtro = filtro

        Return View(dashboard)
    End Function

    <HttpGet>
    Public Function FaturamentoPorPedidos() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.FaturamentoPorPedidos(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_FaturamentoPorPedidos.vbhtml", dashboard)
    End Function

    <HttpGet>
    Public Function DetalhamentoPedidos() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.DetalhamentoPedidos(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_DetalhamentoPedidos.vbhtml", dashboard)
    End Function

    'Vendas Diárias
    Function VendasDiarias() As ActionResult

        ViewData("Title") = "Dashboard - VendasDiarias"

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.TotalVendasDiarias(filtro)

        dashboard.Filtro = filtro

        Return View(dashboard)
    End Function

    <HttpGet>
    Public Function EvolucaoVendasDiarias() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.EvolucaoVendasDiarias(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_EvolucaoVendasDiarias.vbhtml", dashboard)
    End Function

    <HttpGet>
    Public Function DetalhamentoVendasDiarias() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.DetalhamentoVendasDiarias(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_DetalhamentoVendasDiarias.vbhtml", dashboard)
    End Function

    'Vendas Mensal
    Function VendasMensal() As ActionResult

        ViewData("Title") = "Dashboard - VendasMensal"

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.TotalVendasMensais(filtro)

        dashboard.Filtro = filtro

        Return View(dashboard)
    End Function

    <HttpGet>
    Public Function EvolucaoVendasMensais() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.EvolucaoVendasMensais(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_EvolucaoVendasMensais.vbhtml", dashboard)
    End Function

    <HttpGet>
    Public Function DetalhamentoVendasMensais() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.DetalhamentoVendasMensais(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_DetalhamentoVendasMensais.vbhtml", dashboard)
    End Function

    'Consolidado
    Function Consolidado() As ActionResult

        ViewData("Title") = "Dashboard - Consolidado"

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.TotalConsolidado(filtro)

        dashboard.Filtro = filtro

        Return View(dashboard)
    End Function

    <HttpGet>
    Public Function EvolucaoConsolidado() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.EvolucaoConsolidado(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_EvolucaoConsolidado.vbhtml", dashboard)
    End Function

    <HttpGet>
    Public Function DetalhamentoConsolidado() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.DetalhamentoConsolidado(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_DetalhamentoConsolidado.vbhtml", dashboard)
    End Function

    <HttpGet>
    Public Function ComparacaoVolumeXFaturamentoConsolidado() As PartialViewResult

        Dim filtro = SessionHelper.ObterProdutoDashboardFiltro(Request, Session, repository)
        Dim dashboard = repository.ComparacaoVolumeXFaturamentoConsolidado(filtro)

        Return PartialView("~/Views/Dashboard/Partial/_ComparacaoVolumeXFaturamentoConsolidado.vbhtml", dashboard)
    End Function

End Class

