@ModelType  ProdutoDashboard
@Code
    Dim rawPath As String = Request.Url.AbsolutePath ' /ngsApi/controller/action
    Dim basePath As String = Url.Content("~/")       ' /ngsApi/
    Dim relativePath As String = rawPath.Substring(basePath.Length) ' controller/action
    Dim parts As String() = relativePath.Split(New Char() {"/"c}, StringSplitOptions.RemoveEmptyEntries)
    Dim breadcrumbPath As String = ""
End Code

<!DOCTYPE html>
<html>
<head>

    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width" />

    <link id="favicon" rel="shortcut icon" href="@Url.Content("~/admin-lte/assets/img/favicon.png")" />
    <title>@ViewBag.Title</title>

    <!--begin::Fonts-->
    <link rel="stylesheet" href="@Url.Content("~/admin-lte/css/index.css")" />
    <link rel="stylesheet" href="@Url.Content("~/admin-lte/css/style.css")" />
    <link rel="stylesheet" href="@Url.Content("~/admin-lte/css/overlayscrollbars.min.css")" />
    <link rel="stylesheet" href="@Url.Content("~/admin-lte/css/adminlte.css")" />
    <link rel="stylesheet" href="@Url.Content("~/admin-lte/css/apexcharts.css")" />
    <link rel="stylesheet" href="@Url.Content("~/admin-lte/css/bootstrap-icons.min.css")" />

    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/multiselect-dropdown@1.0.1/index.min.css" />


    @Scripts.Render("~/bundles/modernizr")
    @Scripts.Render("~/admin-lte/js")

    @RenderSection("Styles", required:=False)

    <style>
        #btnToggleFiltros {
            background-color: transparent !important;
            border: none !important;
            color: black !important;
            display: flex;
            justify-content: flex-end;
            align-items: center;
            gap: 6px;
            width: 100%;
            padding: 0.375rem 0.75rem; /* padding padrão de btn */
        }

            #btnToggleFiltros:hover, #btnToggleFiltros:focus {
                background-color: #e9ecef; /* leve efeito hover, opcional */
                color: black !important;
            }

        /* Ícone da seta preto */
        #iconeSeta {
            color: black;
            font-weight: bold;
        }

        /* Borda cinza no topo da div filtro, sem fundo */
        #filtroCollapse > div {
            border-top: 1px solid #ccc;
            background-color: transparent !important;
        }

        .dropdown-multiselect {
            position: relative;
            display: inline-block;
            outline: none;
            width: auto;
            min-width: 800px;
        }

        .select-box {
            border: 1px solid #ccc;
            padding: 6px 12px;
            border-radius: 4px;
            cursor: pointer;
            background-color: #fff;
            user-select: none;
            display: flex;
            align-items: center;
            gap: 8px;
            white-space: nowrap;
            transition: width 0.2s ease;
        }

            .select-box span {
                flex: 1;
                overflow: visible;
            }

        .options-box {
            position: absolute;
            background-color: #fff;
            border: 1px solid #ccc;
            max-height: 250px;
            overflow-y: auto;
            overflow-x: hidden;
            margin-top: 2px;
            padding: 8px;
            border-radius: 4px;
            z-index: 1000;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.2);
            white-space: nowrap;
        }

        .dropdown-wrapper {
            min-width: 220px;
            max-width: 500px;
            width: auto;
        }

    </style>

</head>

<body class="layout-fixed sidebar-expand-lg bg-body-tertiary">

    <div class="app-wrapper">
        <!--begin::Header-->
        <nav class="app-header navbar navbar-expand bg-body">
            <!--begin::Container-->
            <div class="container-fluid">
                <!--begin::Start Navbar Links-->
                <ul class="navbar-nav">
                    <li class="nav-item">
                        <a class="nav-link" data-lte-toggle="sidebar" href="#" role="button">
                            <i class="bi bi-list"></i>
                        </a>
                    </li>
                    <li class="nav-item d-none d-md-block"><a href="@Url.Action("Index", "Dashboard")" class="nav-link">Dashboard</a></li>
                    <li class="nav-item d-none d-md-block"><a href="#" class="nav-link">Sistema</a></li>
                </ul>
                <!--end::Start Navbar Links-->
                <!--begin::End Navbar Links-->
                <ul class="navbar-nav ms-auto">
                    <!--begin::Fullscreen Toggle-->
                    <li class="nav-item">
                        <a class="nav-link" href="#" data-lte-toggle="fullscreen">
                            <i data-lte-icon="maximize" class="bi bi-arrows-fullscreen"></i>
                            <i data-lte-icon="minimize" class="bi bi-fullscreen-exit" style="display: none"></i>
                        </a>
                    </li>
                    <!--end::Fullscreen Toggle-->
                    <!-- Logout separado -->
                    <li class="nav-item">
                        <a href="@Url.Action("Logout", "Conta")" title="Sair" class="nav-link d-flex align-items-center gap-1">
                            <i class="bi bi-box-arrow-right"></i>
                            <span class="d-none d-md-inline">Sair</span>
                        </a>
                    </li>
                </ul>
                <!--end::End Navbar Links-->
            </div>
            <!--end::Container-->
        </nav>
        <!--end::Header-->
        <!--begin::Sidebar-->
        <aside class="app-sidebar bg-body-secondary shadow" data-bs-theme="dark">
            <!--begin::Sidebar Brand-->
            <div class="sidebar-brand" style="background-color: white !important;">
                <!--begin::Brand Link-->
                <a href="#" class="brand-link">
                    <!--begin::Brand Image-->
                    <img src="@Url.Content("~/admin-lte/assets/img/ngs.png")"
                         alt="NGS Soluções"
                         class="brand-image" />
                    <!--end::Brand Image-->
                    <!--begin::Brand Text-->
                    <span class="brand-text fw-light">NGS</span>
                    <!--end::Brand Text-->
                </a>
                <!--end::Brand Link-->
            </div>
            <!--end::Sidebar Brand-->
            <!--begin::Sidebar Wrapper-->
            <div class="sidebar-wrapper">
                <nav class="mt-2">
                    <!--begin::Sidebar Menu-->
                    <ul class="nav sidebar-menu flex-column"
                        data-lte-toggle="treeview"
                        role="menu"
                        data-accordion="false">
                        <li class="nav-item menu-open">
                            <a href="#" class="nav-link @(If(Request.RawUrl.Contains("/Dashboard/"), "active", ""))">
                                <svg xmlns="http://www.w3.org/2000/svg" class="lucide lucide-speedometer nav-icon w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M12 12v2" /><path d="M10 10h.01" /><path d="M14 10h.01" /><path d="M4 20a10 10 0 1 1 16 0Z" /></svg>
                                <p>
                                    Dashboard
                                    <i class="nav-arrow bi bi-chevron-right"></i>
                                </p>
                            </a>
                            <ul class="nav nav-treeview">
                                <li class="nav-item">
                                    <a href="@Url.Action("Index", "Dashboard")" class="nav-link @(If(Request.RawUrl.Equals("/Dashboard") Or Request.RawUrl.Equals("/Dashboard/Index"), "active", ""))">
                                        <svg class="lucide lucide-home nav-icon w-5 h-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M3 9L12 2l9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z" /><polyline points="9 22 9 12 15 12 15 22" /></svg>
                                        <p>Home</p>
                                    </a>
                                </li>
                                <li class="nav-item">
                                    <a href="@Url.Action("Consolidado", "Dashboard")" class="nav-link @(If(Request.RawUrl.Contains("/Dashboard/Consolidado"), "active", ""))">
                                        <svg class="lucide lucide-trending-up nav-icon w-5 h-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><polyline points="22 7 13.5 15.5 8.5 10.5 2 17" /><polyline points="16 7 22 7 22 13" /></svg>
                                        <p>Consolidado</p>
                                    </a>
                                </li>
                                <li class="nav-item">
                                    <a href="@Url.Action("Produtos", "Dashboard")" class="nav-link @(If(Request.RawUrl.Contains("/Dashboard/Produtos"), "active", ""))">
                                        <svg class="lucide lucide-chart-pie nav-icon w-5 h-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 0-9 9" /><path d="M21 12A9 9 0 0 0 12 3v9z" /></svg>
                                        <p>Produtos</p>
                                    </a>
                                </li>
                                <li class="nav-item">
                                    <a href="@Url.Action("Clientes", "Dashboard")" class="nav-link @(If(Request.RawUrl.Contains("/Dashboard/Clientes"), "active", ""))">
                                        <svg class="lucide lucide-chart-bar nav-icon w-5 h-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18" /><path d="M7 16h8" /><path d="M7 11h12" /><path d="M7 6h3" /></svg>
                                        <p>Clientes</p>
                                    </a>
                                </li>
                                <li class="nav-item">
                                    <a href="@Url.Action("Representantes", "Dashboard")" class="nav-link @(If(Request.RawUrl.Contains("/Dashboard/Representantes"), "active", ""))">
                                        <svg class="lucide lucide-chart-bar nav-icon w-5 h-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18" /><path d="M7 16h8" /><path d="M7 11h12" /><path d="M7 6h3" /></svg>
                                        <p>Representantes</p>
                                    </a>
                                </li>
                                <li class="nav-item">
                                    <a href="@Url.Action("Cidades", "Dashboard")" class="nav-link @(If(Request.RawUrl.Contains("/Dashboard/Cidades"), "active", ""))">
                                        <svg class="lucide lucide-map nav-icon w-5 h-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M9 3v15" /><path d="M15 5v15" /><path d="M3.3 7l7.7 4.7a2 2 0 0 0 2 0l7.7-4.7" /><path d="M21 4.6v12.8a1 1 0 0 1-.6.9l-4.6 2.3a2 2 0 0 1-1.8 0l-4.2-2.1a2 2 0 0 0-1.8 0l-3.7 1.8A1 1 0 0 1 3 19.4V6.6a1 1 0 0 1 .6-.9l4.6-2.3a2 2 0 0 1 1.8 0l4.2 2.1a2 2 0 0 0 1.8 0l3.7-1.8A1 1 0 0 1 21 4.6Z" /></svg>
                                        <p>Cidades</p>
                                    </a>
                                </li>
                                <li class="nav-item">
                                    <a href="@Url.Action("Estados", "Dashboard")" class="nav-link @(If(Request.RawUrl.Contains("/Dashboard/Estados"), "active", ""))">
                                        <svg class="lucide lucide-map nav-icon w-5 h-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M9 3v15" /><path d="M15 5v15" /><path d="M3.3 7l7.7 4.7a2 2 0 0 0 2 0l7.7-4.7" /><path d="M21 4.6v12.8a1 1 0 0 1-.6.9l-4.6 2.3a2 2 0 0 1-1.8 0l-4.2-2.1a2 2 0 0 0-1.8 0l-3.7 1.8A1 1 0 0 1 3 19.4V6.6a1 1 0 0 1 .6-.9l4.6-2.3a2 2 0 0 1 1.8 0l4.2 2.1a2 2 0 0 0 1.8 0l3.7-1.8A1 1 0 0 1 21 4.6Z" /></svg>
                                        <p>Estados</p>
                                    </a>
                                </li>
                                <li class="nav-item">
                                    <a href="@Url.Action("Pedidos", "Dashboard")" class="nav-link @(If(Request.RawUrl.Contains("/Dashboard/Pedidos"), "active", ""))">
                                        <svg class="lucide lucide-package nav-icon w-5 h-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M3.3 7l7.7 4.7a2 2 0 0 0 2 0l7.7-4.7" /><path d="M21 4.6v12.8a1 1 0 0 1-.6.9l-4.6 2.3a2 2 0 0 1-1.8 0l-4.2-2.1a2 2 0 0 0-1.8 0l-3.7 1.8A1 1 0 0 1 3 19.4V6.6a1 1 0 0 1 .6-.9l4.6-2.3a2 2 0 0 1 1.8 0Z" /><path d="M12 22V12" /></svg>
                                        <p>Pedidos</p>
                                    </a>
                                </li>
                                <li class="nav-item">
                                    <a href="@Url.Action("VendasDiarias", "Dashboard")" class="nav-link @(If(Request.RawUrl.Contains("/Dashboard/VendasDiarias"), "active", ""))">
                                        <svg class="lucide lucide-calendar nav-icon w-5 h-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><rect width="18" height="18" x="3" y="4" rx="2" /><path d="M8 2v4" /><path d="M16 2v4" /><path d="M3 10h18" /></svg>
                                        <p>Vendas Diárias</p>
                                    </a>
                                </li>
                                <li class="nav-item">
                                    <a href="@Url.Action("VendasMensal", "Dashboard")" class="nav-link @(If(Request.RawUrl.Contains("/Dashboard/VendasMensal"), "active", ""))">
                                        <svg class="lucide lucide-chart-column nav-icon w-5 h-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18" /><path d="M18 17V9" /><path d="M13 17V5" /><path d="M8 17v-3" /></svg>
                                        <p>Vendas Mensais</p>
                                    </a>
                                </li>
                            </ul>
                        </li>
                    </ul>
                    <!--end::Sidebar Menu-->
                </nav>
            </div>
            <!--end::Sidebar Wrapper-->
        </aside>
        <!--end::Sidebar-->
        <!--begin::App Main-->
        <main class="app-main">
            <!--begin::App Content Header-->
            <div class="app-content-header">
                <!--begin::Container-->
                <div class="container-fluid">
                    <!--begin::Row-->
                    <div class="row">
                        <div class="col-sm-6"><h3 class="mb-0"></h3></div>
                        <div class="col-sm-6">
                            <ol class="breadcrumb float-sm-end">
                                @For i As Integer = 0 To parts.Length - 1
                                @Code
                                    breadcrumbPath &= If(i = 0, parts(i), "/" & parts(i))
                                End Code

                                @If i < parts.Length - 1 Then
                                @<li class="breadcrumb-item">
                                    <a href="@Url.Content("~/" & breadcrumbPath)">
                                        @parts(i)
                                    </a>
                                </li>
                                Else
                                @<li class="breadcrumb-item active" aria-current="page">
                                    @parts(i)
                                </li>
                                End If
                                    Next
                            </ol>
                        </div>

                    </div>
                    <!--end::Row-->
                </div>
                <button id="btnToggleFiltros" type="button" data-bs-toggle="collapse" data-bs-target="#filtroCollapse" aria-expanded="false" aria-controls="filtroCollapse">
                    <span>Filtros</span>
                    <span id="iconeSeta">&#x25BC;</span>
                </button>

                <div class="collapse p-3" id="filtroCollapse">
                    <div class="container-fluid">
                        <form>
                            <table class="table-borderless">
                                <tbody>
                                    <!-- Linha 1: Empresa -->
                                    <tr>
                                        <td><label for="empresaSelecionada" class="mb-0">Empresa:</label></td>
                                        <td>
                                            <div class="dropdown-multiselect" tabindex="0" onblur="fecharLista()">
                                                <div class="select-box" onclick="toggleLista()" style="width: 300px;">
                                                    <span id="lblSelecionados" style="flex: 1;">Selecione empresas</span>
                                                    <i class="bi bi-chevron-down ms-2"></i>
                                                </div>

                                                <div id="listaEmpresas" class="options-box" style="display: none; width: 300px;">
                                                    <div class="mb-2">
                                                        <button type="button" class="btn btn-sm btn-outline-primary me-2" onclick="marcarTodasEmpresas()">Marcar todas</button>
                                                        <button type="button" class="btn btn-sm btn-outline-secondary" onclick="desmarcarTodasEmpresas()">Desmarcar todas</button>
                                                    </div>

                                                    @If Model IsNot Nothing AndAlso Model.Filtro IsNot Nothing Then

                                                        For Each empresa In Model.Filtro.Empresas
                                                            Dim isChecked = Model.Filtro.EmpresasSelecionada IsNot Nothing AndAlso
                                                                            Model.Filtro.EmpresasSelecionada.Contains(empresa.Value)

                                                            @<div class="form-check">
                                                                <input type="checkbox"
                                                                       name="Filtro.EmpresasSelecionadas"
                                                                       value="@empresa.Value"
                                                                       class="form-check-input chkEmpresa"
                                                                       id="empresa_@empresa.Value"
                                                                       @(If(isChecked, "checked", ""))
                                                                       onchange="atualizarSelecionados()" />

                                                                <label class="form-check-label" for="empresa_@empresa.Value">@empresa.Text</label>
                                                            </div>
                                                        Next
                                                    End If
                                                </div>
                                            </div>
                                        </td>
                                    </tr>

                                    <!-- Linha 2: Seguimento -->
                                    <tr>
                                        <td><label for="seguimentoSelecionado" class="mb-0">Seguimento:</label></td>
                                        <td>
                                            @If Not Model Is Nothing AndAlso Not Model.Filtro Is Nothing Then
                                                @Html.DropDownListFor(Function(m) m.Filtro.SeguimentoSelecionado,
                  CType(Model.Filtro.Seguimentos, List(Of SelectListItem)),
                  "-- Selecione --",
                  New With {.class = "form-control", .style = "width: 300px;"})
                                            End If
                                        </td>
                                    </tr>

                                    <!-- Linha 3: Data De / Até -->
                                    <tr>
                                        <td><label for="dataInicio" class="mb-0">Período De:</label></td>
                                        <td>
                                            <div class="d-flex align-items-center gap-2">
                                                <input type="date" id="dataInicio" name="dataInicio"
                                                       value="@Model.Filtro.DataInicio.ToString("yyyy-MM-dd")"
                                                       class="form-control" style="width: 150px;" />

                                                <label for="dataFim" class="mb-0">Até:</label>

                                                <input type="date" id="dataFim" name="dataFim"
                                                       value="@Model.Filtro.DataFim.ToString("yyyy-MM-dd")"
                                                       class="form-control" style="width: 150px;" />
                                            </div>
                                        </td>
                                    </tr>

                                    <!-- Linha 4: Botão -->
                                    <tr>
                                        <td></td>
                                        <td class="text-end">
                                            <button id="btnFiltrar" class="btn btn-primary">Filtrar</button>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </form>
                    </div>
                </div>


                <!--end::Container-->
            </div>
            <!--end::App Content Header-->
            <!--begin::App Content-->
            <div class="app-content">
                <!--begin::Container-->
                <div class="container-fluid">
                    <!--begin::Row-->
                    <div class="row">
                        @RenderBody()
                    </div>
                    <!-- /.row (main row) -->
                </div>
                <!--end::Container-->
            </div>
            <!--end::App Content-->
        </main>
        <!--end::App Main-->
        <!--begin::Footer-->
        <footer class="app-footer">
            <!--begin::To the end-->
            <div class="float-end d-none d-sm-inline"></div>
            <!--end::To the end-->
            <!--begin::Copyright-->
            <strong>
                Copyright &copy; @DateTime.Now.Year.ToString()&nbsp;
                <a href="https://ngssolucoes.com.br/" class="text-decoration-none">NGS</a>.
            </strong>
            All rights reserved.
            <!--end::Copyright-->
        </footer>
        <!--end::Footer-->
    </div>


    <script src="@Url.Content("~/admin-lte/js/overlayscrollbars.browser.es6.min.js")"></script>
    <script src="@Url.Content("~/admin-lte/js/popper.min.js")"></script>
    <script src="@Url.Content("~/admin-lte/js/bootstrap.min.js")"></script>
    <script src="@Url.Content("~/admin-lte/js/adminlte.js")"></script>

    <script>

        const SELECTOR_SIDEBAR_WRAPPER = '.sidebar-wrapper';
        const Default = {
            scrollbarTheme: 'os-theme-light',
            scrollbarAutoHide: 'leave',
            scrollbarClickScroll: true,
        };

        document.addEventListener('DOMContentLoaded', function () {
            const sidebarWrapper = document.querySelector(SELECTOR_SIDEBAR_WRAPPER);
            if (sidebarWrapper && typeof OverlayScrollbarsGlobal?.OverlayScrollbars !== 'undefined') {
              OverlayScrollbarsGlobal.OverlayScrollbars(sidebarWrapper, {
                scrollbars: {
                theme: Default.scrollbarTheme,
                autoHide: Default.scrollbarAutoHide,
                clickScroll: Default.scrollbarClickScroll,
            },
            });
        }
        });

    </script>

    <script src="@Url.Content("~/admin-lte/js/Sortable.min.js")"></script>

    <script>
        const connectedSortables = document.querySelectorAll('.connectedSortable');
        connectedSortables.forEach((connectedSortable) => {
            let sortable = new Sortable(connectedSortable, {
                group: 'shared',
                handle: '.card-header',
            });
        });

        const cardHeaders = document.querySelectorAll('.connectedSortable .card-header');
        cardHeaders.forEach((cardHeader) => {
            cardHeader.style.cursor = 'move';
        });
    </script>

    <script src="@Url.Content("~/admin-lte/js/apexcharts.min.js")"></script>

    <script>
        // NOTICE!! DO NOT USE ANY OF THIS JAVASCRIPT
        // IT'S ALL JUST JUNK FOR DEMO
        // ++++++++++++++++++++++++++++++++++++++++++

        const sales_chart_options = {
            series: [
              {
                  name: 'Digital Goods',
                  data: [28, 48, 40, 19, 86, 27, 90],
              },
              {
                  name: 'Electronics',
                  data: [65, 59, 80, 81, 56, 55, 40],
              },
            ],
            chart: {
                height: 300,
                type: 'area',
                toolbar: {
                    show: false,
                },
            },
            legend: {
                show: false,
            },
            colors: ['#0d6efd', '#20c997'],
            dataLabels: {
                enabled: false,
            },
            stroke: {
                curve: 'smooth',
            },
            xaxis: {
                type: 'datetime',
                categories: [
                  '2023-01-01',
                  '2023-02-01',
                  '2023-03-01',
                  '2023-04-01',
                  '2023-05-01',
                  '2023-06-01',
                  '2023-07-01',
                ],
            },
            tooltip: {
                x: {
                    format: 'MMMM yyyy',
                },
            },
        };

        const sales_chart = new ApexCharts(
          document.querySelector('#revenue-chart'),
          sales_chart_options,
        );
        sales_chart.render();
    </script>

    <script src="@Url.Content("~/admin-lte/js/jsvectormap.min.js")"></script>
    <script src="@Url.Content("~/admin-lte/js/world.js")"></script>

    <script src="https://cdn.jsdelivr.net/npm/multiselect-dropdown@1.0.1/index.min.js"></script>

    <script>

        const visitorsData = {
            US: 398, // USA
            SA: 400, // Saudi Arabia
            CA: 1000, // Canada
            DE: 500, // Germany
            FR: 760, // France
            CN: 300, // China
            AU: 700, // Australia
            BR: 600, // Brazil
            IN: 800, // India
            GB: 320, // Great Britain
            RU: 3000, // Russia
        };

        // World map by jsVectorMap
        const map = new jsVectorMap({
            selector: '#world-map',
            map: 'world',
        });

        // Sparkline charts
        const option_sparkline1 = {
            series: [
              {
                  data: [1000, 1200, 920, 927, 931, 1027, 819, 930, 1021],
              },
            ],
            chart: {
                type: 'area',
                height: 50,
                sparkline: {
                    enabled: true,
                },
            },
            stroke: {
                curve: 'straight',
            },
            fill: {
                opacity: 0.3,
            },
            yaxis: {
                min: 0,
            },
            colors: ['#DCE6EC'],
        };

        const sparkline1 = new ApexCharts(document.querySelector('#sparkline-1'), option_sparkline1);
        sparkline1.render();

        const option_sparkline2 = {
            series: [
              {
                  data: [515, 519, 520, 522, 652, 810, 370, 627, 319, 630, 921],
              },
            ],
            chart: {
                type: 'area',
                height: 50,
                sparkline: {
                    enabled: true,
                },
            },
            stroke: {
                curve: 'straight',
            },
            fill: {
                opacity: 0.3,
            },
            yaxis: {
                min: 0,
            },
            colors: ['#DCE6EC'],
        };

        const sparkline2 = new ApexCharts(document.querySelector('#sparkline-2'), option_sparkline2);
        sparkline2.render();

        const option_sparkline3 = {
            series: [
              {
                  data: [15, 19, 20, 22, 33, 27, 31, 27, 19, 30, 21],
              },
            ],
            chart: {
                type: 'area',
                height: 50,
                sparkline: {
                    enabled: true,
                },
            },
            stroke: {
                curve: 'straight',
            },
            fill: {
                opacity: 0.3,
            },
            yaxis: {
                min: 0,
            },
            colors: ['#DCE6EC'],
        };

        const sparkline3 = new ApexCharts(document.querySelector('#sparkline-3'), option_sparkline3);
        sparkline3.render();

        document.addEventListener("DOMContentLoaded", function () {
            document.getElementById("btnFiltrar").addEventListener("click", function () {

                var inicio = document.getElementById("dataInicio").value;
                var fim = document.getElementById("dataFim").value;

                // Coleta todas as empresas selecionadas
                var empresa = Array.from(document.querySelectorAll('input[name="Filtro.EmpresasSelecionadas"]:checked'))
                    .map(cb => cb.value)
                    .join(',');

                var seguimento = document.getElementById("Filtro_SeguimentoSelecionado").value;
                var consolidarEmpresa = document.getElementById("chkConsolidar").checked;;
                window.location.href = "?empresaSelecionada=" + empresa + "&seguimentoSelecionado=" + seguimento + "&dataInicio=" + inicio + "&dataFim=" + fim;

            });
        });

        const iconeSeta = document.getElementById('iconeSeta');
        const filtroCollapse = document.getElementById('filtroCollapse');

        filtroCollapse.addEventListener('show.bs.collapse', () => {
            iconeSeta.innerHTML = '&#x25B2;'; // seta para cima
        });

        filtroCollapse.addEventListener('hide.bs.collapse', () => {
            iconeSeta.innerHTML = '&#x25BC;'; // seta para baixo
        });

        function toggleLista() {
            const lista = document.getElementById('listaEmpresas');
            lista.style.display = (lista.style.display === 'none' || lista.style.display === '') ? 'block' : 'none';
        }

        function fecharLista() {
            setTimeout(() => {
                document.getElementById('listaEmpresas').style.display = 'none';
            }, 200); // pequeno delay para não fechar antes do clique
        }

        function marcarTodasEmpresas() {
            document.querySelectorAll(".chkEmpresa").forEach(chk => chk.checked = true);
            atualizarSelecionados();
        }

        function desmarcarTodasEmpresas() {
            document.querySelectorAll(".chkEmpresa").forEach(chk => chk.checked = false);
            atualizarSelecionados();
        }

        function ajustarLarguraSelectBox() {
            const label = document.getElementById("lblSelecionados");
            const selectBox = document.querySelector(".select-box");
            const listaBox = document.getElementById("listaEmpresas");

            // Cria span oculto para medir o texto
            const span = document.createElement("span");
            span.style.visibility = "hidden";
            span.style.position = "absolute";
            span.style.whiteSpace = "nowrap";
            span.style.font = window.getComputedStyle(label).font;
            span.style.fontSize = window.getComputedStyle(label).fontSize;
            span.style.fontWeight = window.getComputedStyle(label).fontWeight;
            span.innerText = label.innerText;

            document.body.appendChild(span);
            const larguraTexto = span.offsetWidth + 48; // margem + ícone + padding
            document.body.removeChild(span);

            const larguraFinal = Math.max(200, Math.min(1000, larguraTexto));
            selectBox.style.width = '100%'; //`${larguraFinal}px`;
            listaBox.style.width = '100%'; //`${larguraFinal}px`; // dropdown igual

        }

        function atualizarSelecionados() {
            const checks = document.querySelectorAll(".chkEmpresa:checked");
            const label = document.getElementById("lblSelecionados");

            if (checks.length === 0) {
                label.innerText = "Selecione empresas";
            } else if (checks.length === 1) {
                const id = checks[0].id;
                const texto = document.querySelector(`label[for='${id}']`).innerText;
                label.innerText = texto;
            } else {
                label.innerText = `${checks.length} selecionadas`;
            }

            ajustarLarguraSelectBox();
        }

        document.addEventListener("DOMContentLoaded", atualizarSelecionados);

    </script>

    @RenderSection("scripts", required:=False)

</body>
</html>
