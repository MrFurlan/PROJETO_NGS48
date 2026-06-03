@ModelType  ProdutoDashboard
@Code
    Layout = "~/Views/Shared/_LayoutDashboard.vbhtml"
End Code

@Section Styles

    <style>

        /* Garantir altura uniforme de 250px para todos os gráficos */
        .card-body canvas {
            height: 250px !important;
        }

        @@Media (max-width: 768px) {
            .chart-col {
                flex: 0 0 100%;
                max-width: 100%;
            }
        }

    </style>

End Section


@Code
    Dim totalQuantidade As Decimal = Model.QuantidadeFiscal
    Dim totalValor As Decimal = Model.Valor
End Code

<!-- Content Header (Page header) -->
<section class="content-header"></section>

<!-- Main content -->
<section class="content">
    <section class="content-header">
        <div class="row">
            <!--begin::Col-->
            <div class="col-lg-3 col-6">
                <!--begin::Small Box Widget 3-->
                <div class="small-box text-bg-warning">
                    <div class="inner">
                        <h3>@totalQuantidade.ToString("N0")</h3>
                        <p>Volume Faturado</p>
                    </div>
                    <svg class="small-box-icon"
                         fill="currentColor"
                         viewBox="0 0 24 24"
                         xmlns="http://www.w3.org/2000/svg"
                         aria-hidden="true">
                        <path d="M6.25 6.375a4.125 4.125 0 118.25 0 4.125 4.125 0 01-8.25 0zM3.25 19.125a7.125 7.125 0 0114.25 0v.003l-.001.119a.75.75 0 01-.363.63 13.067 13.067 0 01-6.761 1.873c-2.472 0-4.786-.684-6.76-1.873a.75.75 0 01-.364-.63l-.001-.122zM19.75 7.5a.75.75 0 00-1.5 0v2.25H16a.75.75 0 000 1.5h2.25v2.25a.75.75 0 001.5 0v-2.25H22a.75.75 0 000-1.5h-2.25V7.5z"></path>
                    </svg>
                    <a href="#"
                       class="small-box-footer link-dark link-underline-opacity-0 link-underline-opacity-50-hover">
                        + info <i class="bi bi-link-45deg"></i>
                    </a>
                </div>
                <!--end::Small Box Widget 3-->
            </div>
            <!--end::Col-->
            <div class="col-lg-3 col-6">
                <!--begin::Small Box Widget 1-->
                <div class="small-box text-bg-primary">
                    <div class="inner">
                        <h3>@totalValor.ToString("C")</h3>
                        <p>Total Faturado</p>
                    </div>
                    <svg class="small-box-icon"
                         fill="currentColor"
                         viewBox="0 0 24 24"
                         xmlns="http://www.w3.org/2000/svg"
                         aria-hidden="true">
                        <path d="M2.25 2.25a.75.75 0 000 1.5h1.386c.17 0 .318.114.362.278l2.558 9.592a3.752 3.752 0 00-2.806 3.63c0 .414.336.75.75.75h15.75a.75.75 0 000-1.5H5.378A2.25 2.25 0 017.5 15h11.218a.75.75 0 00.674-.421 60.358 60.358 0 002.96-7.228.75.75 0 00-.525-.965A60.864 60.864 0 005.68 4.509l-.232-.867A1.875 1.875 0 003.636 2.25H2.25zM3.75 20.25a1.5 1.5 0 113 0 1.5 1.5 0 01-3 0zM16.5 20.25a1.5 1.5 0 113 0 1.5 1.5 0 01-3 0z"></path>
                    </svg>
                    <a href="#"
                       class="small-box-footer link-light link-underline-opacity-0 link-underline-opacity-50-hover">
                        + info <i class="bi bi-link-45deg"></i>
                    </a>
                </div>
                <!--end::Small Box Widget 1-->
            </div>
            <!--end::Col-->
            @*<div class="col-lg-3 col-6">
                    <!--begin::Small Box Widget 2-->
                    <div class="small-box text-bg-success">
                        <div class="inner">
                            <h3>cvcvcvcv</h3>
                            <p>Volume Vendido (kg)</p>
                        </div>
                        <svg class="small-box-icon"
                             fill="currentColor"
                             viewBox="0 0 24 24"
                             xmlns="http://www.w3.org/2000/svg"
                             aria-hidden="true">
                            <path d="M18.375 2.25c-1.035 0-1.875.84-1.875 1.875v15.75c0 1.035.84 1.875 1.875 1.875h.75c1.035 0 1.875-.84 1.875-1.875V4.125c0-1.036-.84-1.875-1.875-1.875h-.75zM9.75 8.625c0-1.036.84-1.875 1.875-1.875h.75c1.036 0 1.875.84 1.875 1.875v11.25c0 1.035-.84 1.875-1.875 1.875h-.75a1.875 1.875 0 01-1.875-1.875V8.625zM3 13.125c0-1.036.84-1.875 1.875-1.875h.75c1.036 0 1.875.84 1.875 1.875v6.75c0 1.035-.84 1.875-1.875 1.875h-.75A1.875 1.875 0 013 19.875v-6.75z"></path>
                        </svg>
                        <a href="#"
                           class="small-box-footer link-light link-underline-opacity-0 link-underline-opacity-50-hover">
                            + info <i class="bi bi-link-45deg"></i>
                        </a>
                    </div>
                    <!--end::Small Box Widget 2-->
                </div>*@
            <!--end::Col-->

        </div>
        <h3>Análise de Produtos</h3>
    </section>
    <!-- Default box -->
    <div class="wrapper">
        <div class="content-wrapper p-4">
            <section class="content">
                <div class="container-fluid">
                    <div class="row" id="dashboard-charts">
                        <div class="col-md-6 mb-4">
                            <div class="chart-container" style="position: relative; width: 100%; height: 400px;">
                                <div style="position: absolute; top: 0; left: 0; width: 100%; height: 100%;">
                                    @Html.Action("FaturamentoPorProduto", "Dashboard")
                                </div>
                            </div>
                        </div>
                        <div class="col-md-6 mb-4">
                            <div class="chart-container" style="position: relative; width: 100%; height: 400px;">
                                <div style="position: absolute; top: 0; left: 0; width: 100%; height: 100%;">
                                    @Html.Action("DistribuicaoDeFaturamentoPorProduto", "Dashboard")
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="row" id="dashboard-charts">
                        @Html.Action("DetalhamentoProdutos", "Dashboard")
                    </div>
                </div>
            </section>
        </div>
    </div>
    <!-- /.box -->


</section>
<!-- /.content -->

@section scripts

    <script src="@Url.Content("~/admin-lte/js/chart.js")"></script>

    <script type="text/javascript">

        $(document).ready(function () {
            if (typeof (Storage) !== "undefined") {
                sessionStorage.removeItem("divSelecionada");
            }
        });
        
    </script>

End Section
