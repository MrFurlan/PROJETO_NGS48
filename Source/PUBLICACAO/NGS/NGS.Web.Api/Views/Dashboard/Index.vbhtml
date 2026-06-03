@Code
    Layout = "~/Views/Shared/_LayoutDashboard.vbhtml"
End Code

@Section Styles
    <style>
        .chart-container {
            margin-bottom: 30px;
            height: 240px;
        }

            .chart-container h5 {
                font-weight: normal;
            }

            .chart-container canvas {
                height: 200px !important;
                width: 100% !important;
            }
    </style>
End Section

<section class="content-header">
    <h3>📊 Dashboard ERP – Indicadores Gerenciais</h3>
</section>

<section class="content">
    <div class="container-fluid">
        <div class="row">
            @Code
                Dim titulos = {
                    "Pedidos Aguardando Liberação",
                    "Caminhões Aguardando para Descarga",
                    "Títulos a Pagar",
                    "Títulos a Receber",
                    "Total Faturado no Dia",
                    "Indicador de Atrasos",
                    "Pedidos com Divergência",
                    "Volume de Vendas por Região",
                    "Notas Faturadas por Dia",
                    "Notas Canceladas por Semana",
                    "Total Descarregado no Dia",
                    "Recebimentos por Cliente (Top 5)",
                    "Pedidos Faturados por Vendedor",
                    "Volume Diário de Produção",
                    "Inadimplência (%)",
                    "Custo Médio por Produto",
                    "Tempo Médio de Entrega",
                    "Evolução de Compras Mensal",
                    "Saldo de Estoque Crítico",
                    "Taxa de Retorno de Clientes"
                }

                For i As Integer = 0 To titulos.Length - 1
            End Code
            <div class="col-lg-3 col-md-6 chart-container">
                <div class="card p-3">
                    <h5 class="mb-3">@titulos(i)</h5>
                    <canvas id="grafico@(i)"></canvas>
                </div>
            </div>
            @Code Next End Code
        </div>
    </div>
</section>

@section Scripts
    <script src="@Url.Content("~/admin-lte/js/chart.js")"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2"></script>
    <script>
        document.addEventListener("DOMContentLoaded", () => {
            const tipos = [
                "bar", "bar", "pie", "doughnut", "line", "bar", "pie", "polarArea",
                "bar", "bar", "bar", "bar", "bar", "bar", "doughnut", "bar", "line", "bar", "bar", "pie"
            ];

            const coresDinamicas = [
                ["#0052cc", "#2684ff", "#4c9aff", "#b3d4ff"],
                ["#ff5630", "#ff7452", "#ff8f73", "#ffc1b3"],
                ["#36b37e", "#57d9a3", "#79f2c0", "#abf5d1"],
                ["#ffab00", "#ffc400", "#ffe380", "#fff0b3"],
                ["#6554c0", "#8777d9", "#998dd9", "#c0b6f2"],
                ["#00b8d9", "#79e2f2", "#b3f5ff", "#e6fcff"]
            ];

            const dados = [
                { labels: ["Comercial", "Financeiro", "Logística"], data: [10, 15, 8] },
                { labels: ["Aguardando Descarga"], data: [25] },
                { labels: ["Fornecedores", "Folha", "Outros"], data: [38, 42, 20] },
                { labels: ["Em aberto", "Quitado"], data: [30, 70] },
                { labels: ["Dia 1", "5", "10", "15", "20", "25"], data: [12, 35, 40, 50, 60, 75] },
                { labels: ["Entregas", "Faturamento", "Estoque"], data: [4, 6, 5] },
                { labels: ["Fiscais", "Financeiros", "Logísticos"], data: [3, 4, 6] },
                { labels: ["Norte", "Sul", "Leste", "Oeste"], data: [35, 28, 30, 30] },
                { labels: ["Seg", "Ter", "Qua", "Qui", "Sex"], data: [15, 23, 40, 12, 33] },
                { labels: ["Semana 1", "Semana 2", "Semana 3"], data: [5, 7, 12] },
                { labels: ["Total Descarregado"], data: [50] },
                { labels: ["Cliente A", "B", "C", "D", "E"], data: [20, 15, 10, 8, 5] },
                { labels: ["João", "Maria", "Carlos"], data: [18, 25, 16] },
                { labels: ["Volume Produzido"], data: [40] },
                { labels: ["Inadimplente", "Regular"], data: [20, 80] },
                { labels: ["Produto A", "Produto B", "Produto C"], data: [12.5, 10.2, 14.8] },
                { labels: ["Jan", "Fev", "Mar", "Abr"], data: [7, 6, 8, 5] },
                { labels: ["Jan", "Fev", "Mar", "Abr", "Mai"], data: [120, 150, 180, 160, 140] },
                { labels: ["Item 1", "Item 2", "Item 3"], data: [2, 4, 1] },
                { labels: ["Retornaram", "Não retornaram"], data: [60, 40] }
            ];

            for (let i = 0; i < dados.length; i++) {
                new Chart(document.getElementById("grafico" + i), {
                    type: tipos[i],
                    data: {
                        labels: dados[i].labels,
                        datasets: [{
                            label: "Valor",
                            data: dados[i].data,
                            backgroundColor: coresDinamicas[i % coresDinamicas.length],
                            borderColor: "#fff",
                            borderWidth: 1,
                            fill: tipos[i] === "line"
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            legend: { display: true },
                            datalabels: {
                                anchor: "end",
                                align: "top",
                                color: "#000",
                                font: { weight: "normal" },
                                formatter: Math.round
                            }
                        }
                    },
                    plugins: [ChartDataLabels]
                });
            }
        });
    </script>
End Section
