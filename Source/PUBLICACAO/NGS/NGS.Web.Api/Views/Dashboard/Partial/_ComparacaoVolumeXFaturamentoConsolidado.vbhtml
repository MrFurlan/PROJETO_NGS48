@ModelType IEnumerable(Of ProdutoDashboard)

@Code
    ' Cores fixas para os datasets
    Dim corFaturamento = "#3366CC" ' Azul
    Dim corVolume = "#DC3912"      ' Vermelho

    ' Montagem do objeto JSON para Chart.js
    Dim chartDataComparacaoVolumeXFaturamento = New With {
        .labels = Model.Select(Function(p) p.Periodo).ToArray(),
        .datasets = New Object() {
            New With {
                .label = "Faturamento (R$)",
                .data = Model.Select(Function(p) p.Faturamento).ToArray(),
                .backgroundColor = Enumerable.Repeat(corFaturamento, Model.Count()).ToArray()
            },
            New With {
                .label = "Volume (Kg)",
                .data = Model.Select(Function(p) p.VolumeKg).ToArray(),
                .backgroundColor = Enumerable.Repeat(corVolume, Model.Count()).ToArray()
            }
        }
    }
End Code

<!-- Script para armazenar os dados JSON (Chart.js vai consumir isso) -->
<script type="application/json" id="chartDataComparacaoVolumeXFaturamento">
    @Html.Raw(Json.Encode(chartDataComparacaoVolumeXFaturamento))
</script>

<!-- Canvas para renderização do gráfico -->
<div class="card">
    <div class="card-header">
        <h3 class="card-title">Comparação Faturamento e Volume</h3>
    </div>
    <div class="card-body">
        <canvas id="dataComparacaoVolumeXFaturamento"></canvas>
    </div>
</div>

<!-- Script de inicialização do Chart.js -->
<script>
    $(document).ready(function () {
        try {
            const data = JSON.parse(document.getElementById('chartDataComparacaoVolumeXFaturamento').textContent);

            const ctx = document.getElementById('dataComparacaoVolumeXFaturamento').getContext('2d');
            new Chart(ctx, {
                type: 'bar',
                data: data,
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        tooltip: {
                            callbacks: {
                                label: function (tooltipItem) {
                                    const valor = tooltipItem.raw;
                                    const label = tooltipItem.dataset.label;

                                    if (label.includes("Faturamento")) {
                                        return `${label}: ${valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })}`;
                                    } else {
                                        return `${label}: ${valor.toLocaleString('pt-BR')} Kg`;
                                    }
                                }
                            }
                        },
                        legend: { position: 'top' }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            grid: {
                                drawBorder: true
                            }
                        },
                        x: {
                            grid: {
                                drawBorder: true
                            }
                        }
                    }
                }
            });

        } catch (e) {
            console.error("Erro ao carregar o gráfico:", e.message);
        }
    });
</script>
