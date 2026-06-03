@ModelType IEnumerable(Of ProdutoDashboard)

@Code
    ' Montagem do objeto JSON que será usado pelo Chart.js
    Dim chartDataEvolucaoConsolidado = New With {
        .labels = Model.Select(Function(p) p.Periodo).ToArray(),
        .datasets = New Object() {
            New With {
                .label = "Vendas",
                .data = Model.Select(Function(p) p.Faturamento).ToArray(),
                .borderColor = "#007bff", ' Cor da linha
                .backgroundColor = "transparent", ' Nenhum preenchimento abaixo da linha
                .tension = 0.3 ' Suaviza a curva
            }
        }
    }
End Code

<!-- Script para armazenar os dados JSON (Chart.js vai consumir isso) -->
<script type="application/json" id="chartDataEvolucaoConsolidado">
    @Html.Raw(Json.Encode(chartDataEvolucaoConsolidado))
</script>

<!-- Canvas para renderização do gráfico -->
<div class="card">
    <div class="card-header">
        <h3 class="card-title">Evolução de Faturamento</h3>
    </div>
    <div class="card-body">
        <canvas id="evolucaoConsolidado"></canvas>
    </div>
</div>

<script>
    $(document).ready(function () {
        try {
            // Carrega os dados do elemento oculto gerado pelo Razor
            const data = JSON.parse(document.getElementById('chartDataEvolucaoConsolidado').textContent);

            // Quantidade e Valor para cada ponto (assumindo mesmo índice que labels)
            const quantidades = @Html.Raw(Json.Encode(Model.Select(Function(p) p.VolumeKg).ToArray()));
            const valores = @Html.Raw(Json.Encode(Model.Select(Function(p) p.Faturamento).ToArray()));

            const ctx = document.getElementById('evolucaoConsolidado').getContext('2d');
            new Chart(ctx, {
                type: 'line',
                data: data,
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        tooltip: {
                            callbacks: {
                                label: function (context) {
                                    const index = context.dataIndex;
                                    const qtd = quantidades[index];
                                    const val = valores[index];

                                    const valorFormatado = val.toLocaleString('pt-BR', {
                                        style: 'currency',
                                        currency: 'BRL'
                                    });

                                    return `Qtd: ${qtd} | Valor: ${valorFormatado}`;
                                }
                            }
                        },
                        legend: {
                            position: 'top'
                        }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            title: {
                                display: true,
                                text: 'Valor (R$)'
                            },
                            grid: {
                                drawBorder: true
                            }
                        },
                        x: {
                            title: {
                                display: true,
                                text: 'Data'
                            },
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

