@ModelType IEnumerable(Of ProdutoDashboard)

@Code
    ' Montagem do objeto JSON que será usado pelo Chart.js
    Dim chartDataEvolucaoVendasMensais = New With {
        .labels = Model.Select(Function(p) p.Data).ToArray(),
        .datasets = New Object() {
            New With {
                .label = "Vendas",
                .data = Model.Select(Function(p) p.Valor).ToArray(),
                .borderColor = "#007bff", ' Cor da linha
                .backgroundColor = "transparent", ' Nenhum preenchimento abaixo da linha
                .tension = 0.3 ' Suaviza a curva
            }
        }
    }
End Code

<!-- Script para armazenar os dados JSON (Chart.js vai consumir isso) -->
<script type="application/json" id="chartDataEvolucaoVendasMensais">
    @Html.Raw(Json.Encode(chartDataEvolucaoVendasMensais))
</script>

<!-- Canvas para renderização do gráfico -->
<div class="card">
    <div class="card-header">
        <h3 class="card-title">Evolução de Vendas Mensais</h3>
    </div>
    <div class="card-body">
        <canvas id="evolucaoVendasMensais"></canvas>
    </div>
</div>

<script>
    $(document).ready(function () {
        try {
            // Carrega os dados do elemento oculto gerado pelo Razor
            const data = JSON.parse(document.getElementById('chartDataEvolucaoVendasMensais').textContent);

            // Quantidade e Valor para cada ponto (assumindo mesmo índice que labels)
            const quantidades = @Html.Raw(Json.Encode(Model.Select(Function(p) p.QuantidadeFiscal).ToArray()));
            const valores = @Html.Raw(Json.Encode(Model.Select(Function(p) p.Valor).ToArray()));

            const ctx = document.getElementById('evolucaoVendasMensais').getContext('2d');
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

