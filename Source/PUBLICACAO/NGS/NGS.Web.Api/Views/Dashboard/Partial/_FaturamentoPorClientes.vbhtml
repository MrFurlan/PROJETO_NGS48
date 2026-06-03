@ModelType IEnumerable(Of ProdutoDashboard)

@Code

    ' Montagem do objeto JSON que será usado pelo Chart.js
    Dim chartDataClientes = New With {
.labels = Model.Select(Function(p) p.ClienteFantasia).ToArray(),
.datasets = New Object() {
    New With {
        .label = "Faturamento (R$)",
        .data = Model.Select(Function(p) p.Valor).ToArray(),
        .backgroundColor = Model.Select(Function(p) "#43A047").ToArray()
    }
}
}
End Code

<!-- Script para armazenar os dados JSON (Chart.js vai consumir isso) -->
<script type="application/json" id="chartDataClientes">
    @Html.Raw(Json.Encode(chartDataClientes))
</script>

<!-- Canvas para renderização do gráfico -->
<div class="card">
    <div class="card-header">
        <h3 class="card-title">Valor Total por Cliente (Top 10)</h3>
    </div>
    <div class="card-body">
        <canvas id="faturamentoClientes"></canvas>
    </div>
</div>

<script>

    $(document).ready(function () {
        try {
            // Carregar dados da PartialView
            const data = JSON.parse(document.getElementById('chartDataClientes').textContent);

            // Verificação para debug (opcional)
            console.log("Dados carregados:", data);

            // Inicializar o gráfico
            const ctx = document.getElementById('faturamentoClientes').getContext('2d');
            new Chart(ctx, {
                type: 'bar',
                data: data,
                options: {
                    indexAxis: 'y', // <<< Torna o gráfico horizontal
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        tooltip: {
                            callbacks: {
                                label: function (tooltipItem) {
                                    const index = tooltipItem.dataIndex;
                                    const nomeCliente = @Html.Raw(Json.Encode(Model.Select(Function(p) p.ClienteFantasia).ToArray()));
                                    const valorCliente = tooltipItem.raw.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
                                    return `${nomeCliente[index]}: ${valorCliente}`;
                                }
                            }
                        },
                        legend: { position: 'top' },
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
