@ModelType IEnumerable(Of ProdutoDashboard)

@Code

    ' Lista de cores fixas para os 10 primeiros seguimentos
    Dim coresFixas = New String() {
"#3366CC", '#1
"#DC3912", '#2
"#FF9900", '#3
"#109618", '#4
"#990099", '#5
"#0099C6", '#6
"#DD4477", '#7
"#66AA00", '#8
"#B82E2E", '#9
"#316395"  '#10
}

    ' Montagem do objeto JSON que será usado pelo Chart.js
    Dim chartDataPorPedidos = New With {
        .labels = Model.Select(Function(p) p.Pedidos).ToArray(),
        .datasets = New Object() {
            New With {
                .label = "Faturamento (R$)",
                .data = Model.Select(Function(p) p.Valor).ToArray(),
                .backgroundColor = Model.Select(Function(p) "#3366CC").ToArray()
            }
        }
    }
End Code


<!-- Script para armazenar os dados JSON (Chart.js vai consumir isso) -->
<script type="application/json" id="chartDataPorPedidos">
    @Html.Raw(Json.Encode(chartDataPorPedidos))
</script>

<!-- Canvas para renderização do gráfico -->
<div class="card">
    <div class="card-header">
        <h3 class="card-title">Valor Total por Pedidos (Top 10)</h3>
    </div>
    <div class="card-body">
        <canvas id="faturamentoPedidos"></canvas>
    </div>
</div>

<script>

    $(document).ready(function () {
        try {
            // Carregar dados da PartialView
            const data = JSON.parse(document.getElementById('chartDataPorPedidos').textContent);

            // Verificação para debug (opcional)
            console.log("Dados carregados:", data);

            // Inicializar o gráfico
            const ctx = document.getElementById('faturamentoPedidos').getContext('2d');
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
                                    const index = tooltipItem.dataIndex;
                                    const nomeProduto = @Html.Raw(Json.Encode(Model.Select(Function(p) p.Pedidos).ToArray()));
                                    const valorProduto = tooltipItem.raw.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
                                    return `${nomeProduto[index]}: ${valorProduto}`;
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

