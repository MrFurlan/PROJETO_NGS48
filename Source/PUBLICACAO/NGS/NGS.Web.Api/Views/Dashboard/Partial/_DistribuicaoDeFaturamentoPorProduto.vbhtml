@ModelType IEnumerable(Of ProdutoDashboard)

@Code

    ' Lista de cores fixas para os 10 primeiros seguimentos
    Dim coresFixas = New String() {
"#2274A5", '#1
"#F75C03", '#2
"#F1C40F", '#3
"#D90368", '#4
"#00CC66", '#5
"#0099C6", '#6
"#DD4477", '#7
"#66AA00", '#8
"#B82E2E", '#9
"#316395"  '#10
}

    ' Montagem do objeto JSON que será usado pelo Chart.js
    Dim chartDataDistribuicao = New With {
.labels = Model.Select(Function(p) p.Seguimento).ToArray(),
.datasets = New Object() {
New With {
.label = "Faturamento (R$)",
.data = Model.Select(Function(p) p.Valor).ToArray(),
.backgroundColor = Model.Select(Function(p, i)
                                    If i < coresFixas.Length Then
                                        Return coresFixas(i)
                                    Else
                                        Dim seguimento = p.Seguimento.ToUpperInvariant()
                                        Dim hash = Math.Abs(seguimento.GetHashCode())
                                        Return HslToHex(hash Mod 360, 70, 50)
                                    End If
                                End Function).ToArray()
}
}
}
End Code

<!-- Script para armazenar os dados JSON (Chart.js vai consumir isso) -->
<script type="application/json" id="chartDataDistribuicao">
    @Html.Raw(Json.Encode(chartDataDistribuicao))
</script>

<!-- Canvas para renderização do gráfico -->
<div class="card">
    <div class="card-header">
        <h3 class="card-title">Distribuição de Faturamento por Produto</h3>
    </div>
    <div class="card-body">
        <canvas id="distribuicaoFaturamentoProduto"></canvas>
    </div>
</div>

<script>

    $(document).ready(function () {
        try {
            // Carregar dados da PartialView
            const data = JSON.parse(document.getElementById('chartDataDistribuicao').textContent);

            // Verificação para debug (opcional)
            console.log("Dados carregados:", data);

            // Inicializar o gráfico
            const ctx = document.getElementById('distribuicaoFaturamentoProduto').getContext('2d');
            new Chart(ctx, {
                type: 'pie',
                data: data,
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        tooltip: {
                            callbacks: {
                                label: function (tooltipItem) {
                                    const index = tooltipItem.dataIndex;
                                    const nomeProduto = @Html.Raw(Json.Encode(Model.Select(Function(p) p.Seguimento).ToArray()));
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
