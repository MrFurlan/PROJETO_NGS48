@ModelType List(Of ProdutoDashboard)

<h2>Resumo Consolidado</h2>

<table id="tabelaProdutos" class="table table-bordered table-striped" style="width: 100%;">
    <thead>
        <tr>
            <th>Periodo</th>
            <th>Pedidos</th>
            <th>Volume (Kg)</th>
            <th>Faturamento</th>
            <th>Volume Médio</th>
            <th>Faturamento Médio</th>
        </tr>
    </thead>
    <tbody>
        @For Each produto In Model
            @<tr>
                <td>@produto.Periodo</td>
                <td>@produto.Pedidos</td>
                <td>@produto.VolumeKg.ToString("N0")</td>
                 <td>@produto.Faturamento.ToString("C")</td>
                <td>@produto.VolumeMedio.ToString("N0")</td>
                <td>@produto.FaturamentoMedio.ToString("C")</td>
            </tr>
        Next
    </tbody>
</table>

<style>
    table {
        width: 100%;
        margin-top: 20px;
    }

    th {
        background-color: #007bff;
        color: white;
        text-align: left;
    }
</style>

<script src="@Url.Content("~/admin-lte/js/jquery-3.7.1.min.js")"></script>

<!-- Bootstrap (já incluso com AdminLTE normalmente) -->
<link rel="stylesheet" href="@Url.Content("~/admin-lte/css/bootstrap.min.css")">
<script src="@Url.Content("~/admin-lte/js/bootstrap.bundle.min.js")"></script>

<!-- DataTables com Bootstrap 4 (AdminLTE usa Bootstrap 4) -->
<link rel="stylesheet" href="@Url.Content("~/admin-lte/css/dataTables.bootstrap4.min.css")">
<script src="@Url.Content("~/admin-lte/js/jquery.dataTables.min.js")"></script>
<script src="@Url.Content("~/admin-lte/js/dataTables.bootstrap4.min.js")"></script>

<script>
    $(function () {
        $('#tabelaProdutos').DataTable({
            "responsive": true,
            "lengthChange": true,
            "autoWidth": false,
            "pageLength": 10,
            "ordering": false, // <-- Isso desativa a ordenação automática
            "lengthMenu": [[10, 25, 50, -1], [10, 25, 50, "Todos"]],
            "language": {
                "lengthMenu": "Mostrar _MENU_ registros por página",
                "zeroRecords": "Nenhum registro encontrado",
                "info": "Mostrando página _PAGE_ de _PAGES_",
                "infoEmpty": "Sem registros disponíveis",
                "infoFiltered": "(filtrado de _MAX_ registros totais)",
                "search": "Filtrar:",
                "paginate": {
                    "previous": "Anterior",
                    "next": "Próximo"
                }
            }
        });
    });
</script>
