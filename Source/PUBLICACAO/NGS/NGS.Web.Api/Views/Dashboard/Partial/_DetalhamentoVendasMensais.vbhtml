@ModelType List(Of ProdutoDashboard)

@Code
    Dim totalQuantidadePedido As Decimal = Model.Sum(Function(p) p.QuantidadePedido)
    Dim totalQuantidade As Decimal = Model.Sum(Function(p) p.QuantidadeFiscal)
    Dim totalValor As Decimal = Model.Sum(Function(p) p.Valor)
    Dim totalProdutos As Integer = Model.Count
End Code

<h2>Detalhes de Vendas Mensais</h2>
 
<table id="tabelaVendasMensais" class="table table-bordered table-striped" style="width: 100%;">
    <thead>
        <tr>
            <th>Data</th>
            <th>Pedido</th>
            <th>Quantidade</th>
            <th>Valor</th>
        </tr>
    </thead>
    <tbody>
        @For Each produto In Model
            @<tr>
                <td>@produto.Data</td>
                <td>@produto.QuantidadePedido.ToString("N0")</td>
                <td>@produto.QuantidadeFiscal.ToString("N0")</td>
                <td>@produto.Valor.ToString("C")</td>
            </tr>
        Next
    </tbody>
    <tfoot>
        <tr style="background-color: #d3d3d3; font-weight: bold;">
            <td>Total:</td>
            <td>@totalQuantidadePedido.ToString("N0")</td>
            <td>@totalQuantidade.ToString("N0")</td>
            <td>@totalValor.ToString("C")</td>
        </tr>
    </tfoot>
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
<script src="@Url.Content("~/admin-lte/js/moment.min.js")"></script>
<script src="@Url.Content("~/admin-lte/js/datetime-moment.js")"></script>

<script>
    $(document).ready(function () {
        // Registra o formato 'DD/MM/YYYY' para o DataTables
        $.fn.dataTable.moment('DD/MM/YYYY');

        $('#tabelaVendasMensais').DataTable({
            "responsive": true,
            "lengthChange": true,
            "autoWidth": false,
            "pageLength": 10,
            "order": [[0, "asc"]], // Coluna 0 = Data
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