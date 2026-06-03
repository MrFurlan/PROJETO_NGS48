@ModelType List(Of ProdutoDashboard)

@Code
    Dim totalQuantidade As Decimal = Model.Sum(Function(p) p.QuantidadeFiscal)
    Dim totalValor As Decimal = Model.Sum(Function(p) p.Valor)
    Dim totalCidades As Integer = Model.Count
    Dim totalEstados As Integer = Model.Select(Function(p) p.Estado).Distinct().Count()
End Code

<h2>Detalhes das Cidades</h2>

<table id="tabelaProdutos" class="table table-bordered table-striped" style="width: 100%;">
    <thead>
        <tr>
            <th>Cidade</th>
            <th>Estado</th>
            <th>Quantidade</th>
            <th>Valor</th>
        </tr>
    </thead>
    <tbody>
        @For Each produto In Model
            @<tr>
                <td>@produto.Cidade</td>
                <td>@produto.Estado </td>
                 <td>@produto.QuantidadeFiscal.ToString("N0")</td>
                 <td>@produto.Valor.ToString("C")</td>
            </tr>
        Next
    </tbody>
    <tfoot>
        <tr style="background-color: #d3d3d3; font-weight: bold;">
            <td>Total: @totalCidades</td>
            <td>Total:@totalEstados</td>
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

<script>
    $(function () {
        $('#tabelaProdutos').DataTable({
            "responsive": true,
            "lengthChange": true,
            "autoWidth": false,
            "pageLength": 10,
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
