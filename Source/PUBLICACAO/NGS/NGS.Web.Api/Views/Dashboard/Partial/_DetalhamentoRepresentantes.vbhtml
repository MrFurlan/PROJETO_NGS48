@ModelType List(Of ProdutoDashboard)

@Code
    Dim totalQuantidade As Decimal = Model.Sum(Function(p) p.QuantidadeFiscal)
    Dim totalValor As Decimal = Model.Sum(Function(p) p.Valor)
    Dim totalRepresentantes As Integer = Model.Count
End Code

<h2>Detalhes dos Representantes</h2>

<table id="tabelaRepresentantes" class="table table-bordered table-striped" style="width: 100%;">
    <thead>
        <tr>
            <th>Representante</th>
            <th>Nome</th>
            <th>Quantidade</th>
            <th>Valor</th>
        </tr>
    </thead>
    <tbody>
        @For Each Representante In Model
            @<tr>
                <td>@Representante.RepresentanteId</td>
                <td>@Representante.RepresentanteFantasia</td>
                <td>@Representante.QuantidadeFiscal.ToString("N0")</td>
                <td>@Representante.Valor.ToString("C")</td>
            </tr>
        Next
    </tbody>
    <tfoot>
        <tr style="background-color: #d3d3d3; font-weight: bold;">
            <td>Total (@totalRepresentantes Representante)</td>
            <td></td>
            <td>@totalQuantidade.ToString("N0")</td>
            <td>@totalValor.ToString("C")</td>
        </tr>
    </tfoot>
</table>

<br />
<br />

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
        $('#tabelaRepresentantes').DataTable({
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