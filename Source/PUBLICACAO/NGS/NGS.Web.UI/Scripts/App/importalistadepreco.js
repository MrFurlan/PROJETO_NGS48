$(document).ready(function () {
    inicializacaoImportaListaDePrecos();

    var prmImportaListaDePrecos = Sys.WebForms.PageRequestManager.getInstance();
    prmImportaListaDePrecos.add_endRequest(inicializacaoImportaListaDePrecos);
});

function inicializacaoImportaListaDePrecos() {
    $("#lnkImportar").click(function (e) {
        //e.preventDefault();
        if ($("#lstEmpresa").val() == '') {
            alert("Informe a Empresa.");
            return false;
        }
        else if ($("#filUpload").val() == '') {
            alert("Nenhum Arquivo Selecionado jquery");
            return false;
        }

        return confirm("Tem certeza que deseja Importar os dados?");
    });
};