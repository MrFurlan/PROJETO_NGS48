$(document).ready(function () {
    inicializacaoImportaFolha();

    var prmImportaFolha = Sys.WebForms.PageRequestManager.getInstance();
    prmImportaFolha.add_endRequest(inicializacaoImportaFolha);
});

function inicializacaoImportaFolha() {
    $("#lnkImportar").click(function (e) {
        //e.preventDefault();
        if ($("#lstUnidadeNegocio").val() == '') {
            alert("Informe a unidade de negócio.");
            return false;
        }
        else if ($("#lstEmpresa").val() == '') {
            alert("informe a empresa.");
            return false;
        }
        else if ($("#filUpload").val() == '') {
            alert("Nenhum Arquivo Selecionado jquery");
            return false;
        }

        return confirm("Tem certeza que deseja Importar os dados?");
    });
};


