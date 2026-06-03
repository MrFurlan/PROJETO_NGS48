$(document).ready(function () {
    FiltroPorPedido();

    var prmRelVenda = Sys.WebForms.PageRequestManager.getInstance();
    prmRelVenda.add_endRequest(FiltroPorPedido);
});

function FiltroPorPedido() {
    //filtros de inicializacao
    if ($("#MainContent_rbPorPedido").attr("checked") == "checked") {
        $(".paramporpedido").show();
        $(".paramexercicio").hide();
    }
    else {
        $(".paramporpedido").hide();
        $(".paramexercicio").show();
    }

    //evento click
    $("input[type=radio]").click(function () {
        if ($(this).val() == "rbPorPedido") {
            $(".paramporpedido").show("drop");
            $(".paramexercicio").hide("drop");
        }
        else {
            $(".paramporpedido").hide("drop");
            $(".paramexercicio").show("drop");
        }

        if ($(this).val() == "rbPorCliente") {
            $("#MainContent_lblCliente").html("Cliente:");
        } else {
            $("#MainContent_lblCliente").html("Representante:");
        }
    });
};

