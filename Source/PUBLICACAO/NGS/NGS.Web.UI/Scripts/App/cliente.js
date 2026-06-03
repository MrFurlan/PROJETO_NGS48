function clienteload() {
    var inputsCEP = $('#MainContent_TabContainer1_TabPanelCadastro_txtEndereco, #MainContent_TabContainer1_TabPanelCadastro_txtBairro, #MainContent_TabContainer1_TabPanelCadastro_txtCidade, #MainContent_TabContainer1_TabPanelCadastro_ddlEstado, #MainContent_TabContainer1_TabPanelCadastro_txtCodMunicipio');
    var inputsRUA = $('#MainContent_TabContainer1_TabPanelCadastro_txtCep, #MainContent_TabContainer1_TabPanelCadastro_txtBairro, #MainContent_TabContainer1_TabPanelCadastro_txtCodMunicipio');
    var validacep = /^[0-9]{8}$/;

    function limpa_formulario_cep(msg) {
        if (msg !== undefined) {
            alert(msg);
        }
        inputsCEP.val('');
    }

    function get(url) {
        $.get(url, function (data) {

            if (!("erro" in data)) {

                if (Object.prototype.toString.call(data) === '[object Array]') {
                    var data = data[0];
                }
                $("#MainContent_TabContainer1_TabPanelCadastro_txtCep").val(data.cep);
                $("#MainContent_TabContainer1_TabPanelCadastro_txtEndereco").val(data.logradouro);
                $("#MainContent_TabContainer1_TabPanelCadastro_txtBairro").val(data.bairro);

                //$("#MainContent_TabContainer1_TabPanelCadastro_txtCidade").removeClass("aspNetDisabled");
                //$("#MainContent_TabContainer1_TabPanelCadastro_txtCidade").attr("disabled", false);
                $("#MainContent_TabContainer1_TabPanelCadastro_txtCidade").val(data.localidade);

                $('#MainContent_TabContainer1_TabPanelCadastro_ddlEstado>option[value=""]').attr('selected', false);
                $('#MainContent_TabContainer1_TabPanelCadastro_ddlEstado>option[value="' + data.uf + '"]').attr('selected', true);

                $("#MainContent_TabContainer1_TabPanelCadastro_txtCodMunicipio").val(data.ibge);
                $("#MainContent_TabContainer1_TabPanelCadastro_txtNumero").focus();

                $('#MainContent_TabContainer1_TabPanelCadastro_ddlPais>option:eq(0)').attr('selected', false);
                $('#MainContent_TabContainer1_TabPanelCadastro_ddlPais>option:eq(1)').attr('selected', true);
                //$.each(data, function (nome, info) {
                //    console.log(nome);
                //    console.log(info);
                //    $('#' + nome).val(nome === 'cep' ? info.replace(/\D/g, '') : info).attr('info', nome === 'cep' ? info.replace(/\D/g, '') : info);
                //});
            } else {
                limpa_formulario_cep("CEP não encontrado.");
            }
        });
    }

    function SelecionaImagem(upld) {
        __doPostBack("", "MostrarImagem");
    };

    if ($("#MainContent_TabContainer1_TabPanelCadastro_imgFicha").val() == undefined)
        $('#MainContent_TabContainer1_TabPanelCadastro_btnConsultaCadastro').show();
    else
        $('#MainContent_TabContainer1_TabPanelCadastro_btnConsultaCadastro').hide();

    // Digitando RUA/CIDADE/MainContent_TabContainer1_TabPanelCadastro_ddlEstado
    //$('#MainContent_TabContainer1_TabPanelCadastro_txtEndereco, #MainContent_TabContainer1_TabPanelCadastro_txtCidade, #MainContent_TabContainer1_TabPanelCadastro_ddlEstado').on('blur', function (e) {
    //    if ($('#MainContent_TabContainer1_TabPanelCadastro_txtEndereco').val() !== '' && $('#MainContent_TabContainer1_TabPanelCadastro_txtEndereco').val() !== $('#MainContent_TabContainer1_TabPanelCadastro_txtEndereco').attr('info') && $('#MainContent_TabContainer1_TabPanelCadastro_txtCidade').val() !== '' && $('#MainContent_TabContainer1_TabPanelCadastro_txtCidade').val() !== $('#MainContent_TabContainer1_TabPanelCadastro_txtCidade').attr('info') && $('#MainContent_TabContainer1_TabPanelCadastro_ddlEstado').val() !== '' && $('#MainContent_TabContainer1_TabPanelCadastro_ddlEstado').val() !== $('#MainContent_TabContainer1_TabPanelCadastro_ddlEstado').attr('info')) {
    //        inputsRUA.val('...');
    //        get('https://viacep.com.br/ws/' + $('#MainContent_TabContainer1_TabPanelCadastro_ddlEstado').val() + '/' + $('#MainContent_TabContainer1_TabPanelCadastro_txtCidade').val() + '/' + $('#MainContent_TabContainer1_TabPanelCadastro_txtEndereco').val() + '/json/');
    //    }
    //});

    // Digitando CEP
    //$('#MainContent_TabContainer1_TabPanelCadastro_txtCep').on('blur', function (e) {
    //    var cep = $('#MainContent_TabContainer1_TabPanelCadastro_txtCep').val().replace(/\D/g, '');
    //    if (cep !== "" && validacep.test(cep)) {
    //        inputsCEP.val('...');
    //        get('https://viacep.com.br/ws/' + cep + '/json/');
    //    } else {
    //        limpa_formulario_cep(cep == "" ? undefined : "Formato de CEP inválido.");
    //    }
    //});
};
