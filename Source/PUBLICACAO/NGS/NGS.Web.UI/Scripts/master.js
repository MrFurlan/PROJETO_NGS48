//<![CDATA[
/*Exibe a mensagem de carga ao abandonar a tela*/
window.onbeforeunload = function () { $("#dvWait:first").show(); };
/*Validação de CPF*/
function isValidCPF(value) { value = $.trim(value); var cpf = value.replace(/\.|-|\//gi, ''); while (cpf.length < 11) cpf = "0" + cpf; var expReg = /^0+$|^1+$|^2+$|^3+$|^4+$|^5+$|^6+$|^7+$|^8+$|^9+$/; var a = []; var b = new Number; var c = 11; for (i = 0; i < 11; i++) { a[i] = cpf.charAt(i); if (i < 9) b += (a[i] * --c); } if ((x = b % 11) < 2) { a[9] = 0 } else { a[9] = 11 - x } b = 0; c = 11; for (y = 0; y < 10; y++) b += (a[y] * c--); if ((x = b % 11) < 2) { a[10] = 0; } else { a[10] = 11 - x; } if ((cpf.charAt(9) != a[9]) || (cpf.charAt(10) != a[10]) || cpf.match(expReg)) return false; return true; }
/*Função para substituir a função alert*/
function jAlert(msg, title, callback) { if (title == null) title = "Aviso"; msg = msg.replace(/\n/g, '<br />'); $("body").append("<div id='dialog-alerts' style='display:none;' title='" + title + "'><p><table><tr><td style='width:10%;'><span class='ui-icon ui-icon-alert' style='float:left; margin:0 7px 20px 0;'></span></td><td>" + msg + "</td></tr></table></p></div> "); $(function () { $("#dialog").dialog("destroy"); $("#dialog-alerts").dialog({ resizable: true, minHeight: 200, maxHeight: 600, width: 700, minWidth: 500, maxWidth: 800, modal: true/*, show: 'scale', hide: 'clip'*/, buttons: { "OK": function () { if (callback) callback(true); $(this).dialog('close'); } }, close: function () { $("#dialog-alerts").dialog("destroy"); $("#dialog-alerts").remove(); } }); }); return false; }
/*Função para substituir a função confirm*/
function jConfirm(msg, title, callback) { if (title == null) title = "Confirmação"; msg = msg.replace(/\n/g, '<br />'); $("body").append("<div id='dialog-alerts' style='display:none;' title='" + title + "'><p><span class='ui-icon ui-icon-info' style='float:left; margin:0 7px 20px 0;'></span>" + msg + "</p></div> "); $(function () { $("#dialog").dialog("destroy"); $("#dialog-alerts").dialog({ resizable: true, minHeight: 150, maxHeight: 600, width: 500, minWidth: 250, maxWidth: 800, modal: true, /*show: 'scale', hide: 'clip',*/buttons: { "Confirmar": function () { if (callback) callback(true); $(this).dialog('close'); }, "Cancelar": function () { if (callback) callback(false); $(this).dialog('close'); } }, close: function () { $("#dialog-alerts").dialog("destroy"); $("#dialog-alerts").remove(); } }); }); return false; }
/*Formata as datas recebidas via JSON ASPNet*/
function formatJSONDate(dateValue) { var result = ""; if (dateValue != null) if (dateValue != "") if (dateValue != "undefined") result = $.datepicker.formatDate('d/mm/yy', new Date(parseInt(eval(dateValue.replace(/\/Date\((\d+)\)\//gi, "$1"))))); return result; }
$(document).ready(function () {
    $.ajaxSetup({ global: false, cache: true, type: "post", dataType: "json", async: true, timeout: 15000 });
    /*Customização do componente DatePicker do jQuery.UI*/
    $(function ($) { $.datepicker.regional['pt-BR'] = { closeText: 'Fechar', prevText: '&#x3c;Anterior', nextText: 'Pr&oacute;ximo&#x3e;', currentText: 'Hoje', monthNames: ['Janeiro', 'Fevereiro', 'Mar&ccedil;o', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'], monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'], dayNames: ['Domingo', 'Segunda-feira', 'Ter&ccedil;a-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sabado'], dayNamesShort: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sab'], dayNamesMin: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sab'], dateFormat: 'dd/mm/yy', firstDay: 0, isRTL: false }; $.datepicker.setDefaults($.datepicker.regional['pt-BR']); });
    /*Configura os campos do tipo data*/
    $(".datefield").datepicker().attr("maxlength", "10").keypress(function (e) { if (isNaN(String.fromCharCode(e.which)) && e.which != 0 && e.which != 8 && e.which != 13) return false; if ($(this).val().length == 2 || $(this).val().length == 5) $(this).val($(this).val() + "/"); }).blur(function () { var day = $(this).val().substr(0, $(this).val().indexOf("/")); var month = $(this).val().substr($(this).val().indexOf("/") + 1, 2); var year = $(this).val().substr($(this).val().lastIndexOf("/") + 1, 4); var date = new Date(); date.setFullYear(year, month - 1, day); if (date.getDate() != day || date.getMonth() != month - 1 || date.getFullYear() != year || parseInt(year) < 1900) $(this).val(""); });
    /*Configura os campos do tipo número inteiro*/
    $(".numberfield").keypress(function (e) { if (isNaN(String.fromCharCode(e.which)) && e.which != 0 && e.which != 8 && e.which != 13) return false; });
    /*Configura os campos do tipo número decimal*/
    $(".decimalfield").keypress(function (e) { if (isNaN(String.fromCharCode(e.which)) && e.which != 0 && e.which != 8 && e.which != 13 && e.which != 44) return false; });
    /*Configura os campos do tipo telefone*/
    $(".phonefield").attr("maxlength", "12").keypress(function (e) { if (isNaN(String.fromCharCode(e.which)) && e.which != 0 && e.which != 8 && e.which != 13) return false; /*document.selection.clear();*/var value = $.trim($(this).val()).replace(" ", ""); if (value.length >= 2) value = $.trim(value.substr(0, 2)) + " " + $.trim(value.substr(2, value.length - 1)); $(this).val(value); });
    $(".phonefield").each(function () { var value = $.trim($(this).val()).replace(" ", ""); if (value.length >= 2) value = $.trim(value.substr(0, 2)) + " " + $.trim(value.substr(2, value.length - 1)); $(this).val(value); });
    /*Configura os campos do tipo hora*/
    $(".timefield").attr("maxlength", "5").keypress(function (e) { $(this).val().replace(":", ""); if (isNaN(String.fromCharCode(e.which)) && e.which != 0 && e.which != 8 && e.which != 13) return false; var value = $.trim($(this).val()).replace(" ", ""); if (value.length >= 2) value = $.trim(value.substr(0, 2)) + ":" + $.trim(value.substr(2, value.length - 1)); $(this).val(value.replace("::", ":")); }).blur(function (e) { $(this).val($(this).val().replace(":", "")); if ($(this).val().length >= 2) { var prefix = $.trim($(this).val().substr(0, 2)); var sufix = $.trim($(this).val().substr(2, $(this).val().length - 1)); if (prefix == "") prefix = "00"; if (sufix == "") sufix = "00"; if (parseInt(prefix) >= 24) prefix = "00"; if (parseInt(sufix) >= 60) sufix = "00"; $(this).val(prefix + ":" + sufix); $(this).val($(this).val().replace("::", ":")); } });
    $(".timefield").each(function () { var value = $.trim($(this).val()).replace(" ", ""); if (value.length >= 2) value = $.trim(value.substr(0, 2)) + ":" + $.trim(value.substr(2, value.length - 1)); $(this).val(value).attr("maxlength", "5"); });
    /*Configura os campos do tipo CPF*/
    $(".cpffield").attr("maxlength", "11").keypress(function (e) { if (isNaN(String.fromCharCode(e.which)) && e.which != 0 && e.which != 8 && e.which != 13) return false; }).blur(function () { var value = $.trim($(this).val()).replace(" ", "").replace(".", "").replace("-", ""); var part1 = ""; var part2 = ""; var part3 = ""; var part4 = ""; if (value.length >= 3) part1 = $.trim(value.substr(0, 3)) + "."; if (value.length >= 4) part2 = $.trim(value.substr(3, 3)) + "."; if (value.length >= 7) part3 = $.trim(value.substr(6, 3)) + "-"; if (value.length >= 9) part4 = $.trim(value.substr(9, value.length - 1)); $(this).val(part1 + part2 + part3 + part4); });
    $("#dvWait").hide();
});
/*Deixa a Tab atual como ativa*/
function SetTab(tab) {
    $(document).ready(function () {
        $(function () {
            $(".tabs").find(".ui-state-active").removeClass("ui-state-active");
            $(".tabs").find("li:eq(" + tab + ") a").click();
        });
    });
}
// ]]>