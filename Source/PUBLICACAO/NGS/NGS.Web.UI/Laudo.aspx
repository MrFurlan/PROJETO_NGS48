<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Laudo.aspx.vb" Inherits="NGS.Web.UI.Laudo" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }

        table.pontilhado {
            border-collapse: collapse;
            width: 100%;
        }

            table.pontilhado > tbody > tr > td {
                border: 1px dotted #000000;
            }

        .none {
            display: none;
        }

        .fonte6 {
            font-size: 0.8em;
        }

        .borda_pontilhado {
            border: 1px solid #BBBBBB;
        }

        .tamanho50 {
            width: 50%;
            position: relative;
        }
    </style>

    <script type="text/javascript">
        function pageLoadLaudo() {
            $('.txt', '#MainContent_TabContainer1_TabPanel1_gridDescontos').on("keydown", function (e) {
                //get the next index of text input element
                var next_idx = $('.txt', '#MainContent_TabContainer1_TabPanel1_gridDescontos').index(this) + 1;
                //get number of text input element in a html document
                var tot_idx = $('#MainContent_TabContainer1_TabPanel1_gridDescontos').find('.txt').length;
                //enter button in ASCII code
                if (e.keyCode == 13) {
                    if (next_idx == ($("#MainContent_TabContainer1_TabPanel1_gridDescontos > tbody > tr").size() - 1)) {
                        //go to the next button input element
                        console.log("MainContent_TabContainer1_TabPanel1_btnCalcular");
                        $("#MainContent_TabContainer1_TabPanel1_btnCalcular").click();
                    }
                    else if (tot_idx == next_idx) {
                        //go to the first text element if focused in the last text input element
                        $('.txt:eq(0)', '#MainContent_TabContainer1_TabPanel1_gridDescontos').focus();
                    }
                    else {
                        //go to the next text input element
                        $('.txt:eq(' + next_idx + ')', '#MainContent_TabContainer1_TabPanel1_gridDescontos').focus();
                    }
                    return false;
                }
            });

            $("#MainContent_TabContainer1_TabPanel1_txtPrimeiraPesagem").keydown(function (e) {
                if (e.keyCode == 13) {
                    $("#MainContent_TabContainer1_TabPanel1_txtSegundaPesagem").val("").focus();

                    $('select', '#MainContent_TabContainer1_TabPanel1_gridDescontos').each(function () {
                        $(this).removeClass("aspNetDisabled");
                        $(this).removeAttr("disabled");
                    });
                }
            });

            $("#MainContent_TabContainer1_TabPanel1_txtSegundaPesagem").keydown(function (e) {
                if (e.keyCode == 13) {

                    if ($("#MainContent_TabContainer1_TabPanel1_txtSegundaPesagem").val() != "" && $("#MainContent_TabContainer1_TabPanel1_txtSegundaPesagem").val() != undefined && parseInt($("#MainContent_TabContainer1_TabPanel1_txtSegundaPesagem").val()) > 0) {
                        $('input[type=text]', '#MainContent_TabContainer1_TabPanel1_gridDescontos').each(function () {
                            $(this).removeAttr("disabled");
                            $(this).removeAttr("readonly");
                        });

                        $('select', '#MainContent_TabContainer1_TabPanel1_gridDescontos').each(function () {
                            $(this).removeClass("aspNetDisabled");
                            $(this).removeAttr("disabled");
                        });

                        var primeiraPesagem = 0;
                        if ($("#MainContent_TabContainer1_TabPanel1_txtPrimeiraPesagem").val() != "" && $("#MainContent_TabContainer1_TabPanel1_txtPrimeiraPesagem").val() != undefined)
                            primeiraPesagem = parseInt($("#MainContent_TabContainer1_TabPanel1_txtPrimeiraPesagem").val());
                        var segundaPesagem = 0;
                        if ($("#MainContent_TabContainer1_TabPanel1_txtSegundaPesagem").val() != "" && $("#MainContent_TabContainer1_TabPanel1_txtSegundaPesagem").val() != undefined)
                            segundaPesagem = parseInt($("#MainContent_TabContainer1_TabPanel1_txtSegundaPesagem").val());

                        if (primeiraPesagem > segundaPesagem) {
                            $("#MainContent_TabContainer1_TabPanel1_txtPesoBruto").val(primeiraPesagem - segundaPesagem);
                        } else {
                            $("#MainContent_TabContainer1_TabPanel1_txtPesoBruto").val(segundaPesagem - primeiraPesagem);
                        }

                        var pesoBruto = parseInt($("#MainContent_TabContainer1_TabPanel1_txtPesoBruto").val());
                        var desconto = parseInt($("#MainContent_TabContainer1_TabPanel1_txtDesconto").val());

                        $("#MainContent_TabContainer1_TabPanel1_txtLiquido").val(pesoBruto - desconto);
                        $("#MainContent_TabContainer1_TabPanel1_btnCalcular").removeAttr("disabled aria-disabled").removeClass("aspNetDisabled ui-button-disabled ui-state-disabled");
                        $("#MainContent_TabContainer1_TabPanel1_btnCalcular").attr("onclick", "javascript:__doPostBack('ctl00$MainContent$TabContainer1$TabPanel1$btnCalcular','')");

                        $("#btnAutorizacao").click();
                        //                        $("##MainContent_TabContainer1_TabPanel1_btnAutorizacao").click();
                        //                        $('input[type=text].txtDecimal:eq(0)', '#MainContent_TabContainer1_TabPanel1_gridDescontos').focus();
                    } else {
                        $('input[type=text]', '#MainContent_TabContainer1_TabPanel1_gridDescontos').each(function () {
                            console.log("2");
                            $(this).val("0,00").attr("disabled", true);
                        });

                        var primeiraPesagem = 0;
                        if ($("#MainContent_TabContainer1_TabPanel1_txtPrimeiraPesagem").val() != "" && $("#MainContent_TabContainer1_TabPanel1_txtPrimeiraPesagem").val() != undefined)
                            primeiraPesagem = parseInt($("#MainContent_TabContainer1_TabPanel1_txtPrimeiraPesagem").val());

                        $("#MainContent_TabContainer1_TabPanel1_txtPesoBruto").removeAttr("readonly");
                        $("#MainContent_TabContainer1_TabPanel1_txtDesconto").removeAttr("readonly");
                        $("#MainContent_TabContainer1_TabPanel1_txtLiquido").removeAttr("readonly");

                        $("#MainContent_TabContainer1_TabPanel1_txtPesoBruto").attr("disabled", true).attr("value", primeiraPesagem);
                        $("#MainContent_TabContainer1_TabPanel1_txtLiquido").attr("disabled", true).attr("value", primeiraPesagem);
                        $("#MainContent_TabContainer1_TabPanel1_txtDesconto").attr("disabled", true).attr("value", 0);

                        $("#MainContent_TabContainer1_TabPanel1_btnCalcular").attr("disabled", true).addClass("aspNetDisabled ui-button-disabled ui-state-disabled");
                        $("#MainContent_TabContainer1_TabPanel1_btnIncluir").removeAttr("disabled aria-disabled").removeClass("aspNetDisabled ui-button-disabled ui-state-disabled");
                        $("#MainContent_TabContainer1_TabPanel1_btnIncluir").attr("onclick", "javascript:__doPostBack('ctl00$MainContent$TabContainer1$TabPanel1$btnIncluir','')");
                        $("#MainContent_TabContainer1_TabPanel1_txtNotaProdutor").val("").focus();
                    }
                    return false;
                }
            });

            $("#MainContent_TabContainer1_TabPanel1_txtNotaProdutor").keydown(function (e) {
                if (e.keyCode == 13) {
                    $("#MainContent_TabContainer1_TabPanel1_txtSerieNotaProdutor").val("").focus();
                    return false;
                }
            });

            $("#MainContent_TabContainer1_TabPanel1_txtSerieNotaProdutor").keydown(function (e) {
                if (e.keyCode == 13) {
                    $("#MainContent_TabContainer1_TabPanel1_txtPesoNotaProdutor").val("").focus();
                    return false;
                }
            });

            $("#MainContent_TabContainer1_TabPanel1_txtPesoNotaProdutor").keydown(function (e) {
                if (e.keyCode == 13) {
                    $("#MainContent_TabContainer1_TabPanel1_txtObservacao").val("").focus();
                    return false;
                }
            });

            $("#txtLaudo").keydown(function (e) {
                if (e.keyCode == 13) {
                    $("#MainContent_TabContainer1_TabPanel1_txtNotaProdutor").focus();
                }
            });
        }

        function ConfirmarPeso(primeiraPesagem, peso) {
            if (primeiraPesagem == "1") {
                $("#dialog").text("Primeira pesagem: " + peso)
            } else {
                $("#dialog").text("Segunda pesagem: " + peso)
            }
            //abre a janela para confirmar
            $("#dialog").dialog({
                modal: true, title: 'Confirma a pesagem?', zIndex: 10000, autoOpen: true,
                width: '260', resizable: false, closeOnEscape: false,
                buttons: {
                    "Sim": function () {
                        $(this).dialog("destroy");
                        $("#btnAutorizacao").click();
                    },
                    "Não": function () {
                        $(this).dialog("close");
                        $("#btnLerPeso").click();
                    }
                },
                open: function (type, data) {
                    $('.ui-dialog-titlebar-close').hide();
                }
            });
        }

        $(document).ready(function () {
            pageLoadLaudo();
            $("#MainContent_TabContainer1_TabPanel1_gridDescontos_txtPercentual_0").focus();
            var prmLaudo = Sys.WebForms.PageRequestManager.getInstance();
            prmLaudo.add_endRequest(pageLoadLaudo);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngsLaudoDePesagem" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlLaudoDePesagem" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Pesagem
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Pesagem
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div style="float: left; width: 50%">
                            <div style="width: 100%;">
                                <fieldset class="borda_pontilhado">
                                    <legend class="fonte6"><strong><em>EMPRESA EMITENTE</em></strong></legend>
                                    <div style="float: left; width: 60%">
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">FILIAL</span></strong></label><br />
                                        <asp:Button ID="btnEmpresa" runat="server" UseSubmitBehavior="False" Text=" > " OnClick="btnEmpresa_Click"
                                            Enabled="True"></asp:Button>
                                        <asp:Label ID="txtNomeDaEmpresa" runat="server" Font-Bold="False" ForeColor="Blue"
                                            Font-Italic="False" Font-Names="Tahoma" />
                                    </div>
                                    <div style="float: left; width: 20%;">
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">CNPJ</span> </strong>
                                        </label>
                                        <br />
                                        <asp:Label ID="txtCnpjDaEmpresa" runat="server" ForeColor="Blue" Font-Italic="False"
                                            Font-Names="Tahoma" />
                                    </div>
                                    <div style="float: left; width: 15%">
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">MUNICÍPIO</span></strong></label><br />
                                        <asp:Label ID="txtMunicipioDaEmpresa" runat="server" ForeColor="Blue" Font-Italic="False"
                                            Font-Names="Tahoma" />
                                    </div>
                                    <div style="float: left; width: 5%">
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">UF</span></strong></label><br />
                                        <asp:Label ID="txtEstadoDaEmpresa" runat="server" ForeColor="Blue" Font-Italic="False"
                                            Font-Names="Tahoma" />
                                    </div>
                                    <asp:HiddenField ID="hdnCodigoEndEmpresa" runat="server" />
                                </fieldset>
                            </div>
                            <div style="width: 100%">
                                <fieldset class="borda_pontilhado">
                                    <legend class="fonte6"><strong><em>LAUDO</em></strong></legend>
                                    <div style="float: left; width: 20%">
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">LAUDO</span></strong></label><br />
                                        <asp:TextBox ID="txtLaudo" runat="server" Style="color: Blue; text-align: right"
                                            Font-Names="Arial" Font-Size="15pt" Width="100px" CssClass="borda_pontilhado"
                                            OnTextChanged="txtLaudo_TextChanged" AutoPostBack="True" ClientIDMode="Static" />&nbsp;
                                        <asp:Label ID="lblEntradaSaida" runat="server" Font-Bold="True" Font-Size="Small"
                                            ForeColor="Red" />
                                    </div>
                                    <div style="float: left; width: 20%">
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">ROMANEIO</span></strong></label><br />
                                        <asp:TextBox ID="txtRomaneio" runat="server" Style="color: Blue; text-align: right"
                                            Font-Names="Arial" Font-Size="15pt" Width="100px" CssClass="borda_pontilhado"
                                            ReadOnly="True" Enabled="False" ClientIDMode="Static" />
                                    </div>
                                    <div style="float: left; width: 60%">
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">SITUAÇÃO</span></strong></label><br />
                                        <asp:DropDownList ID="ddlSituacao" runat="server" Width="230px" Enabled="False">
                                        </asp:DropDownList>
                                    </div>
                                </fieldset>
                            </div>
                            <div style="width: 100%">
                                <fieldset class="borda_pontilhado">
                                    <legend class="fonte6"><strong><em>CLIENTE</em></strong></legend>
                                    <div style="float: left; width: 100%">
                                        <div style="float: left; width: 60%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">
                                                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente." data-ToolTip="default"
                                                        runat="server" />NOME/RAZÃO SOCIAL</span></strong></label><br />
                                            <asp:Button ID="btnCliente" runat="server" Text=" > " UseSubmitBehavior="False" />
                                            <asp:Label ID="txtNomeDoCliente" Font-Names="Tahoma" runat="server" />
                                        </div>
                                        <div style="float: left; width: 20%;">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">CPF/CNPJ</span></strong></label><br />
                                            <asp:Label ID="txtCnpjDoCliente" Font-Names="Tahoma" runat="server" />
                                            <asp:HiddenField ID="hdnCodigoEndCliente" runat="server" />
                                        </div>
                                        <div style="float: left; width: 15%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">MUNICÍPIO</span></strong></label><br />
                                            <asp:Label ID="txtMunicipioDoCliente" Font-Names="Tahoma" runat="server" />
                                        </div>
                                        <div style="float: left; width: 5%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">UF</span></strong></label><br />
                                            <asp:Label ID="txtEstadoDoCliente" Font-Names="Tahoma" runat="server" />
                                        </div>
                                    </div>
                                    <div style="float: left; width: 100%">
                                        <div style="float: left; width: 50%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label><br />
                                            <asp:Label ID="txtEnderecoDoCliente" Font-Names="Tahoma" runat="server" />
                                        </div>
                                        <div style="float: left; width: 20%;">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label><br />
                                            <asp:Label ID="txtComplementoDoCliente" Font-Names="Tahoma" runat="server" />
                                        </div>
                                        <div style="float: left; width: 15%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">BAIRRO/DISTRITO</span></strong></label><br />
                                            <asp:Label ID="txtBairroDoCliente" Font-Names="Tahoma" runat="server" />
                                        </div>
                                        <div style="float: left; width: 10%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">CEP</span></strong></label><br />
                                            <asp:Label ID="txtCepDoCliente" Font-Names="Tahoma" runat="server" />
                                        </div>
                                    </div>
                                    <div style="float: left; width: 100%">
                                        <div style="float: left; width: 60%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">FONE/FAX</span></strong></label><br />
                                            <asp:Label ID="txtTelefoneDoCliente" Font-Names="Tahoma" runat="server" />
                                        </div>
                                        <div style="float: left; width: 40%;">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label><br />
                                            <asp:Label ID="txtInscricaoDoCliente" Font-Names="Tahoma" runat="server" />
                                        </div>
                                    </div>
                                </fieldset>
                            </div>
                            <div style="width: 100%">
                                <fieldset class="borda_pontilhado">
                                    <legend class="fonte6"><strong><em>PEDIDO</em></strong></legend>
                                    <div style="float: left; width: 100%">
                                        <div style="float: left; width: 20%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">NÚMERO</span></strong></label><br />
                                            <asp:Button ID="btnPedido" runat="server" Text=" > " UseSubmitBehavior="False" />&nbsp;
                                            <asp:Label ID="txtPedido" Font-Names="Tahoma" runat="server" />
                                            <asp:ImageButton ID="imgExtratoPedido" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                ImageAlign="AbsMiddle" data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido"
                                                Visible="False" OnClick="imgExtratoPedido_Click" Style="border: 0px none;" CausesValidation="False"></asp:ImageButton>&nbsp;
                                            <asp:HiddenField ID="hdnSaldoPedido" Value="0" runat="server" />
                                            <asp:HiddenField ID="hdnAutorizacaoDeRetirada" Value="0" runat="server" />
                                            <asp:Button ID="btnRateio" runat="server" Text="Rateio" CssClass="botao" Visible="False"
                                                Width="60px" UseSubmitBehavior="False" />&nbsp;
                                            <asp:ImageButton ID="imgRateio" Visible="False" runat="server" Height="30px" ImageUrl="~/Images/compartilhar.jpg"
                                                data-ToolTip="default" ToolTip="Rateio" Enabled="False" />
                                        </div>
                                        <div style="float: left; width: 40%;">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">OPERAÇÃO</span></strong></label><br />
                                            <asp:DropDownList ID="ddlOperacao" runat="server" Width="90%" AutoPostBack="True">
                                            </asp:DropDownList>
                                        </div>
                                        <div style="float: left; width: 40%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">SUB OPERAÇÃO</span></strong></label><br />
                                            <asp:DropDownList ID="ddlSubOperacao" runat="server" Width="90%" AutoPostBack="True">
                                            </asp:DropDownList>
                                        </div>
                                    </div>
                                    <div style="float: left; width: 100%">
                                        <div style="float: left; width: 40%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">PRODUTO</span></strong></label><br />
                                            <asp:Label ID="txtProduto" Font-Names="Tahoma" runat="server" />
                                        </div>
                                        <div style="float: left; width: 40%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">SAFRA</span></strong></label><br />
                                            <asp:Label ID="txtSafra" Font-Names="Tahoma" runat="server" />
                                        </div>
                                        <div style="float: left; width: 20%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt; color: red">SALDO PEDIDO</span></strong></label><br />
                                            <asp:Label ID="txtSaldoPedido" Font-Names="Tahoma" CssClass="color: red" runat="server" />
                                        </div>
                                    </div>
                                    <div style="float: left; width: 100%">
                                        <asp:Panel ID="pnlRateio" runat="server" Width="99%" Height="220px" BorderStyle="Double"
                                            BorderColor="#336699" Visible="False">
                                            <table width="99%" border="0">
                                                <tr>
                                                    <td>
                                                        <strong><em style="color: red">RATEIO</em></strong>
                                                        <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                                            <tr>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                                                    &nbsp;<asp:Button ID="btnDestinatario" runat="server" Text=" > " UseSubmitBehavior="False" />
                                                                    <asp:Label ID="txtNomeDoDestinatario" runat="server" />
                                                                </td>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt">&nbsp;CPF/CNPJ</span></strong></label><br />
                                                                    &nbsp;<asp:Label ID="txtCnpjDoDestinatario" runat="server" />
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                                            <tr>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt">&nbsp;ENDEREÇO</span></strong></label><br />
                                                                    &nbsp;<asp:Label ID="txtEnderecoDoDestinatario" runat="server" />
                                                                </td>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt">&nbsp;COMPLEMENTO</span></strong></label><br />
                                                                    &nbsp;<asp:Label ID="txtComplementoDoDestinatario" runat="server" />
                                                                </td>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt">&nbsp;BAIRRO/DISTRITO</span></strong></label><br />
                                                                    &nbsp;<asp:Label ID="txtBairroDoDestinatario" runat="server" />
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                                            <tr>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label><br />
                                                                    &nbsp;<asp:Label ID="txtMunicipioDoDestinatario" runat="server" />
                                                                </td>
                                                                <td>
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label><br />
                                                                    &nbsp;<asp:Label ID="txtEstadoDoDestinatario" runat="server" />
                                                                </td>
                                                                <td style="background-color: #ff9966">
                                                                    <label class="titulo">
                                                                        <strong><span style="font-size: 6pt">&nbsp;QUANTIDADE</span></strong></label><br />
                                                                    &nbsp;<asp:TextBox ID="txtQuantidadeDestino" runat="server" Style="color: Blue; text-align: right"
                                                                        Font-Names="Arial" Font-Size="15pt" Width="100px" BorderStyle="None" ReadOnly="True" />
                                                                </td>
                                                                <td align="center">
                                                                    <asp:ImageButton ID="imgQuantidadeRateio" runat="server" ImageUrl="~/images/confirmar.gif" />
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                                        </asp:Panel>
                                    </div>
                                    <div style="float: left; width: 100%">
                                        <asp:Panel ID="pnlDestinatarios" Visible="False" runat="server" Width="98%" Height="80px"
                                            BorderStyle="Double" BorderColor="#336699" ScrollBars="Vertical">
                                            <asp:GridView ID="GridRateio" runat="server" AutoGenerateColumns="False" BackColor="White"
                                                BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3" Width="99%">
                                                <RowStyle ForeColor="#000066" />
                                                <Columns>
                                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente" />
                                                    <asp:BoundField DataField="Endereco" HeaderText="End" />
                                                    <asp:BoundField DataField="Nome" HeaderText="Nome" />
                                                    <asp:BoundField DataField="Pedido" HeaderText="Pedido" />
                                                    <asp:BoundField DataField="Operacao" HeaderText="OP" />
                                                    <asp:BoundField DataField="SubOperacao" HeaderText="SO" />
                                                    <asp:BoundField DataField="Autorizacao" HeaderText="Autoriza&#231;&#227;o">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Bruto" DataFormatString="{0:N0}" HeaderText="Bruto">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Desconto" DataFormatString="{0:N0}" HeaderText="Desconto">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Liquido" DataFormatString="{0:N0}" HeaderText="Liquido">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                </Columns>
                                                <FooterStyle BackColor="White" ForeColor="#000066" />
                                                <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                                                <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                                                <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                                            </asp:GridView>
                                        </asp:Panel>
                                    </div>
                                </fieldset>
                            </div>
                            <div style="width: 100%;">
                                <fieldset class="borda_pontilhado">
                                    <legend class="fonte6"><strong><em>DEPOSITANTE</em></strong></legend>
                                    <asp:Button ID="btnDepositante" runat="server" Text=" > " UseSubmitBehavior="False" />
                                    <asp:Label ID="txtDepositante" Font-Names="Tahoma" runat="server" />
                                    <asp:HiddenField ID="hdnCodigoEndDepositante" runat="server" />
                                </fieldset>
                            </div>
                            <div style="width: 100%;">
                                <fieldset class="borda_pontilhado">
                                    <legend class="fonte6"><strong><em>DEPÓSITO</em></strong></legend>
                                    <asp:Button ID="btnDeposito" runat="server" Text=" > " UseSubmitBehavior="False" />
                                    <asp:Label ID="txtDeposito" Font-Names="Tahoma" runat="server" />
                                    <asp:HiddenField ID="hdnCodigoEndDeposito" runat="server" />
                                </fieldset>
                            </div>
                            <div style="width: 100%;">
                                <fieldset class="borda_pontilhado">
                                    <legend class="fonte6"><strong><em>DADOS DO TRANSPORTADOR</em></strong></legend>
                                    <div style="float: left; width: 100%">
                                        <div style="float: left; width: 50%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">TRANSPORTADOR</span></strong></label><br />
                                            <asp:Button ID="btnTransportador" runat="server" Text=" > " UseSubmitBehavior="False" />
                                            <asp:Label ID="txtTransportador" Font-Names="Tahoma" runat="server" />
                                            <asp:HiddenField ID="hdnCodigoEndTransportador" runat="server" />
                                        </div>
                                        <div style="float: left; width: 20%;">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">PLACA</span></strong></label><br />
                                            <asp:Button ID="btnPlaca" runat="server" Text=" > " UseSubmitBehavior="False" />
                                            <asp:Label ID="txtPlaca" Font-Names="Tahoma" runat="server" />
                                        </div>
                                        <div style="float: left; width: 30%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">MOTORISTA</span></strong></label><br />
                                            <asp:Label ID="txtMotorista" Font-Names="Tahoma" runat="server" />
                                            <asp:HiddenField ID="hdnCodigoEndMotorista" runat="server" />
                                            <asp:HiddenField ID="hdnViaDeTransporte" runat="server" />
                                        </div>
                                        <div style="float: left; width: 5%">
                                        </div>
                                    </div>
                                </fieldset>
                            </div>
                            <div style="width: 100%; padding: 10px 10px 10px 10px">
                                <asp:Button ID="btnPesagem" runat="server" Text="Pesagem" CssClass="botao" Enabled="False"
                                    UseSubmitBehavior="False" />
                                <asp:Button ID="btnEncerrar" runat="server" CssClass="botao" Text="Encerrar" Enabled="False"
                                    UseSubmitBehavior="False" />
                                <asp:Button ID="btnIncluir" runat="server" Text="Gravar" CssClass="botao" Enabled="False"
                                    UseSubmitBehavior="False" OnClick="btnIncluir_Click" />
                                <asp:Button ID="btnLimpar" runat="server" CssClass="botao" Enabled="False" Text="Limpar"
                                    UseSubmitBehavior="False" />
                                <asp:Button ID="btnCancelar" runat="server" CssClass="botao" Enabled="False" Text="Cancelar"
                                    OnClientClick="if (!confirm('Deseja realmente cancelar o laudo?')) return false;" UseSubmitBehavior="False" />
                                <asp:Button ID="btnReimprimir" runat="server" CssClass="botao" Enabled="False" Text="Reimprimir"
                                    UseSubmitBehavior="False" />
                            </div>
                        </div>
                        <div style="float: right; width: 50%">
                            <div style="width: 100%; margin-bottom: 0.3%">
                                <div style="width: 30%; float: right;">
                                    <fieldset class="borda_pontilhado">
                                        <legend class="fonte6"><strong><em>MOVIMENTO</em></strong></legend>
                                        <asp:TextBox ID="txtMovimento" runat="server" CssClass="calendario borda_pontilhado"
                                            ClientIDMode="Static" /><br />
                                        <br />
                                        <asp:Button ID="btnConsultar" runat="server" CssClass="botao" Enabled="False" Text="Consultar"
                                            UseSubmitBehavior="False" />
                                    </fieldset>
                                </div>
                                <div style="width: 70%;">
                                    <fieldset class="borda_pontilhado">
                                        <legend class="fonte6"><strong><em>INFORMAÇÕES PARA IMPRESSÃO</em></strong></legend>
                                        <div style="float: left; width: 40%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">&nbsp;TIPO DE IMPRESSÃO</span></strong></label><br />
                                            <asp:RadioButton ID="radImpressora" runat="server" GroupName="Impressao" Text="Impressora"
                                                AutoPostBack="True" />
                                            <asp:RadioButton ID="radArquivo" runat="server" GroupName="Impressao" Text="Arquivo"
                                                Checked="True" AutoPostBack="True" />
                                        </div>
                                        <div style="float: left; width: 30%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">NÚMERO DE CÓPIAS</span></strong></label><br />
                                            <asp:TextBox ID="txtNumeroDeCopias" CssClass="txtNumerico" runat="server" Width="20px" />
                                        </div>
                                        <div style="float: left; width: 30%">
                                            <label class="titulo">
                                                <span style="font-size: 6pt; font-weight: bold">ACESSO USUÁRIO:</span></label>
                                            <asp:Label ID="lblAcessoUsuario" runat="server" Width="40px" />
                                        </div>
                                        <div style="float: left; width: 80%">
                                            <asp:DropDownList ID="ddlImpressora" runat="server" Width="70%" Enabled="False">
                                            </asp:DropDownList>
                                        </div>
                                        <div style="float: left; width: 80%">
                                            <table width="99%" border="0">
                                                <tr>
                                                    <td>
                                                        <label class="titulo">
                                                            <strong><span style="font-size: 8pt;">Módulo:</span></strong></label><asp:Label ID="lblModulo"
                                                                runat="server" Width="30px" />
                                                    </td>
                                                    <td>
                                                        <label class="titulo">
                                                            <strong><span style="font-size: 8pt">&nbsp;Porta:</span></strong></label>
                                                        <asp:Label ID="lblPorta" runat="server" Width="30px" />
                                                    </td>
                                                    <td>
                                                        <label class="titulo">
                                                            <strong><span style="font-size: 8pt">&nbsp;Acesso Bal.:</span></strong></label>
                                                        <asp:Label ID="lblAcesso" runat="server" />
                                                    </td>
                                                    <td>
                                                        <asp:Button ID="btnTesteBalanca" runat="server" CssClass="botao" Text="Ler Peso"
                                                            UseSubmitBehavior="False" />
                                                        <asp:HiddenField ID="hdnBaudRateBalanca" runat="server" />
                                                        <asp:HiddenField ID="hdnDataBits" runat="server" />
                                                        <asp:HiddenField ID="hdnParity" runat="server" />
                                                        <asp:HiddenField ID="hdnStopBits" runat="server" />
                                                        <asp:HiddenField ID="hdnEletronico" runat="server" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </div>
                                    </fieldset>
                                </div>
                            </div>
                            <div style="width: 100%;">
                                <fieldset class="borda_pontilhado">
                                    <legend class="fonte6"><strong><em>ANÁLISES</em></strong></legend>
                                    <label class="titulo">
                                        <strong><span style="font-size: 6pt">TABELA DE CLASSIFICAÇÕES</span></strong></label><br />
                                    <div style="float: left; width: 50%">
                                        <asp:DropDownList ID="ddlTabelaDeClassificacao" runat="server" AutoPostBack="True"
                                            Width="300px">
                                        </asp:DropDownList>
                                    </div>
                                    <div style="float: right; width: 50%">
                                        &nbsp;<asp:Label ID="lblPesagemMecanica" runat="server" CssClass="texto_vermelho"
                                            Text="Pesagem Eletrônica" />
                                        &nbsp;<asp:CheckBox ID="chkLiberado" runat="server" Enabled="False" Text="Liberado" />
                                    </div>
                                    <div class="bordaLaudoAnalises" id="divAnalise" runat="server">
                                        <asp:GridView ID="gridDescontos" runat="server" CellPadding="4" ForeColor="#333333"
                                            GridLines="None" Width="98%" AutoGenerateColumns="False">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EditRowStyle BackColor="#999999" />
                                            <Columns>
                                                <asp:BoundField DataField="CodigoAnalise" />
                                                <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o" />
                                                <asp:TemplateField HeaderText="Percentual">
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtPercentual" runat="server" Width="98%" CssClass="txtDecimal txt"
                                                            Enabled="false" Text='<%# IIf(Convert.ToString(Eval("Percentual")) = "0", "",  Eval("Percentual", "{0:N2}"))%>' />
                                                        <asp:DropDownList ID="ddlOpcao" runat="server" Visible="False" Width="98%" CssClass="txt">
                                                        </asp:DropDownList>
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Indice">
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtIndice" runat="server" Width="98%" Enabled="false" Text='<%# Eval("Indice")%>' />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Desconto">
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtDesconto" runat="server" Width="98%" Enabled="false" Text='<%# Eval("Desconto")%>' />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                            </Columns>
                                        </asp:GridView>


                                        <asp:Button ID="btnCalcular" runat="server" CssClass="botao" Text="Calcular" Enabled="False"
                                            UseSubmitBehavior="False" />
                                    </div>
                                    <div style="float: left; width: 27%">
                                        <div style="float: left; width: 90%;">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">1ª. PESAGEM</span></strong></label><br />
                                            <asp:TextBox ID="txtPrimeiraPesagem" runat="server" Style="color: Blue; text-align: right;"
                                                Font-Names="Arial" Font-Size="12pt" Width="80px" BorderStyle="None" CssClass="txtNumerico"
                                                onkeydown="return (event.keyCode!=13);" Enabled="False" />
                                        </div>
                                        <div style="float: left; width: 90%;">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">2ª. PESAGEM</span></strong></label><br />
                                            <asp:TextBox ID="txtSegundaPesagem" runat="server" Style="color: Blue; text-align: right;"
                                                Font-Names="Arial" Font-Size="12pt" Width="80px" BorderStyle="None" CssClass="txtNumerico"
                                                Enabled="False" />
                                        </div>
                                        <div style="float: left; width: 90%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">PESO BRUTO</span></strong></label><br />
                                            <asp:TextBox ID="txtPesoBruto" runat="server" Style="color: Black; text-align: right"
                                                Font-Names="Arial" Font-Size="12pt" Width="80px" BorderStyle="None" Enabled="False" />
                                        </div>
                                        <div style="float: left; width: 90%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">DESCONTO</span></strong></label><br />
                                            <asp:TextBox ID="txtDesconto" runat="server" Style="color: Black; text-align: right"
                                                Font-Names="Arial" Font-Size="12pt" Width="80px" BorderStyle="None" Enabled="False" />
                                        </div>
                                        <div style="float: left; width: 90%">
                                            <label class="titulo">
                                                <strong><span style="font-size: 6pt">LIQUIDO</span></strong></label><br />
                                            <asp:TextBox ID="txtLiquido" runat="server" Style="color: Red; text-align: right"
                                                Font-Names="Arial" Font-Size="12pt" Width="80px" BorderStyle="None" Enabled="False" />
                                        </div>
                                        <div style="float: left; width: 90%">
                                            <asp:Button ID="btnLerPeso" OnClick="btnLerPeso_Click" Width="120px" runat="server"
                                                CssClass="botao none" Text="Capturar Peso" UseSubmitBehavior="False" ClientIDMode="Static" />
                                            <asp:Button ID="btnAutorizacao" OnClick="btnAutorizacao_Click" runat="server" CssClass="botao none"
                                                Text="Autorizacao" UseSubmitBehavior="False" ClientIDMode="Static" />
                                        </div>
                                    </div>
                                </fieldset>
                            </div>
                            <div style="width: 100%; clear: both;">
                                <fieldset class="borda_pontilhado">
                                    <legend class="fonte6"><strong><em>NOTA FISCAL PRODUTOR</em></strong></legend>

                                    <div style="float: left; width: 23%">
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">NOTA FISCAL/SÉRIE</span></strong></label><br />
                                        <asp:TextBox ID="txtNotaProdutor" CssClass="borda_pontilhado txtNumerico" runat="server"
                                            Style="text-align: right" Width="80px" TabIndex="1" />
                                        -
                                        <asp:TextBox ID="txtSerieNotaProdutor" CssClass="borda_pontilhado" runat="server"
                                            MaxLength="3" Width="30px" TabIndex="2" />
                                    </div>
                                    <div style="float: left;">
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">PESO</span></strong></label><br />
                                        <asp:TextBox ID="txtPesoNotaProdutor" CssClass="borda_pontilhado txtNumerico" runat="server"
                                            Style="text-align: right" Width="80px" TabIndex="3" />
                                    </div>
                                    <div style="float: left; width: 62%; margin-left: 4px;">
                                    </div>
                                    <legend class="fonte6"><strong><em>OBSERVAÇÃO</em></strong></legend>
                                    <asp:TextBox ID="txtObservacao" runat="server" CssClass="borda_pontilhado" Font-Names="Tahoma" Height="30px" TabIndex="4" TextMode="MultiLine" Width="99%"></asp:TextBox>
                                </fieldset>
                            </div>
                            </fieldset>
                            <br />
                        </div>
                        <div style="width: 100%;">
                            <fieldset class="borda_pontilhado">
                                <legend class="fonte6"><strong><em>HISTÓRICO</em></strong></legend>
                                <asp:Label ID="lblHistorico" runat="server" Font-Names="Tahoma"></asp:Label>
                            </fieldset>
                        </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="gdvLaudo" runat="server" CellPadding="4" ForeColor="#333333" GridLines="None"
                                Width="98%" AutoGenerateColumns="False" OnSelectedIndexChanged="Gridlaudo_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Font-Names="Tahoma" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Font-Names="Tahoma" />
                                <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" Font-Names="Tahoma" />
                                <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" Font-Names="Tahoma" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:BoundField DataField="Codigo" HeaderText="Laudo" />
                                    <asp:BoundField DataField="CodigoEmpresa">
                                        <HeaderStyle CssClass="none" />
                                        <ItemStyle CssClass="none" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EnderecoEmpresa">
                                        <HeaderStyle CssClass="none" />
                                        <ItemStyle CssClass="none" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoRomaneio" HeaderText="Romaneio" />
                                    <asp:BoundField DataField="CodigoPlaca" HeaderText="Placa" />
                                    <asp:TemplateField HeaderText="Empresa">
                                        <ItemTemplate>
                                            <%# Eval("CodigoEmpresa")%>&nbsp;-&nbsp;<%# Eval("Empresa.Nome")%>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Cliente">
                                        <ItemTemplate>
                                            <%# Eval("CodigoCliente")%>&nbsp;-&nbsp;<%# Eval("Cliente.Nome")%>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Produto">
                                        <ItemTemplate>
                                            <%# Eval("CodigoProduto")%>&nbsp;-&nbsp;<%# Eval("Produto.Descricao")%>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoOperacao" HeaderText="OP" />
                                    <asp:BoundField DataField="CodigoSubOperacao" HeaderText="Sub" />
                                    <asp:BoundField DataField="PrimeiraPesagem" HeaderText="1ª Pesagem" />
                                    <asp:BoundField DataField="SegundaPesagem" HeaderText="2ª Pesagem" />
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgImprimir" runat="server" ImageUrl="~/Images/impressora3.jpg"
                                                Style="border: 0;" data-ToolTip="default" ToolTip="Imprimir" OnClientClick="return confirm('Deseja reimprimir o Laudo?');"
                                                OnClick="imgImprimir_Click" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
            <input type="hidden" runat="server" id="hdnControlePopup" />
        </ContentTemplate>
    </asp:UpdatePanel>
    <div id="dialog" data-tooltip="default" title="Pesagem">
    </div>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaClientesDireto ID="ucConsultaClientesDireto" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:ConsultaAutorizacaoDeRetirada ID="ucConsultaAutorizacaoDeRetirada" runat="server" />
    <uc:ConsultaPlacas ID="ucConsultaPlacas" runat="server" />
    <uc:ConsultaEstados ID="ucConsultaEstados" runat="server" />
    <uc:ConsultaCodMunicipios ID="ucConsultaCodMunicipios" runat="server" />
</asp:Content>
