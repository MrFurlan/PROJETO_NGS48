<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AlterarLaudo.aspx.vb" Inherits="NGS.Web.UI.AlterarLaudo" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        input.botao {
            font-size: 11px;
        }
    </style>

    <script type="text/javascript">
        function pageLoadLaudo() {
            $('.txt', '#MainContent_gridDescontos').on("keydown", function (e) {
                //get the next index of text input element
                var next_idx = $('.txt', '#MainContent_gridDescontos').index(this) + 1;
                //get number of text input element in a html document
                var tot_idx = $('#MainContent_gridDescontos').find('.txt').length;
                //enter button in ASCII code
                if (e.keyCode == 13) {
                    if (next_idx == ($("#MainContent_gridDescontos > tbody > tr").size() - 1)) {
                        //go to the next button input element
                        console.log("MainContent_cmdCalcular");
                        $("#MainContent_cmdCalcular").click();
                    }
                    else if (tot_idx == next_idx) {
                        //go to the first text element if focused in the last text input element
                        $('.txt:eq(0)', '#MainContent_gridDescontos').focus();
                    }
                    else {
                        //go to the next text input element
                        $('.txt:eq(' + next_idx + ')', '#MainContent_gridDescontos').focus();
                    }
                    return false;
                }
            });

            $("#MainContent_txtPrimeiraPesagem").keydown(function (e) {
                if (e.keyCode == 13) {
                    $("#MainContent_txtSegundaPesagem").val("").focus();

                    $('select', '#MainContent_gridDescontos').each(function () {
                        $(this).removeClass("aspNetDisabled");
                        $(this).removeAttr("disabled");
                    });
                }
            });

            $("#MainContent_txtSegundaPesagem").keydown(function (e) {
                if (e.keyCode == 13) {

                    if ($("#MainContent_txtSegundaPesagem").val() != "" && $("#MainContent_txtSegundaPesagem").val() != undefined && parseInt($("#MainContent_txtSegundaPesagem").val()) > 0) {
                        $('input[type=text]', '#MainContent_gridDescontos').each(function () {
                            $(this).removeAttr("disabled");
                            $(this).removeAttr("readonly");
                        });

                        $('select', '#MainContent_gridDescontos').each(function () {
                            $(this).removeClass("aspNetDisabled");
                            $(this).removeAttr("disabled");
                        });

                        var primeiraPesagem = 0;
                        if ($("#MainContent_txtPrimeiraPesagem").val() != "" && $("#MainContent_txtPrimeiraPesagem").val() != undefined)
                            primeiraPesagem = parseInt($("#MainContent_txtPrimeiraPesagem").val());
                        var segundaPesagem = 0;
                        if ($("#MainContent_txtSegundaPesagem").val() != "" && $("#MainContent_txtSegundaPesagem").val() != undefined)
                            segundaPesagem = parseInt($("#MainContent_txtSegundaPesagem").val());

                        if (primeiraPesagem > segundaPesagem) {
                            $("#MainContent_txtPesoBruto").val(primeiraPesagem - segundaPesagem);
                        } else {
                            $("#MainContent_txtPesoBruto").val(segundaPesagem - primeiraPesagem);
                        }

                        var pesoBruto = parseInt($("#MainContent_txtPesoBruto").val());
                        var desconto = parseInt($("#MainContent_txtDesconto").val());

                        $("#MainContent_txtLiquido").val(pesoBruto - desconto);
                        $("#MainContent_cmdCalcular").removeAttr("disabled aria-disabled").removeClass("aspNetDisabled ui-button-disabled ui-state-disabled");
                        $("#MainContent_cmdCalcular").attr("onclick", "javascript:__doPostBack('ctl00$MainContent$TabContainer1$TabPanel1$btnCalcular','')");

                        $("#btnAutorizacao").click();
                        //                        $("##MainContent_btnAutorizacao").click();
                        //                        $('input[type=text].txtDecimal:eq(0)', '#MainContent_gridDescontos').focus();
                    } else {
                        $('input[type=text]', '#MainContent_gridDescontos').each(function () {
                            console.log("2");
                            $(this).val("0,00").attr("disabled", true);
                        });

                        var primeiraPesagem = 0;
                        if ($("#MainContent_txtPrimeiraPesagem").val() != "" && $("#MainContent_txtPrimeiraPesagem").val() != undefined)
                            primeiraPesagem = parseInt($("#MainContent_txtPrimeiraPesagem").val());

                        $("#MainContent_txtPesoBruto").removeAttr("readonly");
                        $("#MainContent_txtDesconto").removeAttr("readonly");
                        $("#MainContent_txtLiquido").removeAttr("readonly");

                        $("#MainContent_txtPesoBruto").attr("disabled", true).attr("value", primeiraPesagem);
                        $("#MainContent_txtLiquido").attr("disabled", true).attr("value", primeiraPesagem);
                        $("#MainContent_txtDesconto").attr("disabled", true).attr("value", 0);

                        $("#MainContent_cmdCalcular").attr("disabled", true).addClass("aspNetDisabled ui-button-disabled ui-state-disabled");
                        $("#MainContent_btnIncluir").removeAttr("disabled aria-disabled").removeClass("aspNetDisabled ui-button-disabled ui-state-disabled");
                        $("#MainContent_btnIncluir").attr("onclick", "javascript:__doPostBack('ctl00$MainContent$TabContainer1$TabPanel1$btnIncluir','')");
                        $("#MainContent_txtNotaProdutor").val("").focus();
                    }
                    return false;
                }
            });

            $("#MainContent_txtNotaProdutor").keydown(function (e) {
                if (e.keyCode == 13) {
                    $("#MainContent_txtSerieNotaProdutor").val("").focus();
                    return false;
                }
            });

            $("#MainContent_txtSerieNotaProdutor").keydown(function (e) {
                if (e.keyCode == 13) {
                    $("#MainContent_txtPesoNotaProdutor").val("").focus();
                    return false;
                }
            });

            $("#MainContent_txtPesoNotaProdutor").keydown(function (e) {
                if (e.keyCode == 13) {
                    $("#MainContent_txtObservacao").val("").focus();
                    return false;
                }
            });

            $("#txtLaudo").keydown(function (e) {
                if (e.keyCode == 13) {
                    $("#MainContent_txtNotaProdutor").focus();
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
            $("#MainContent_gridDescontos_txtPercentual_0").focus();
            var prmLaudo = Sys.WebForms.PageRequestManager.getInstance();
            prmLaudo.add_endRequest(pageLoadLaudo);
        });
    </script>
</asp:Content>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngAlterarLaudo" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAlterarLaudo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Alteração de Pesagem
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkVincularNF" Text="Vincular NF" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkDesVincularNF" Text="Desvincular NF" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="618px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged" Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="618px" Enabled="False" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Laudo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtLaudo" runat="server" Width="100px" CssClass="txtNumerico" />
                </div>
                <div class="collbl" style="width: 140px; margin-left: 30px;">
                    Romaneio:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtEntradaSaida" runat="server" />
                    <asp:TextBox ID="TxtRomaneio" runat="server" Enabled="False" Width="126px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Movimento:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtMovimento" CssClass="calendario" runat="server" Width="100px"
                        data-ToolTip="default" ToolTip="Data da movimentação." />
                </div>
                <div class="collbl" style="width: 140px; margin-left: 10px;">
                    Nota Fiscal:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNotaFiscal" runat="server" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultarNota" runat="server" CssClass="btn" Text=" > " Enabled="False" UseSubmitBehavior="False" OnClick="btnConsultarNota_Click" data-ToolTip="default" ToolTip="Buscar Nota Fiscal." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtCliente" runat="server" Width="580px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnConsultarCliente" runat="server" OnClick="BtnConsultarCliente_Click"
                        CssClass="btn" Text=" > " Enabled="False" UseSubmitBehavior="False" data-ToolTip="default"
                        ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtPedido" runat="server" Width="100px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnConsultarPedido" runat="server" OnClick="BtnConsultarPedido_Click"
                        CssClass="btn" Text=" > " Enabled="False" UseSubmitBehavior="False" data-ToolTip="default"
                        ToolTip="Informar o número do pedido." />
                </div>
                <div class="collbl" style="width: 140px;">
                    Autorização de Retirada:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAutorizacao" runat="server" Enabled="False" Width="126px" />
                    <asp:HiddenField ID="txtCodigoAutorizacao" runat="server" />
                    <asp:HiddenField ID="txtSaldoAutorizacao" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnAutorizacao" runat="server" OnClick="btnAutorizacao_Click"
                        CssClass="btn" Text=">" UseSubmitBehavior="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlOperacao" runat="server" Width="463px" Enabled="False" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depositante:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoDepositante" runat="server" />
                    <asp:TextBox ID="TxtDepositante" runat="server" Width="580px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnConsultarDepositante" runat="server" OnClick="BtnConsultarDepositante_Click"
                        CssClass="btn" Text=" > " Enabled="False" UseSubmitBehavior="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoDeposito" runat="server" />
                    <asp:TextBox ID="TxtDeposito" runat="server" Enabled="False" Width="580px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnConsultarDeposito" runat="server" Enabled="False" OnClick="BtnConsultarDeposito_Click"
                        Text=" > " UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Local de armazenamento da mercadoria." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Transportador:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoTransportador" runat="server" />
                    <asp:TextBox ID="TxtTransportador" runat="server" Enabled="False" Width="580px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnConsultarTransportador" runat="server" Enabled="False" OnClick="BtnConsultarTransportador_Click"
                        Text=" > " UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Nome da pessoa/empresa responsável pelo transporte da mercadoria." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Placa:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoPlaca" runat="server" />
                    <asp:TextBox ID="txtPlaca" runat="server" Enabled="False" Width="100px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultarPlaca" runat="server" Enabled="False" OnClick="btnConsultarPlaca_Click"
                        Text=" > " UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Identificação do veículo de transporte." />
                </div>
                <div class="collbl">
                    Classificação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlClassificacao" runat="server" Width="193px" Enabled="False"
                        AutoPostBack="True" OnSelectedIndexChanged="DdlClassificacao_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtProduto" runat="server" Width="611px" Enabled="False" />
                </div>
            </div>
            <div class="subtitulodiv">
                Pesagem
            </div>

            <div class="painelleft">
                <div class="row">
                    <div class="collbl">
                        1a_Pesagem:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPrimeiraPesagem" runat="server" TabIndex="160" Width="80px" Enabled="False" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        2a_Pesagem:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtSegundaPesagem" runat="server" TabIndex="170" Width="80px" Enabled="False" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Peso Bruto:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPesoBruto" runat="server" Width="80px" Enabled="False" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Desconto:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDesconto" runat="server" Width="80px" Enabled="False" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Líquido:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtLiquido" runat="server" Width="80px" Enabled="False" />
                    </div>
                </div>
            </div>
            <div class="painelleft">
                <div class="subtitulodiv">
                    Descontos
                </div>
                <div class="bordagrid" style="width: 520px; min-height: 123px; height: auto;">

                    <asp:GridView ID="gridDescontos" runat="server" CellPadding="4" ForeColor="#333333"
                        GridLines="None" Width="98%" AutoGenerateColumns="False">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Font-Names="Tahoma" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Font-Names="Tahoma" />
                        <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" Font-Names="Tahoma" />
                        <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" Font-Names="Tahoma" />
                        <EditRowStyle BackColor="#999999" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:BoundField DataField="CodigoAnalise" />
                            <asp:BoundField DataField="Descricao" HeaderText="Descrição" />
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
                </div>
                <div class="row">
                    <div class="coltxt" style="float: right;">
                        <asp:Button ID="cmdCalcular" runat="server" CssClass="botao" OnClick="cmdCalcular_Click"
                            TabIndex="180" Text="Calcular" Enabled="False" UseSubmitBehavior="False" />
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaCodMunicipios ID="ucConsultaCodMunicipios" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:ConsultaPlacas ID="ucConsultaPlacas" runat="server" />
    <uc:ConsultaAutorizacaoDeRetirada ID="ucConsultaAutorizacaoDeRetirada" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
</asp:Content>
