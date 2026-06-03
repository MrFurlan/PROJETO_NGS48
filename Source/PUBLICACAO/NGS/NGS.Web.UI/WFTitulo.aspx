<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="WFTitulo.aspx.vb" Inherits="NGS.Web.UI.WFTitulo" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1180px !important;
        }

        .collbl {
            width: 125px;
        }

        .w100 {
            width: 100px;
        }

        .w120 {
            width: 120px;
        }

        .lbls {
            min-height: 18px;
            display: block;
            white-space: normal;
            line-height: 16px;
            padding-left: 7px;
            text-indent: 0;
            text-transform: capitalize;
            white-space: normal;
            padding: 4px 0px 4px 7px;
        }

        .txtDecimal {
        }
    </style>
    <script type="text/javascript">
        function pageLoadTitulo() {
            $("#txtProrrogacao:enabled").datepicker({
                dateFormat: 'dd/mm/yy',
                dayNames: ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'],
                dayNamesMin: ['D', 'S', 'T', 'Q', 'Q', 'S', 'S', 'D'],
                dayNamesShort: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'],
                monthNames: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
                monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                nextText: 'Próximo',
                prevText: 'Anterior',
                showOn: "button",
                buttonImage: "Images/calendar.png",
                buttonImageOnly: true,
                showButtonPanel: true
            });

            $("#txtMovimento:enabled").datepicker({
                dateFormat: 'dd/mm/yy',
                dayNames: ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'],
                dayNamesMin: ['D', 'S', 'T', 'Q', 'Q', 'S', 'S', 'D'],
                dayNamesShort: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'],
                monthNames: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
                monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                nextText: 'Próximo',
                prevText: 'Anterior',
                showOn: "button",
                buttonImage: "Images/calendar.png",
                buttonImageOnly: true,
                showButtonPanel: true
            });

            $("#txtProrrogacao").setMask("date");

            $("#txtMovimento").setMask("date");

            $("#txtCodigoDeBarras").live('keypress', function (e) {
                var code = (e.keyCode ? e.keyCode : e.which);
                if (code == 13) {
                    $("#MainContent_BtValidarCodBarras").click();
                    return false;
                }
            });

            autoComplete();
        }

        $(document).ready(function () {
            pageLoadTitulo();
            var prmTitulo = Sys.WebForms.PageRequestManager.getInstance();
            prmTitulo.add_endRequest(pageLoadTitulo);
        });


        function validaSituacaoDoCliente() {
            var situacaoCliente = $("#<%=txtCodigoSituacaoCliente.ClientID%>").val();
            if (situacaoCliente != undefined && situacaoCliente != 0 && situacaoCliente != 1 && situacaoCliente != 50) {
                if (confirm("O Cliente está com a situação " + $("#<%=txtSituacaoCliente.ClientID%>").val() + "\nDeseja realmente gravar este título?")) {
                    return true;
                }
                return false;
            }
            return true;
        }

        function Arquivo() {
            $("#btnAdicionar").click();
        }

        function autoComplete() {

            var list = [];

            $('#MainContent_TabContainer1_TabPanel1_dltHistorico option').each(function (index) {
                var id = $(this).val();
                var name = $(this).text();
                list.push({ value: id, label: name });
            });

            $("#<%=txtHistoricoAutoComplete.ClientID%>").autocomplete({
                source: list,
                minLength: 3,
                select: function (event, ui) {
                    $("#<%=hdnHistoricoValue.ClientID%>").val(ui.item.value);
                    $("#<%=txtHistoricoAutoComplete.ClientID%>").val(ui.item.label);
                    $("#<%=txtHistorico.ClientID%>").val(ui.item.label);
                    return false;
                },
                change: function (event, ui) {
                    $("#<%=hdnHistoricoValue.ClientID%>").val(ui.item.value);
                    $("#<%=txtHistoricoAutoComplete.ClientID%>").val(ui.item.label);
                    $("#<%=txtHistorico.ClientID%>").val(ui.item.label);
                    return false;
                }
            });
        }
    </script>

</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngWFTitulo" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="50000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlWFTitulo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="HTabela" runat="server" Value="ContasAPagar" />
            <div class="titulodiv">
                <asp:Label ID="lblForm" runat="server" Text="Contas A Pagar" />
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%"
                Style="margin-top: 4px;">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Títulos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <asp:HiddenField ID="txtLiberarTitulo" runat="server" />
                        <asp:HiddenField ID="txtUsuarioLiberarTitulo" runat="server" />
                        <asp:HiddenField ID="txtUsuarioLiberarPedido" runat="server" />
                        <asp:HiddenField ID="txtUsuarioLiberarCheque" runat="server" />
                        <asp:HiddenField ID="txtUsuarioLiberarTituloData" runat="server" />
                        <asp:HiddenField ID="txtUsuarioLiberarPedidoData" runat="server" />
                        <asp:HiddenField ID="txtUsuarioLiberarChequeData" runat="server" />
                        <asp:HiddenField ID="txtTipoDoDocumento" runat="server" />
                        <asp:HiddenField ID="txtSlip" runat="server" />
                        <asp:HiddenField ID="txtSituacaoCliente" runat="server" />
                        <asp:HiddenField ID="txtCodigoSituacaoCliente" runat="server" />
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovo" runat="server" Text="Gravar" OnClientClick="return validaSituacaoDoCliente();" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconExcluir" ID="lnkExcluir" runat="server" Text="Excluir"
                                            OnClientClick="if(!confirm('Deseja realmente excluir este título?')) return false;" />
                                    </li>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultar" runat="server" Text="Consultar" />
                                    </li>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkSituacaoBancaria" runat="server" Text="Liberar Cobrança" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorio" runat="server" Text="Recibo" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorioAgrupamento" runat="server" Text="Relatório" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" />
                                    </li>
                                    <li runat="server" style="float: right;">
                                        <div class="row" style="margin-top: 0;">
                                            <div class="coltxt">
                                                <asp:Image ID="imgUsuarios" runat="server" Height="20px" ImageAlign="AbsMiddle" ImageUrl="~/Images/man2.png"
                                                    data-ToolTip="default" ToolTip="Usuário Lançamento" Width="18px" />
                                            </div>
                                            <div class="coltxt">
                                                <asp:DropDownList ID="ddlUsuarios" runat="server" Width="200px" />
                                            </div>
                                        </div>
                                    </li>
                                    <li runat="server" style="float: right;">
                                        <div class="row">
                                            <div class="coltxt" runat="server">
                                                <asp:CheckBox runat="server" ID="chkEmitirRecibo" Text="Emitir Recibo" />
                                                <asp:CheckBox runat="server" ID="chkManterLancamento" Text="Manter dados do lançamento" />
                                            </div>
                                        </div>
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            <asp:Label ID="lblCabCliFor" runat="server" Text="Dados do Fornecedor" />
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    Registro:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtRegistro" runat="server" Enabled="false" CssClass="txtNumerico9" Style="text-align: right;"
                                        Width="150px" Font-Size="10pt" Font-Bold="True" data-ToolTip="default" ToolTip="Clique em Consultar para buscar o Número do lançamento." />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgBloqueio" OnClick="imgBloqueio_Click" runat="server" CssClass="btn"
                                        ImageUrl="~/Images/liberar.png" ToolTip="Liberar Registro"
                                        ImageAlign="AbsMiddle" Visible="False" />
                                </div>
                                <div class="coltxt" style="padding-top: 7px;">
                                    <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                                        <ContentTemplate>
                                            <asp:ImageButton ID="imgBoletoPDF" Visible="False" runat="server" ImageUrl="~/images/icopdf16X16.jpg"
                                                ImageAlign="AbsMiddle" OnCommand="btnBoleto_Click" CommandName='<%# Eval("Codigo") %>' data-ToolTip="default"
                                                ToolTip="Boleto Bancário" />
                                        </ContentTemplate>
                                        <Triggers>
                                            <asp:PostBackTrigger ControlID="imgBoletoPDF" />
                                        </Triggers>
                                    </asp:UpdatePanel>
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="txtMestre" runat="server" CssClass="primario" Style="color: Red" />
                                </div>
                                <div class="coltxt" runat="server">
                                    <asp:CheckBox ID="ChkLiberado" runat="server" Text="Liberado" Enabled="False" />
                                    <asp:HiddenField ID="txtDataConciliacao" runat="server" />
                                    <asp:CheckBox ID="chkConciliado" runat="server" AutoPostBack="True" Text="Conciliado" Enabled="False" />
                                </div>

                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Unidade:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlUnidadeDeNegocioEmpresaCliente" runat="server" Width="582px"
                                        OnSelectedIndexChanged="DdlUnidadeDeNegocioEmpresaCliente_SelectedIndexChanged"
                                        AutoPostBack="True" />
                                </div>
                                <div class="coltxt">
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Empresa:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlEmpresaCliente" runat="server" Width="582px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    <asp:Label ID="lblCliente" runat="server" Text="Fornecedor:" />
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="txtCodigoFornecedor" runat="server" />
                                    <asp:TextBox ID="txtFornecedor" runat="server" Width="543px" Enabled="False" />
                                    <asp:Button ID="cmdClientesTitulo" CssClass="btn" OnClick="cmdClientesTitulo_Click"
                                        runat="server" UseSubmitBehavior="False" Text=">" data-ToolTip="default" ToolTip="Selecionar o fornecedor desejado." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Pedido:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPedido" runat="server" Width="120px" Enabled="False" Style="text-align: right;" AutoPostBack="True" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="cmdPedido" OnClick="cmdPedido_Click" runat="server" Text=">" UseSubmitBehavior="False"
                                        CssClass="btn" />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgExtratoPedido" OnClick="imgExtratoPedido_Click" runat="server"
                                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                        ToolTip="Visualizar Extrato do Pedido" ImageAlign="Middle" />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgLimparPedido" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/borracha.JPG"
                                        OnClick="imgLimparPedido_Click" ToolTip="Remover vínculo com Pedido"
                                        CssClass="btn" />
                                </div>
                                <div class="coltxt" style="width: 110px;">
                                    &nbsp;
                                </div>
                                <div runat="server">
                                    <div class="collbl">
                                        Procuração:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtCessaoDeCredito" runat="server" Width="120px" Style="text-align: right" />
                                    </div>
                                    <div runat="server" visible="False">
                                        <div class="collbl">
                                            Solicitação:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtSolicitacao" runat="server" Width="20px" data-ToolTip="default"
                                                ToolTip="Informar o número do contrato financeiro." />
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Favorecido:
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="txtCodigoFavorecido" runat="server" />
                                    <asp:TextBox ID="txtFavorecido" runat="server" Width="540px" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="cmdFavorecido" OnClick="cmdFavorecido_Click" runat="server" Text=">"
                                        CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Informar o favorecido." />
                                </div>
                            </div>
                            <div id="divBanco" runat="server" class="row">
                                <div class="collbl">
                                    Banco:
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnDadosBancarios" runat="server" OnClick="btnDadosBancarios_Click"
                                        Text=" &gt; " UseSubmitBehavior="False" CssClass="btn" />
                                    &nbsp;
                                <asp:Label ID="lblBanco" runat="server" Text="Banco" Font-Names="Calibri" Font-Size="9pt" />
                                    &nbsp;
                                |
                                &nbsp;
                                <asp:Label ID="lblAgencia" runat="server" Text="Agência" Font-Names="Calibri" Font-Size="9pt" />
                                    &nbsp;
                                 |
                                &nbsp;
                                <asp:Label ID="lblContaCorrente" runat="server" Text="Conta" Font-Names="Calibri" Font-Size="9pt" />
                                    &nbsp;
                                 |
                                &nbsp;
                                <asp:Label ID="lblTipoConta" runat="server" Text="Tipo da Conta" Font-Names="Calibri" Font-Size="9pt" />
                                    &nbsp;
                                 |
                                &nbsp;
                                <asp:ImageButton ID="imgContaObs" runat="server" ImageUrl="~/Images/ico_OBS_ativo.gif" />

                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgLimparContaBancaria" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/borracha.JPG"
                                        ToolTip="Remover Conta Bancária."
                                        CssClass="btn" />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 35%;">
                            <div class="row">
                                <div class="collbl w120">
                                    Data Entrada Sistema:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataEntradaSistema" runat="server" CssClass="calendario" Width="75px" Font-Size="9pt" Font-Bold="False" ClientIDMode="Static" Enabled="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w120">
                                    Venc. Original:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="lblVencOriginal" runat="server" Width="80px" Text="Label" Font-Bold="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w120">
                                    Venc. Atual:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtProrrogacao" Width="75px" runat="server" AutoPostBack="True" ClientIDMode="Static"
                                        OnTextChanged="txtProrrogacao_TextChanged" />
                                    <asp:HiddenField ID="hdnProrrogacaoOriginal" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w120" style="color: red;">
                                    Previsão Baixa:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlProvisoes" runat="server" Width="127px" OnSelectedIndexChanged="DdlProvisoes_SelectedIndexChanged"
                                        AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div runat="server">
                                    <div class="collbl  w120">
                                        Data Da Baixa:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtMovimento" Width="75px" runat="server" AutoPostBack="True" ClientIDMode="Static" 
                                            OnTextChanged="txtMovimento_TextChanged" data-ToolTip="default" ToolTip="Data do movimento para contabilização." />
                                        <asp:HiddenField ID="hdnMovimentoOriginal" runat="server" />
                                    </div>
                                    <div class="painelleft" runat="server" style="width: 100%; font-weight: bold;">
                                        <div class="row" runat="server" id="divlCot" style="border: 1px solid; text-align: center;">
                                            <asp:Label ID="lblDescCotacao" runat="server" />
                                            <br />
                                            <asp:Label ID="lblCotacao" runat="server" />
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="row" id="divArquivos" runat="server" visible="False">
                                <div class="coltxt">
                                    <div class="collbl" style="width: 103px">
                                        Arquivo:
                                    </div>
                                    <div class="coltxt">
                                        <uc:File ID="ucFile" runat="server" />
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            Dados da Empresa
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    <asp:Label ID="lblEmpresa" runat="server" Text="Empresa Pagadora:" />
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlEmpresaPagadora" runat="server" Width="550px" OnSelectedIndexChanged="DdlEmpresaPagadora_SelectedIndexChanged"
                                        AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    <asp:Label ID="lblTipoPgtoRec" runat="server" Text="Tipo Pgto:" />
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlTiposDePagamentos" runat="server" Width="550px" AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row" id="divCheque" runat="server">
                                <div class="collbl">
                                    Cheque:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNumeroCheque" runat="server" Enabled="False" Style="text-align: right"
                                        data-ToolTip="default" ToolTip="Informar o número da folha de cheque." />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgLimparCheque" CssClass="btn" OnClick="imgLimparCheque_Click"
                                        runat="server" ImageUrl="~/Images/borracha.JPG" ImageAlign="AbsMiddle"
                                        data-ToolTip="default" ToolTip="Remover vínculo com cheque" />
                                    <asp:HiddenField ID="txtEmiteCheque" runat="server" />
                                    <asp:HiddenField ID="txtLiberarCheque" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Banco:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlBancoPagador" runat="server" Width="550px" OnSelectedIndexChanged="DdlBancoPagador_SelectedIndexChanged"
                                        AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Conta:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlContaPagadora" runat="server" Width="550px" AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Moeda/Indexador:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlMoeda" runat="server" Width="127px" AutoPostBack="True" />
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlIndexador" runat="server" Width="170px" AutoPostBack="True" />
                                </div>
                            </div>

                            <div class="row">
                                <div class="collbl">
                                    Carteira do Título:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlCarteiraDoTitulo" runat="server" Width="550px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Finalidade Financeira:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlCarteiras" runat="server" Width="550px" OnSelectedIndexChanged="ddlCarteiras_SelectedIndexChanged"
                                        AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row" runat="server">
                                <div class="collbl">
                                    Encargo:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlTributos" runat="server" Width="550px" />
                                </div>
                            </div>
                            <div class="row" id="divNaviosXInvoice" runat="server" visible="False">
                                <div class="collbl">
                                    Invoice:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNaviosXInvoice" runat="server" Width="209px" Enabled="False" ToolTip="Código Invoice e descrição do navio." />
                                    <asp:LinkButton class="iconConsultar" ID="lnkConsultarNaviosXInvoice" runat="server" />
                                </div>
                            </div>
                            <div runat="server" id="divAd">
                                <asp:HiddenField ID="txtLiberarPedido" runat="server" />
                                <asp:HiddenField ID="txtPedidoFixacao" runat="server" />

                                <div class="row">
                                    <div class="collbl">
                                        Safra:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlSafraAdto" runat="server" Width="550px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Adiantamento:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtNumeroAdto" runat="server" Width="120px" Enabled="False" Style="text-align: right" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:Button ID="cmdAdiantamento" runat="server" Text=">"
                                            CssClass="btn" UseSubmitBehavior="False" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:ImageButton ID="imgLimparAdto" runat="server" ImageAlign="AbsMiddle" CssClass="btn"
                                            ImageUrl="~/Images/borracha.JPG" OnClick="imgLimparAdto_Click" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Venc. Adto:
                                    </div>
                                    <div class="coltxt" style="width: 198px;">
                                        <asp:TextBox ID="txtVencimentoAdto" runat="server" Width="120px" CssClass="calendario" data-ToolTip="default" ToolTip="Inserir a data de vencimento." />
                                    </div>
                                    <div class="collbl">
                                        Taxa de Juro:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtTaxaAdto" CssClass="txtDecimal" Style="text-align: right" runat="server" Width="120px" data-ToolTip="default" ToolTip="Informar a taxa de juro." />
                                    </div>
                                </div>
                            </div>
                            <div class="row" runat="server" id="divEndosso">
                                <div class="bordagrid" style="line-height: 10px; height: 50px;">
                                    <asp:GridView ID="gridEndosso" runat="server" AutoGenerateColumns="False"
                                        CellPadding="4" ForeColor="#333333" Width="99%" GridLines="None">
                                        <EditRowStyle BackColor="#999999" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                        <Columns>
                                            <asp:BoundField DataField="Codigo" HeaderText="Código">
                                                <HeaderStyle HorizontalAlign="Left" Width="50px" />
                                                <ItemStyle HorizontalAlign="Left" Width="50px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ClienteEndossoDescricao" HeaderText="Cliente do Endosso">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="NumeroEndosso" HeaderText="Número">
                                                <HeaderStyle HorizontalAlign="Left" Width="50px" />
                                                <ItemStyle HorizontalAlign="Left" Width="50px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento" HtmlEncode="False">
                                                <HeaderStyle HorizontalAlign="Left" Width="80px" />
                                                <ItemStyle HorizontalAlign="Left" Width="80px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Valor" HtmlEncode="False">
                                                <HeaderStyle HorizontalAlign="Right" Width="100px" />
                                                <ItemStyle HorizontalAlign="Right" Width="100px" />
                                            </asp:BoundField>
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </div>
                            <div class="row">
                                <div id="ContratoBanco" runat="server">
                                    <div class="collbl">
                                        Contrato Banco:
                                    </div>
                                    <div class="coltxt" style="width: 198px;">
                                        <asp:TextBox ID="txtContratoBanco" runat="server" Width="90px" Font-Size="9pt" Font-Bold="False"
                                            CssClass="texto" ClientIDMode="Static" MaxLength="20" data-ToolTip="default"
                                            ToolTip="Informar o número do contrato bancário." />
                                    </div>
                                </div>
                                <div class="collbl">
                                    Contrato Financeiro:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContratoFinanceiro" runat="server" MaxLength="30" Width="120px"
                                        data-ToolTip="default" ToolTip="Informar o número do contrato financeiro." />
                                </div>
                                <div id="DivFaturaDeFrete" runat="server" style="display: none;">
                                    <div class="collbl">
                                        Fatura de Frete:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtCodigoFaturaDeFrete" runat="server" CssClass="txtNumerico" Enabled="False" />
                                    </div>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Código de Barras:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCodigoDeBarras" runat="server" Width="295px" data-ToolTip="default"
                                        ToolTip="Informar o código de barras quando houver boleto bancário." />
                                    <asp:Button ID="BtValidarCodBarras" OnClick="BtValidarCodBarras_Click" runat="server"
                                        CssClass="btn" Text="Validar" data-ToolTip="default" ToolTip="Validar o código de barras." />
                                    <asp:CheckBox ID="CkbCodigoDeBarras" runat="server" Text="Digitado" />
                                    <asp:CheckBox ID="ckPreImpresso" runat="server" Text="Pré Impresso" />
                                </div>
                            </div>

                        </div>
                        <div class="painelleft" style="width: 450px;">
                            <div class="row">

                                <div class="coltxt right" style="margin-right: 130px;">
                                    <div class="subtitulodiv" style="width: 100px; margin-right: 5px;">
                                        DÓLAR<asp:Label ID="txtMoeda" runat="server" />
                                    </div>
                                </div>
                                <div class="coltxt right">
                                    <div class="subtitulodiv" style="width: 100px;">
                                        REAL<asp:Label ID="txtOficial" runat="server" />
                                    </div>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w100">
                                    Valor:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtValorDoDocumento" Width="90px" runat="server" CssClass="txtDecimal"
                                        Style="text-align: right" data-ToolTip="default" ToolTip="Inserir o valor total do pagamento." />
                                    <asp:HiddenField ID="HDValorOriginalOficial" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="HDValorOriginalMoeda" runat="server" />
                                    <asp:TextBox ID="txtValorEmMoeda" Width="90px" CssClass="txtDecimal" runat="server"
                                        Style="text-align: right" data-ToolTip="default" ToolTip="Inserir o valor total do pagamento." />
                                </div>
                            </div>
                            <div class="row" runat="server">
                                <div class="collbl w100" style="height: auto; min-height: 26px; margin-bottom: 4px;">
                                    <asp:Label ID="lblDescontos" CssClass="lbls" runat="server" Text="Descontos:"></asp:Label>
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDescontos" Width="90px" CssClass="txtDecimal" runat="server"
                                        Style="text-align: right" data-ToolTip="default" ToolTip="Informar quando tiver um desconto." />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDescontosMoeda" Width="90px" CssClass="txtDecimal" runat="server"
                                        Style="text-align: right" data-ToolTip="default" ToolTip="Informar quando tiver um desconto." />
                                </div>
                                <div class="coltxt">
                                    <label style="color: Red; font-weight: bold;">
                                        (-)</label>
                                </div>
                            </div>
                            <div class="row" runat="server">
                                <div class="collbl w100" style="height: auto; min-height: 26px; margin-bottom: 4px;">
                                    <asp:Label ID="lblDeducoes" CssClass="lbls" runat="server" Text="Deduções:"></asp:Label>
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDeducoes" Width="90px" CssClass="txtDecimal" runat="server" Style="text-align: right"
                                        data-ToolTip="default" ToolTip="Preencher com as deduções quando houver." />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDeducoesMoeda" Width="90px" CssClass="txtDecimal" runat="server"
                                        Style="text-align: right" data-ToolTip="default" ToolTip="Preencher com as deduções quando houver." />
                                </div>
                                <div class="coltxt">
                                    <label style="color: Red; font-weight: bold;">
                                        (-)</label>
                                </div>
                            </div>
                            <div class="row" runat="server">
                                <div class="collbl w100" style="height: auto; min-height: 26px; margin-bottom: 4px;">
                                    <asp:Label ID="lblJuros" CssClass="lbls" runat="server" Text="Multa/Juros:"></asp:Label>
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtJuros" Width="90px" runat="server" CssClass="txtDecimal" Style="text-align: right"
                                        data-ToolTip="default" ToolTip="Preencher com o valor de juros, quando houver." />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtJurosMoeda" Width="90px" CssClass="txtDecimal" runat="server"
                                        Style="text-align: right" data-ToolTip="default" ToolTip="Preencher com o valor de juros, quando houver." />
                                </div>
                                <div class="coltxt">
                                    <label style="color: Blue">
                                        (+)</label>
                                </div>
                            </div>
                            <div class="row" runat="server">
                                <div class="collbl w100" style="height: auto; min-height: 26px; margin-bottom: 4px;">
                                    <asp:Label ID="lblAcrescimos" CssClass="lbls" runat="server" Text="Acréscimos:"></asp:Label>
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtAcrescimos" Width="90px" runat="server" CssClass="txtDecimal"
                                        Style="text-align: right" data-ToolTip="default" ToolTip="Preencher com os acréssimos quando houver." />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtAcrescimosMoeda" Width="90px" CssClass="txtDecimal" runat="server"
                                        Style="text-align: right" data-ToolTip="default" ToolTip="Preencher com os acréssimos quando houver." />
                                </div>
                                <div class="coltxt">
                                    <label style="color: Blue">
                                        (+)</label>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w100">
                                    Valor Líquido:
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="txtValorLiquido" runat="server" />
                                    <asp:TextBox ID="txtValorCobrado" Width="90px" CssClass="txtDecimal" runat="server"
                                        ReadOnly="True" Style="text-align: right" data-ToolTip="default" ToolTip="Valor líquido a pagar após os descontos/acréscimos." />
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="txtValorLiquidoMoeda" runat="server" />
                                    <asp:TextBox ID="txtValorCobradoMoeda" Width="90px" CssClass="txtDecimal" runat="server"
                                        ReadOnly="True" Style="text-align: right" data-ToolTip="default" ToolTip="Valor líquido a pagar após os descontos/acréscimos." />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="ImgValores" runat="server" CssClass="btn"
                                        ImageUrl="~/Images/liberar.png" ToolTip="Liberar Valores"
                                        ImageAlign="AbsMiddle" />

                                </div>
                            </div>
                        </div>
                        <%-- <div class="row">
                            <div class="collbl">
                                Histórico:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSelecionarHist" runat="server" AutoPostBack="True" Style="width: 550px;"
                                    MaxLength="198" />
                            </div>
                            
                        </div>--%>
                        <%-- <div class="row">
                            <div class="collbl">
                                Histórico 2:
                            </div>
                            <div class="coltxt">
                                <input id="txtHist" list="MainContent_TabContainer1_TabPanel1_dlCountries" style="width: 540px;"
                                    runat="server" />
                              
                            </div>
                            <div class="collbl w100" style="margin-left: 3px;">
                                Observação:
                            </div>
                        </div>--%>
                            <div class="row" runat="server" id="divSelecaoAdBX" style="line-height: 10px;">
                                <asp:GridView ID="gridAdiantamentosDisponiveis" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Font-Size="10px">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:TemplateField HeaderText="Pedido">
                                            <ItemTemplate>
                                                <%# IIf(Eval("RegistroPedido") = 0, "Avulso", Eval("RegistroPedido"))%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="CodigoTitulo" HeaderText="Título" />
                                        <asp:BoundField DataField="Codigo" HeaderText="Adto" />
                                        <asp:BoundField DataField="Moeda.Simbolo" HeaderText="Moeda" />

                                        <asp:TemplateField HeaderText="C.Contábil">
                                            <ItemTemplate>
                                                <%--<%# Eval("DescContaAdto")%>--%>
                                                <%# Eval("CodigoContaAdto")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento" />
                                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento" />
                                        <asp:BoundField DataField="VlrAdto" DataFormatString="{0:N2}" HeaderText="Vlr Adto" />
                                        <asp:BoundField DataField="VlrBaixa" DataFormatString="{0:N2}" HeaderText="Vlr Baixado" />
                                        <asp:BoundField DataField="VlrSaldo" DataFormatString="{0:N2}" HeaderText="Saldo" />
                                        <asp:BoundField DataField="Taxa" DataFormatString="{0:N2}" HeaderText="Taxa" />
                                        <asp:TemplateField HeaderText="Baixa Oficial">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtVlrBaixaOficial" runat="server" DataFormatString="{0:N2}" CssClass="txtDecimal" Text='<%# Eval("VlrABaixarOficial", "{0:N2}")%>' Width="65px" OnTextChanged="txtVlrBaixaOficial_TextChanged" AutoPostBack="True" />
                                            </ItemTemplate>
                                        </asp:TemplateField>

                                        <asp:TemplateField HeaderText="Baixa Moeda">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtVlrBaixaMoeda" runat="server" DataFormatString="{0:N2}" CssClass="txtDecimal" Text='<%# Eval("VlrABaixarMoeda", "{0:N2}")%>' Width="65px" OnTextChanged="txtVlrBaixaMoeda_TextChanged" AutoPostBack="True" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="VlrRealBaixaOficial" DataFormatString="{0:N2}" HeaderText="Equiv.Adto" />
                                        <asp:BoundField DataField="VlrCalcVariacao" DataFormatString="{0:N2}" HeaderText="Variação" />
                                    </Columns>
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                </asp:GridView>
                            </div>
                            <div class="row" runat="server" id="divAdBaixados">
                                <asp:GridView ID="gridBaixasAdiantamentos" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:BoundField DataField="Adiantamento.CodigoTitulo" DataFormatString="{0:N0}" HeaderText="Titulo Adto" />
                                        <asp:BoundField DataField="Adiantamento.Codigo" DataFormatString="{0:N0}" HeaderText="Adiant." />
                                        <asp:BoundField DataField="CodigoTitulo" DataFormatString="{0:N0}" HeaderText="Titulo Bx." />
                                        <asp:BoundField DataField="Sequencia" DataFormatString="{0:N0}" HeaderText="Seq." />
                                        <asp:BoundField DataField="DataBaixa" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento" />
                                        <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor Oficial" />
                                        <asp:BoundField DataField="VariacaoOficial" DataFormatString="{0:N2}" HeaderText="Valor Variacao" />
                                        <asp:BoundField DataField="ValorMoeda" DataFormatString="{0:N2}" HeaderText="Valor Moeda" />
                                    </Columns>
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                </asp:GridView>
                            </div>
                        <div class="row">
                            <div class="collbl">
                                Histórico:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtHistoricoAutoComplete" Style="width: 540px;"
                                    runat="server" />
                                <asp:HiddenField ID="hdnHistoricoValue" runat="server" />
                                <datalist id="dltHistorico" runat="server">
                                </datalist>
                            </div>
                            <div class="collbl w100" style="margin-left: 3px;">
                                Observação:
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:TextBox ID="txtHistorico" runat="server" Width="540px" MaxLength="200" Style="margin-left: 137px;"
                                    TextMode="MultiLine" data-ToolTip="default" ToolTip="Selecionar o histórico adequado da operação." />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacoes" runat="server" Width="420px" Height="45px" TextMode="MultiLine"
                                    data-ToolTip="default" ToolTip="Preencher quando houver alguma observação relevante." />
                            </div>
                        </div>
                        <div class="row" id="idRetornoBancario" visible="False" runat="server">
                            <div class="collbl">
                                Retorno Bancário:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObsRetornoBancario" runat="server" Width="540px" Height="45px" MaxLength="200" TextMode="MultiLine" data-ToolTip="default" ToolTip="Selecionar o histórico adequado da operação." />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel3" runat="server">
                    <HeaderTemplate>
                        Pedidos/Parcelas
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="GridPedidos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Registro" HeaderText="Registro">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Baixa" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Baixa"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Historico" HeaderText="Hist&#243;rico">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorDoDocumento" DataFormatString="{0:N}" HeaderText="Valor da Parcela"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descontos" DataFormatString="{0:N}" HeaderText="Descontos"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Deducoes" DataFormatString="{0:N}" HeaderText="Dedu&#231;&#245;es"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Juros" DataFormatString="{0:N}" HeaderText="Juros" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Acrescimos" DataFormatString="{0:N}" HeaderText="Acr&#233;scimos"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorLiquido" DataFormatString="{0:N}" HeaderText="Valor da Baixa"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Saldo" DataFormatString="{0:N}" HeaderText="Saldo" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Documento" ID="TabDocumento">
                    <HeaderTemplate>
                        Documento(s)
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row" runat="server">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricaoDocumento" runat="server" MaxLength="100" Width="500px" data-ToolTip="default"
                                    ToolTip="Descrição do Documento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Arquivo:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeDoArquivo" runat="server" Enabled="false" Style="width: 200px;" />
                                <asp:FileUpload ID="fupArquivo" OnChange="Arquivo()" runat="server" Width="120px" Font-Size="11px" ClientIDMode="Static" />
                            </div>
                        </div>
                        <div class="painelleft bordagrid" style="width: 99%; margin-left: 0.5%; height: 225px;">
                            <asp:GridView ID="gridDocumentos" runat="server" AutoGenerateColumns="False" ForeColor="#333333"
                                GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Codigo" HeaderText="Código">
                                        <HeaderStyle HorizontalAlign="Center" Width="70px" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                                        <HeaderStyle HorizontalAlign="Left" Width="500px" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeDoArquivo" HeaderText="Nome do arquivo">
                                        <HeaderStyle HorizontalAlign="Left" Width="300px" />
                                        <ItemStyle HorizontalAlign="Left" Width="300px" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Arquivo">
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgDownload" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/download32x32.png" Style="margin-top: 0;"
                                                Height="18px" Width="18px" OnClick="imgDownload_Click" ToolTip="Baixar Arquivo" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Excluir">
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirDocumento" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                OnClick="imgExcluirDocumento_Click" data-ToolTip="default" ToolTip="Excluir"
                                                OnClientClick="return confirm('Deseja realmente excluir o Documento?');" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <asp:UpdatePanel ID="updpnlAdicionar" runat="server">
                            <ContentTemplate>
                                <asp:Button ID="btnAdicionar" runat="server" Text="Adicionar" ClientIDMode="Static"
                                    OnClick="btnAdicionar_Click" CssClass="none" />
                            </ContentTemplate>
                            <Triggers>
                                <asp:PostBackTrigger ControlID="btnAdicionar" />
                            </Triggers>
                        </asp:UpdatePanel>
                        <asp:Label ID="lblMsg" runat="server" />
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel4" runat="server">
                    <HeaderTemplate>
                        Consulta Títulos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultarTitulo" Text="Consultar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                                        <ul>
                                            <li>
                                                <asp:LinkButton class="iconPdf" ID="lnkRelatorioCTitulo" runat="server" Text="Pdf" />
                                            </li>
                                            <li>
                                                <asp:LinkButton class="iconExcel" ID="lnkExcelCTitulo" runat="server" Text="Excel" />
                                            </li>
                                            <li>
                                                <asp:LinkButton class="iconExcel" ID="lnkExcelCTituloDados" runat="server" Text="Excel Dados" />
                                            </li>
                                        </ul>
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparConsultaTitulo" Text="Limpar" runat="server"
                                            data-tooltip="default" ToolTip="Limpa as informações de consulta." />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkAgruparPagamento" Text="Agrupar Pagamento" runat="server"
                                            data-tooltip="default" ToolTip="Agrupar Pagamento referente os Registros Selecionados." />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkBaixarSelecionados" Text="Baixar Selecionados" runat="server"
                                            data-tooltip="default" ToolTip="Baixar Todos os Registros Selecionados." />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkReprogramar" Text="Reprogramar" runat="server"
                                            data-tooltip="default" ToolTip="Habilita campo para informar o novo vencimento." />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRecibo" Text="Recibo" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkSlip" Text="Slip" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlUnidadeConsultaTitulos" runat="server" AutoPostBack="True"
                                    OnSelectedIndexChanged="DdlUnidadeConsultaTitulos_SelectedIndexChanged" Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaConsultaTitulos" runat="server" Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl" style="line-height: 14px; padding-left: 7px; padding-bottom: 2px; text-indent: 0; height: auto; width: 118px; margin-bottom: 4px;">
                                <asp:CheckBox ID="chkClientes" ClientIDMode="Static" runat="server" Text="Cons/Cliente:"
                                    onclick="if($(this).attr('checked') == 'checked') msgbox('Campo utilizado apenas para relatório.','ATENÇÃO!','Info');" /><br />
                                <asp:CheckBox ID="chkClienteEndosso" runat="server" AutoPostBack="true" OnCheckedChanged="chkClienteEndosso_CheckedChanged" Text="Endosso" data-ToolTip="default" ToolTip="Utilizado apenas para Consulta e Relatório." />
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoClienteConsulta" runat="server" />
                                <asp:TextBox ID="txtClienteConsulta" runat="server" Width="556px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdConsultaClientes" CssClass="btn" OnClick="cmdConsultaClientes_Click"
                                    runat="server" UseSubmitBehavior="False" Text=">" data-ToolTip="default" ToolTip="Unificar os clientes para consulta." />
                            </div>
                        </div>
                        <div class="row" runat="server" id="divClienteConsultaEndosso" visible="False">
                            <div class="collbl">
                                Cliente de Endosso:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoClienteConsultaEndosso" runat="server" />
                                <asp:TextBox ID="txtClienteConsultaEndosso" runat="server" Width="556px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdConsultaClientesEndosso" CssClass="btn" OnClick="cmdConsultaClientesEndosso_Click"
                                    runat="server" UseSubmitBehavior="False" Text=">" data-ToolTip="default" ToolTip="Clientes de Endosso para consulta." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedidoConsultaTitulos" runat="server" Width="100px" CssClass="txtNumerico" data-ToolTip="default" ToolTip="Número do pedido." />
                            </div>
                            <div class="coltxt">
                                <asp:Button CssClass="btn" ID="btnBuscaPedidoConsultaTitulos" OnClick="btnBuscaPedidoConsultaTitulos_Click"
                                    runat="server" Text=">" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Número do pedido." />
                            </div>
                            <div class="collbl" style="width: 100px;">
                                Vencto Inicial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" runat="server" Width="80px" CssClass="calendario" data-ToolTip="default" ToolTip="Data inicial do vencimento do pagamento." />
                            </div>
                            <div class="collbl" style="width: 100px;">
                                Vencto Final:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" runat="server" Width="80px" CssClass="calendario" data-ToolTip="default" ToolTip="Data final do vencimento do pagamento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Nota:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNumNota" runat="server" Width="337px" Style="text-align: right" data-ToolTip="default" ToolTip="Número da nota fiscal." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Parametros:
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="ChkAutozizado" runat="server" Text="Autorizados" data-ToolTip="default" ToolTip="Selecionar a opção de acordo com a consulta desejada." />
                                <asp:CheckBox ID="chkObservacao" runat="server" Text="Emitir Historico com Observações" data-ToolTip="default" ToolTip="Selecionar a opção de acordo com a consulta desejada." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Situação:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RbAtivo" runat="server" AutoPostBack="True" Checked="True" Text="Ativos"
                                    OnCheckedChanged="RbAtivo_CheckedChanged" GroupName="Situacao" data-ToolTip="default" ToolTip="Informar se o título está ativo ou cancelado." />
                                <asp:RadioButton ID="RbCancelado" runat="server" Text="Cancelados" OnCheckedChanged="RbCancelado_CheckedChanged"
                                    GroupName="Situacao" AutoPostBack="True" data-ToolTip="default" ToolTip="Informar se o título está ativo ou cancelado." />
                                <asp:RadioButton ID="RbGeral" runat="server" AutoPostBack="True" Text="Todos" OnCheckedChanged="RbGeral_CheckedChanged"
                                    GroupName="Situacao" data-ToolTip="default" ToolTip="Informar se o título está ativo ou cancelado." />
                                &nbsp; &nbsp;
                                <asp:CheckBox ID="chkPrevisao" runat="server" Text="Previsão" Checked="True" data-ToolTip="default" ToolTip="Informar se o título está ativo ou cancelado." />
                                <asp:CheckBox ID="chkProvisao" runat="server" Text="Provisão" data-ToolTip="default" ToolTip="Informar se o título está ativo ou cancelado." />
                                <asp:CheckBox ID="chkBaixado" runat="server" Text="Baixado" data-ToolTip="default" ToolTip="Informar se o título está ativo ou cancelado." />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblAgrupar" runat="server" Width="100px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Totalização:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RbDiaGeral" runat="server" Checked="True" GroupName="Totalizacao"
                                    Text="Diario" Width="66px" data-ToolTip="default" ToolTip="Informar se o total será por diário, filial ou carteira." />
                                <asp:RadioButton ID="RbFilialDiario" runat="server" GroupName="Totalizacao"
                                    Text="Filial" data-ToolTip="default" ToolTip="Informar se o total será por diário, filial ou carteira." />
                                <asp:RadioButton ID="RbCarteiraDia" runat="server" GroupName="Totalizacao"
                                    Text="Carteira" data-ToolTip="default" ToolTip="Informar se o total será por diário, filial ou carteira." />
                            </div>
                        </div>
                        <div class="row" id="pnlReprogramaVencimentos" runat="server">
                            <div class="collbl">
                                <label class="primario">
                                    Novo Vencimento:</label>
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNovoVencimento" runat="server" CssClass="calendario" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnNovoVencimento" runat="server" CssClass="botao" OnClick="btnNovoVencimento_Click"
                                    Text="Confirmar" data-ToolTip="default" ToolTip="Confirmar Reprogramação dos Vencimentos"
                                    UseSubmitBehavior="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnCancelarNonoVencimento" runat="server" CssClass="botao" OnClick="btnCancelarNonoVencimento_Click"
                                    Text="Cancelar" UseSubmitBehavior="False" />
                            </div>
                        </div>
                        <div class="row" id="rowDolar" runat="server" visible="False">
                            <div class="coltxt">
                                <asp:Label ID="lblTotalRegistroAgrupado" runat="server" Font-Bold="True" />
                                <asp:HiddenField ID="txtRealDolar" runat="server" />
                                <asp:HiddenField ID="HiddenIndexador" runat="server" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="GridConsultaTitulos" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="GridConsultaTitulos_SelectedIndexChanged"
                                Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" HeaderText="Sel" SelectText="&gt;" ShowSelectButton="True">
                                        <ItemStyle Width="30px" />
                                    </asp:CommandField>
                                    <asp:TemplateField HeaderText="LB">
                                        <ItemTemplate>
                                            <asp:CheckBox ID="ChkLiberado" runat="server" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <HeaderTemplate>
                                            <asp:CheckBox ID="chkAllTitulos" data-ToolTip="default" ToolTip="Seleciona todos os títulos de mesma moeda e indexador."
                                                Text="CK" runat="server" AutoPostBack="True" OnCheckedChanged="chkAllTitulos_CheckedChanged" />
                                        </HeaderTemplate>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="ChkGridTitulos" runat="server" AutoPostBack="True" OnCheckedChanged="ChkGridTitulos_CheckedChanged" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Registro" HeaderText="Título">
                                        <ItemStyle Width="60px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                                        HtmlEncode="False">
                                        <ItemStyle Width="80px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                        <ItemStyle Width="200px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Historico" HeaderText="Histórico">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" Width="250px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Dolar" DataFormatString="{0:N}" HeaderText="Dolares">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" Width="100px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorLiquido" DataFormatString="{0:N}" HeaderText="Reais"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" Width="100px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Liberado" HeaderText="Autorizante" ShowHeader="False">
                                        <HeaderStyle Width="0px" />
                                        <ItemStyle Width="0px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Moeda">
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Indexador" HeaderText="Ind">
                                        <HeaderStyle HorizontalAlign="Right" Width="20px" />
                                        <ItemStyle HorizontalAlign="Right" Width="20px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Grupado" HeaderText="Agrupado">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" Width="30px" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <HeaderTemplate>
                                            <asp:CheckBox ID="chkReprogramarAll" data-ToolTip="default" ToolTip="Seleciona todos os títulos para efetuar uma reprogramação de vencimentos."
                                                Text="RP" runat="server" AutoPostBack="True" OnCheckedChanged="chkReprogramarAll_CheckedChanged" />
                                        </HeaderTemplate>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkReprogramar" runat="server" data-ToolTip="default" ToolTip="Reprogramar Vencimento" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderImageUrl="~/Images/impressora3.jpg">
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkRecibo" runat="server" data-ToolTip="default" ToolTip="Emitir Recibo" />
                                        </ItemTemplate>
                                        <HeaderStyle Height="23px" Width="23px" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Situacao" HeaderText="" />
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel5" runat="server">
                    <HeaderTemplate>
                        Contabilização
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="gridRazao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField HeaderText="Empresa" DataField="Empresa">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Conta" DataField="Conta">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Cliente" DataField="Cliente">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Titulo" HeaderText="Descri&#231;&#227;o">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Lote" DataField="Lote">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Produto" DataField="Produto">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Custo" DataField="Custo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Hist&#243;rico" DataField="Historico">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="D&#233;bito" DataField="Debito" DataFormatString="{0:n2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Cr&#233;dito" DataField="Credito" DataFormatString="{0:n2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Saldo" DataField="Saldo" DataFormatString="{0:n2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel6" runat="server">
                    <HeaderTemplate>
                        Processa Retorno Bancário
                    </HeaderTemplate>
                    <ContentTemplate>
                        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                            <ContentTemplate>
                                <div class="menu_acoes">
                                    <div class="acoes">
                                        <ul>
                                            <li runat="server">
                                                <asp:LinkButton class="iconConsultar" ID="lnkImportar" runat="server" Text="Importar"
                                                    OnClientClick="if (!confirm('Confirma importação do arquivo de retorno?')) return false;" />
                                            </li>
                                            <li runat="server">
                                                <asp:LinkButton class="iconRelatorio" ID="lnkBaixar" runat="server" Text="Baixar" />
                                            </li>
                                        </ul>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Arquivo:
                                    </div>
                                    <div class="coltxt">
                                        <asp:FileUpload ID="fup" runat="server" Width="575px" />
                                    </div>
                                </div>
                                <div class="bordagrid">
                                    <asp:GridView ID="GridRetornoTitulos" runat="server" AutoGenerateColumns="False"
                                        CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <EditRowStyle BackColor="#999999" />
                                        <Columns>
                                            <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"
                                                HeaderText="Sel">
                                                <ItemStyle Width="30px" />
                                            </asp:CommandField>
                                            <asp:TemplateField HeaderText="CK">
                                                <ItemTemplate>
                                                    <asp:CheckBox ID="ChkGridTitulos" runat="server" AutoPostBack="True" />
                                                </ItemTemplate>
                                                <HeaderStyle Width="30px" />
                                                <ItemStyle Width="30px" />
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="Registro" HeaderText="Titulo">
                                                <ItemStyle Width="60px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                                                HtmlEncode="False">
                                                <ItemStyle Width="80px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                                <ItemStyle Width="200px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Valor" HeaderText="Valor" DataFormatString="{0:n2}" />
                                            <asp:BoundField DataField="DataPgto" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Pgto" />
                                            <asp:BoundField DataField="CodRetorno" HeaderText="CodRetorno" Visible="False">
                                                <ItemStyle Width="250px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="DescRetorno" HeaderText="Retorno" />
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </ContentTemplate>
                            <Triggers>
                                <asp:PostBackTrigger ControlID="lnkImportar" />
                            </Triggers>
                        </asp:UpdatePanel>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaAdiantamentos ID="ucConsultaAdiantamentos" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:DestinoContabil ID="ucDestinoContabil" runat="server" />
    <uc:ConsultaDadosBancarios ID="ucConsultaDadosBancarios" runat="server" />
    <uc:BaixaLoteFinanceiro ID="ucBaixaLoteFinanceiro" runat="server" />
    <uc:ConsultarNaviosXInvoice ID="ucConsultarNaviosXInvoice" runat="server" />
    <uc:ConsultarTitulo ID="ucConsultaTitulo" runat="server" />
</asp:Content>
