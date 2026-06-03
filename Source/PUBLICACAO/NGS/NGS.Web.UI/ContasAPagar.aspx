<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ContasAPagar.aspx.vb" Inherits="NGS.Web.UI.ContasAPagar" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1080px !important;
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
    </style>
    <script type="text/javascript">
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
    </script>

</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngContasAPagar" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlContasAPagar" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                <label>
                    Contas A Pagar
                </label>
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
                        <asp:HiddenField ID="HDSaldoAdiantamento" runat="server" />
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
                                </ul>
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            <label>
                                Dados do Fornecedor
                            </label>
                        </div>

                        <div class="row">
                            <div class="collbl">
                                Registro:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistro" runat="server" CssClass="txtNumerico9" Style="text-align: right;"
                                    Width="150px" Font-Size="10pt" Font-Bold="True" data-ToolTip="default" ToolTip="Número do lançamento." />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgBloqueio" OnClick="imgBloqueio_Click" runat="server" CssClass="btn"
                                    ImageUrl="~/Images/liberar.png" ToolTip="Liberar Registro"
                                    ImageAlign="AbsMiddle" Visible="False" />
                            </div>

                            <div class="coltxt">
                                <asp:Label ID="txtMestre" runat="server" CssClass="primario" Style="color: Red" />
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox runat="server" ID="chkEmitirRecibo" Text="Emitir Recibo" />
                                <asp:CheckBox runat="server" ID="chkManterLancamento" Text="Manter dados do lançamento" />
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
                                <asp:CheckBox ID="ChkLiberado" runat="server" Text="Liberado" Enabled="False" />
                                <asp:HiddenField ID="txtDataConciliacao" runat="server" />
                                <asp:CheckBox ID="chkConciliado" runat="server" Text="Conciliado" Enabled="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaCliente" runat="server" Width="582px" />
                            </div>
                            <div class="coltxt">
                                Inclusão:
                                <asp:Image ID="imgUsuarioIncl" runat="server" Width="18px" Height="20px" ImageUrl="~/Images/man2.png"
                                    ImageAlign="AbsMiddle"></asp:Image>
                                <asp:Label ID="lblUsuarioIncl" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Fornecedor:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoFornecedor" runat="server" />
                                <asp:TextBox ID="txtFornecedor" runat="server" Width="543px" Enabled="False" />
                                <asp:Button ID="cmdClientesTitulo" CssClass="btn" OnClick="cmdClientesTitulo_Click"
                                    runat="server" UseSubmitBehavior="False" Text=">" data-ToolTip="default" ToolTip="Selecionar o fornecedor desejado." />
                            </div>
                            <div class="coltxt">
                                Alteração:
                                <asp:Image ID="imgUsuarioAlt" runat="server" Width="18px" Height="20px" ImageUrl="~/Images/man2.png"
                                    ImageAlign="AbsMiddle"></asp:Image>
                                <asp:Label ID="lblUsuarioAlt" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedido" runat="server" Width="120px" Enabled="False" Style="text-align: right;" />
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
                        <div class="subtitulodiv">
                            Dados da Empresa
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    Empresa Pagadora:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlEmpresaPagadora" runat="server" Width="550px" OnSelectedIndexChanged="DdlEmpresaPagadora_SelectedIndexChanged"
                                        AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Tipo Pgto:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlTiposDePagamentos" runat="server" Width="550px" />
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
                                    <asp:DropDownList ID="DdlContaPagadora" runat="server" Width="550px" />
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
                                    Previsão Baixa:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlProvisoes" runat="server" Width="127px" OnSelectedIndexChanged="DdlProvisoes_SelectedIndexChanged"
                                        AutoPostBack="True" />
                                </div>
                                <div runat="server">
                                    <div class="collbl w100">
                                        Data Da Baixa:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtMovimento" runat="server" Width="75px" Font-Size="9pt" Font-Bold="False"
                                            ClientIDMode="Static" data-ToolTip="default" ToolTip="Data do movimento para contabilização." />
                                        <asp:HiddenField ID="hdnMovimentoOriginal" runat="server" />
                                    </div>
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

                            <div class="row">
                                <asp:HiddenField ID="txtLiberarPedido" runat="server" />
                                <asp:HiddenField ID="txtPedidoFixacao" runat="server" />


                                <div class="row" runat="server">
                                    <div class="collbl">
                                        Carteira Adto:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlCarteirasAdto" runat="server" Width="550px" AutoPostBack="True"
                                            OnSelectedIndexChanged="DdlCarteirasAdto_SelectedIndexChanged" />
                                    </div>
                                </div>

                                <div class="row" runat="server">
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
                                <div class="row" runat="server">
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
                            <div class="row">
                                <div class="collbl">
                                    Histórico:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlSelecionarHist" runat="server" AutoPostBack="True" Style="width: 550px;"
                                        MaxLength="198" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt">
                                    <asp:TextBox ID="txtHistorico" runat="server" Width="540px" Style="margin-left: 137px;"
                                        TextMode="MultiLine" data-ToolTip="default" ToolTip="Selecionar o histórico adequado da operação." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Observação:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtObservacoes" runat="server" Width="540px" Height="40px" TextMode="MultiLine"
                                        data-ToolTip="default" ToolTip="Preencher quando houver alguma observação relevante." />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 350px;">
                            <div class="painelleft" style="width: 70%;">
                                <div class="row">
                                    <div class="collbl w120">
                                        Data Entrada Sistema:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtDataEntradaSistema" runat="server"  CssClass="calendario" Width="75px" Font-Size="9pt" Font-Bold="False" ClientIDMode="Static" Enabled="False" />
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
                                        <asp:TextBox ID="txtProrrogacao" Width="75px" runat="server" data-ToolTip="default"
                                            ToolTip="Data do vencimento do pagamento." AutoPostBack="True" ClientIDMode="Static"
                                            OnTextChanged="txtProrrogacao_TextChanged" />
                                        <asp:HiddenField ID="hdnProrrogacaoOriginal" runat="server" />
                                    </div>
                                </div>
                            </div>
                            <div class="painelleft" style="width: 33%; margin-left: 2px; text-align: center; font-weight: bold;">
                                <div class="row" style="line-height: 16px; text-align: center;">
                                    <asp:Label ID="lblDescCotacao" runat="server" />
                                    <br />
                                    <asp:Label ID="lblCotacao" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt right" style="margin-right: 30px;">
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
                                    <asp:Button ID="ImgCalcular" OnClick="ImgCalcular_Click" runat="server" CssClass="btn"
                                        UseSubmitBehavior="False" Text="=" data-ToolTip="default" ToolTip="Calcular Valores" />
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Dados Bancários
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkGravarDadosBancarios" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkAlterarDadosBancarios" Text="Atualizar"
                                            runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconExcluir" ID="lnkExcluirDadosBancarios" Text="Excluir"
                                            runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparDadosBancarios" Text="Limpar" runat="server" />
                                    </li>
                                </ul>
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
                        <div class="row">
                            <div class="collbl">
                                Banco:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlBancos" runat="server" Width="568px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Agência:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAgencia" runat="server" MaxLength="4" data-ToolTip="default"
                                    ToolTip="Informar em qual agência ocorrerá o pagamento." />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDigitoAgencia" runat="server" Width="30px" MaxLength="1" data-ToolTip="default"
                                    ToolTip="Informar em qual agência ocorrerá o pagamento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conta:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtContaCorrente" runat="server" MaxLength="12" data-ToolTip="default"
                                    ToolTip="Selecionar a conta que realizará o pagamento." />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDigitoDaConta" runat="server" Width="30px" MaxLength="2" data-ToolTip="default"
                                    ToolTip="Selecionar a conta que realizará o pagamento." />
                            </div>
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
                                Tipo Conta:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTipoConta" runat="server" Width="121px">
                                    <asp:ListItem Value="C">C. Corrente</asp:ListItem>
                                    <asp:ListItem Value="P">C. Poupan&#231;a</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observções:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacoesDaConta" runat="server" Width="568px" TextMode="MultiLine"
                                    MaxLength="300" data-ToolTip="default" ToolTip="Preencher quando houver alguma observação relevante." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="GridContasCorrentes" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridContasCorrentes_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                                    <asp:BoundField DataField="Banco_Id" HeaderText="Banco" ReadOnly="True">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Agencia_Id" HeaderText="Agencia">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DigitoAgencia_Id" HeaderText="Dg">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ContaCorrente_Id" HeaderText="Conta">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DigitoConta_Id" HeaderText="Dg">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="TipoConta" HeaderText="Tipo">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Observacoes" HeaderText="Observacoes">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
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
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatorioCTitulo" Text="Relatório" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparConsultaTitulo" Text="Limpar" runat="server"
                                            data-tooltip="default" ToolTip="Limpa as informações de consulta." />
                                    </li>
                                    <li runat="server" style="width: 164px;">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkAgruparPagamento" Text="Agrupar Pagamento"
                                            runat="server" />
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
                            <div class="collbl">
                                <asp:CheckBox ID="chkClientes" ClientIDMode="Static" runat="server" Text="Cons/Cliente:"
                                    onclick="if($(this).attr('checked') == 'checked') msgbox('Campo utilizado apenas para relatório.','ATENÇÃO!','Info');" />
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoClienteConsulta" runat="server" />
                                <asp:TextBox ID="txtClienteConsulta" runat="server" Width="556px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdConsultaClientes" CssClass="btn" OnClick="cmdConsultaClientes_Click"
                                    runat="server" UseSubmitBehavior="False" Text=">" data-ToolTip="default" ToolTip="Unificar os clientes pra consulta." />
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
                        <div class="row" id="rowDolar" runat="server" visible="false">
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
</asp:Content>
