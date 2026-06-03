<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ContasAReceber.aspx.vb" Inherits="NGS.Web.UI.ContasAReceber" %>

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
        

        function selectAll(chkAll) {
            
            var chk = $('#' + chkAll.id);
            var checked = chk.attr('checked') == "checked";
           
            $("input[type='checkbox']", <%=GridConsultaTitulos.ClientID%>).each(function () {
                if ($(this).parent().attr('class') =='chk tooltipstered') {
                    $(this).attr("checked", checked);
                    if (checked==false){
                        $("#HiddenIndexador").value="";
                    }
                }
            });
        };
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngContasAReceber" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlContasAReceber" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="txtUsuarioLiberarTitulo" runat="server" />
            <asp:HiddenField ID="txtUsuarioLiberarPedido" runat="server" />
            <asp:HiddenField ID="txtUsuarioLiberarTituloData" runat="server" />
            <asp:HiddenField ID="txtUsuarioLiberarPedidoData" runat="server" />
            <asp:HiddenField ID="HDSaldoAdiantamento" runat="server" />
            <div class="titulodiv">
                Contas A Receber
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%"
                Style="margin-top: 4px;">
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel1" ID="TabPanel1">
                    <HeaderTemplate>
                        Títulos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovo" runat="server" Text="Gravar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconExcluir" ID="lnkExcluir" runat="server" Text="Excluir"
                                            OnClientClick="return confirm('Tem certeza que deseja excluir este título?');" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" runat="server" Text="Recibo" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                                    </li>
                                    <br />
                                </ul>
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            DADOS DO CLIENTE
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Registro:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtLiberarTitulo" runat="server" />
                                <asp:TextBox ID="txtRegistro" TabIndex="2" runat="server" CssClass="txtNumerico9"
                                    Width="150px" Font-Size="10pt" Font-Bold="True" data-ToolTip="default" ToolTip="Número do lançamento." />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgBloqueio" OnClick="imgBloqueio_Click" runat="server" Width="25px"
                                    Height="25px" ImageUrl="~/images/liberar.png" ImageAlign="AbsMiddle" Visible="False"
                                    data-ToolTip="default" ToolTip="Liberar Registro" />
                            </div>

                            <div class="coltxt">
                                <asp:Label ID="txtMestre" runat="server" CssClass="primario" Style="color: Red" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlUnidadeDeNegocioEmpresaCliente" TabIndex="3" runat="server"
                                    Width="583px" OnSelectedIndexChanged="DdlUnidadeDeNegocioEmpresaCliente_SelectedIndexChanged"
                                    AutoPostBack="True" />
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="chkEmitirRecibo" runat="server" Text="Emitir Recibo" />
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox runat="server" ID="chkManterLancamento" Text="Manter dados do lançamento" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaCliente" TabIndex="4" runat="server" Width="583px" />
                            </div>
                            <div class="coltxt">
                                Inclusão:
                                <asp:Image ID="imgUsuarioIncl" runat="server" Width="18px" Height="20px" ImageUrl="~/images/man2.png"
                                    ImageAlign="AbsMiddle"></asp:Image>
                                <asp:Label ID="lblUsuarioIncl" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                                <asp:TextBox ID="txtCliente" runat="server" Enabled="False" Style="width: 543px;" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdClientesTitulo" OnClick="cmdClientesTitulo_Click" runat="server"
                                    CssClass="btn" Text=">" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                            </div>
                            <div class="coltxt">
                                Alteração:
                                <asp:Image ID="imgUsuarioAlt" runat="server" Height="20px" ImageAlign="AbsMiddle"
                                    ImageUrl="~/images/man2.png" Width="18px" />
                                <asp:Label ID="lblUsuarioAlt" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedido" runat="server" Enabled="False" Style="text-align: right" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdPedido" OnClick="cmdPedido_Click" runat="server" UseSubmitBehavior="False"
                                    CssClass="btn" Text=">" data-ToolTip="default" ToolTip="Número do pedido." />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgExtratoPedido" OnClick="imgExtratoPedido_Click" runat="server"
                                    ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ImageAlign="AbsMiddle"
                                    data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido" />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgLimparPedido" Width="22px" Height="22px" OnClick="imgLimparPedido_Click"
                                    runat="server" CssClass="btn" ImageUrl="~/images/borracha.JPG" ImageAlign="AbsMiddle"></asp:ImageButton>
                                <asp:HiddenField ID="txtLiberarPedido" runat="server" />
                                <asp:HiddenField ID="txtPedidoFixacao" runat="server" />
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            DADOS DA EMPRESA
                        </div>
                        <div class="painelleft">
                            <div class="row">
                                <div class="collbl">
                                    Empresa Recebedora:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlEmpresaPagadora" TabIndex="6" runat="server" Width="550px"
                                        OnSelectedIndexChanged="DdlEmpresaPagadora_SelectedIndexChanged" AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Tipo Rcbto:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlTiposDePagamentos" TabIndex="7" runat="server" Width="550px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Banco:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlBancoRecebedor" runat="server" Width="550px" OnSelectedIndexChanged="DdlBancoRecebedor_SelectedIndexChanged"
                                        AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Conta:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlContaRecebedora" TabIndex="8" runat="server" Width="550px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Moeda / Indexador:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlMoeda" runat="server" Width="127px" />
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
                                    <asp:DropDownList ID="DdlProvisoes" TabIndex="9" runat="server" Width="127px" AutoPostBack="True"
                                        OnSelectedIndexChanged="DdlProvisoes_SelectedIndexChanged" />
                                </div>
                                <div class="collbl">
                                    Movimento:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtMovimento" runat="server" Width="75px" Font-Bold="False" ClientIDMode="Static"
                                        data-ToolTip="default" ToolTip="Selecionar conforme recebimento do titulo." />
                                    <asp:HiddenField ID="hdnMovimentoOriginal" runat="server" />
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
                            <div class="row" runat="server">
                                <div class="collbl">
                                    Adiantamento:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNumeroAdto" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="cmdAdiantamento" OnClick="cmdAdiantamento_Click" runat="server" UseSubmitBehavior="False"
                                        CssClass="btn" Text=">" data-ToolTip="default" />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgLimparAdto" OnClick="imgLimparAdto_Click" runat="server"
                                        CssClass="btn" Width="22px" Height="22px" ImageUrl="~/images/borracha.JPG" ImageAlign="AbsMiddle"></asp:ImageButton>
                                </div>
                            </div>
                            <div class="row" runat="server">
                                <div class="collbl">
                                    Venc.Adto:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtVencimentoAdto" runat="server" Enabled="False" data-ToolTip="default"
                                        ToolTip="Vencimento do Adiantamento." CssClass="calendario" />
                                </div>
                                <div class="collbl">
                                    Taxa de Juro:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtTaxaAdto" CssClass="txtDecimal" Style="text-align: right" runat="server" Width="120px" data-ToolTip="default" ToolTip="Informar a taxa de juro." />
                                </div>
                            </div>
                            <div class="row">
                                <div id="rowContratoBanco" runat="server">
                                    <div class="collbl">
                                        Contrato Banco:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtContratoBanco" runat="server" Width="90px" Font-Bold="False"
                                            CssClass="texto" ClientIDMode="Static" MaxLength="20" data-ToolTip="default"
                                            ToolTip="Informar o número do contrato bancário." />
                                    </div>
                                </div>
                                <div id="DivFaturaDeFrete" runat="server" style="display: none;">
                                    <div class="collbl">
                                        Fatura de Frete:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtCodigoFaturaDeFrete" runat="server" CssClass="txtNumerico" Enabled="False"
                                            data-ToolTip="default" />
                                    </div>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Contrato Financeiro:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContratoFinanceiro" runat="server" MaxLength="30" data-ToolTip="default"
                                        ToolTip="Informar o número do contrato financeiro." />
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
                                <div class="coltxt" style="margin-left: 136px;">
                                    <asp:TextBox ID="txtHistorico" TabIndex="10" runat="server" TextMode="MultiLine"
                                        Style="width: 542px;" MaxLength="198" data-ToolTip="default" ToolTip="Selecionar o histórico adequado da operação." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Observação:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtObservacoes" runat="server" TextMode="MultiLine" TabIndex="20"
                                        Style="width: 542px;" data-ToolTip="default" ToolTip="Preencher quando houver alguma observação relevante." />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 350px;">
                            <div class="painelleft" style="width: 66%;">
                                <div class="row">
                                    <div class="collbl w100">
                                        Venc. Original:
                                    </div>
                                    <div class="coltxt">
                                        <asp:Label ID="lblVencOriginal" runat="server" Width="80px" Text="Label" Font-Bold="True" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl w100">
                                        Venc. Atual:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtProrrogacao" Width="84px" runat="server" AutoPostBack="True"
                                            ClientIDMode="Static" OnTextChanged="txtProrrogacao_TextChanged" data-ToolTip="default"
                                            ToolTip="Data do vencimento do recebimento." />
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
                                        Style="text-align: right" data-ToolTip="default" ToolTip="Inserir o valor total do recebimento." />
                                    <asp:HiddenField ID="txtValorDocumento" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="txtValorMoeda" runat="server" />
                                    <asp:TextBox ID="txtValorEmMoeda" Width="90px" CssClass="txtDecimal" runat="server"
                                        Style="text-align: right" data-ToolTip="default" ToolTip="Inserir o valor total do recebimento." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w100" style="height: auto; min-height: 26px; margin-bottom: 4px;">
                                    <asp:Label ID="lblDescontos" CssClass="lbls" runat="server" Text="Descontos:" />
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
                            <div class="row">
                                <div class="collbl w100" style="height: auto; min-height: 26px; margin-bottom: 4px;">
                                    <asp:Label ID="lblDeducoes" CssClass="lbls" runat="server" Text="Deduções:" />
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
                            <div class="row">
                                <div class="collbl w100" style="height: auto; min-height: 26px; margin-bottom: 4px;">
                                    <asp:Label ID="lblJuros" CssClass="lbls" runat="server" Text="Multa/Juros:" />
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
                            <div class="row">
                                <div class="collbl w100" style="height: auto; min-height: 26px; margin-bottom: 4px;">
                                    <asp:Label ID="lblAcrescimos" CssClass="lbls" runat="server" Text="Acréscimos:" />
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
                                        ReadOnly="True" Style="text-align: right" data-ToolTip="default" ToolTip="Valor líquido a receber após os descontos/acréscimos." />
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="txtValorLiquidoMoeda" runat="server" />
                                    <asp:TextBox ID="txtValorCobradoMoeda" Width="90px" CssClass="txtDecimal" runat="server"
                                        ReadOnly="True" Style="text-align: right" data-ToolTip="default" ToolTip="Valor líquido a receber após os descontos/acréscimos." />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="ImgCalcular" OnClick="ImgCalcular_Click1" runat="server" CssClass="btn"
                                        UseSubmitBehavior="False" Text="=" data-ToolTip="default" ToolTip="Calcular Valores" />
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel3" ID="TabPanel3" >
                    <HeaderTemplate>
                        Parcelas
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="GridPedidos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
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
                                    <asp:BoundField DataField="Baixa" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data da Baixa"
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
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Deducoes" DataFormatString="{0:N}" HeaderText="Dedu&#231;&#245;es"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Juros" DataFormatString="{0:N}" HeaderText="Juros" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Acrescimos" DataFormatString="{0:N}" HeaderText="Acr&#233;scimos"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
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
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel4" ID="TabPanel4">
                    <HeaderTemplate>
                        Consulta Títulos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultarTitulos" Text="Consultar" runat="server"
                                            OnClick="lnkConsultarTitulos_Click" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatorioCons" Text="Relatório" runat="server"
                                            OnClick="lnkRelatorioCons_Click" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparConsultaTitulos" Text="Limpar" runat="server"
                                            OnClick="lnkLimparConsultaTitulos_Click" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkAgrupar" runat="server" Text="Agrupar"
                                            OnClick="lnkAgrupar_Click" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkBaixar" Text="BaixarTodos" runat="server"
                                            OnClick="lnkBaixar_Click" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkReprogramar" runat="server" Text="Reprogramar"
                                            data-tooltip="default" ToolTip="Habilita campo para informar o novo vencimento." />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkEmiteReciboGeral" Text="Recibo" runat="server"
                                            OnClick="lnkEmiteReciboGeral_Click" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 70.5%;">
                            <div class="row">
                                <div class="collbl">
                                    Unidade:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="DdlUnidadeConsultaTitulos" runat="server" Width="596px" OnSelectedIndexChanged="DdlUnidadeConsultaTitulos_SelectedIndexChanged"
                                        AutoPostBack="True" />
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
                                        onclick="if($(this).attr('checked') == 'checked') msgbox('ATENÇÃO!\n\nCampo utilizado apenas para relatório.','ATENÇÃO!','Info');" />
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="txtCodigoClienteConsTitulo" runat="server" />
                                    <asp:TextBox ID="txtClienteConsTitulo" runat="server" Enabled="False" Width="557px" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="cmdConsultaClientes" OnClick="cmdConsultaClientes_Click" runat="server"
                                        CssClass="btn" Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Unificar os clientes pra consulta." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Pedido:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPedidoConsultaTitulos" runat="server" Width="85px" CssClass="txtNumerico"
                                        data-ToolTip="default" ToolTip="Número do pedido." />
                                </div>
                                <div class="coltxt">
                                    <asp:Button CssClass="btn" ID="btnBuscaPedidoConsultaTitulos" OnClick="btnBuscaPedidoConsultaTitulos_Click"
                                        runat="server" Text=">" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default"
                                        ToolTip="Número do pedido." />
                                </div>
                                <div class="collbl w100">
                                    Período Inicial:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" runat="server" Width="85px" CssClass="calendario"
                                        data-ToolTip="default" ToolTip="Informar o data inicial de consulta." />
                                </div>
                                <div class="collbl w100">
                                    Período Final:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" runat="server" Width="85px" CssClass="calendario"
                                        data-ToolTip="default" ToolTip="Informar o data final de consulta." />
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="lblAgrupar" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Nota:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNumNota" runat="server" Style="width: 326px;" data-ToolTip="default"
                                        ToolTip="Número da Nota Fiscal." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Parametros:
                                </div>
                                <div class="coltxt">
                                    <asp:CheckBox ID="chkObservacao" runat="server" Text="Emitir Historico com Observações"
                                        data-ToolTip="default" ToolTip="Selecionar a opção de acordo com a consulta desejada." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Título:
                                </div>
                                <div class="coltxt">
                                    <asp:RadioButton ID="RbGeral" runat="server" Text="Geral" OnCheckedChanged="RbGeral_CheckedChanged"
                                        Checked="True" GroupName="Situacao" AutoPostBack="True" data-ToolTip="default"
                                        ToolTip="Selecionar a situação do título para consulta." />
                                    <asp:RadioButton ID="RbBaixado" runat="server" Text="Baixado" OnCheckedChanged="RbBaixado_CheckedChanged"
                                        GroupName="Situacao" AutoPostBack="True" data-ToolTip="default" ToolTip="Selecionar a situação do título para consulta." />
                                    <asp:RadioButton ID="RbAtivo" runat="server" Text="Ativo" OnCheckedChanged="RbAtivo_CheckedChanged"
                                        GroupName="Situacao" AutoPostBack="True" data-ToolTip="default" ToolTip="Selecionar a situação do título para consulta." />&nbsp;
                                    <asp:RadioButton ID="RbCancelado" runat="server" Text="Cancelados" OnCheckedChanged="RbCancelado_CheckedChanged"
                                        GroupName="Situacao" AutoPostBack="True" data-ToolTip="default" ToolTip="Selecionar a situação do título para consulta." />
                                    <asp:CheckBox ID="chkPrevisao" runat="server" Checked="True" Text="Previsão" data-ToolTip="default"
                                        ToolTip="Selecionar a situação do título para consulta." />&nbsp;
                                    <asp:CheckBox ID="chkProvisao" runat="server" Text="Provisão" data-ToolTip="default"
                                        ToolTip="Selecionar a situação do título para consulta." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Totalização:
                                </div>
                                <div class="coltxt">
                                    <asp:RadioButton ID="RbDiaGeral" runat="server" Text="Diario" Checked="True" GroupName="Totalizacao"
                                        data-ToolTip="default" ToolTip="Informar se o total será por diário, filial ou carteira." />
                                    <asp:RadioButton ID="RbFilialDiario" runat="server" Text="Filial " GroupName="Totalizacao"
                                        data-ToolTip="default" ToolTip="Informar se o total será por diário, filial ou carteira." />
                                    <asp:RadioButton ID="RbCarteiraDia" runat="server" Text="Carteira" GroupName="Totalizacao"
                                        data-ToolTip="default" ToolTip="Informar se o total será por diário, filial ou carteira." />
                                </div>
                            </div>
                            <div id="Div1" class="row" runat="server" visible="False">
                                <div class="coltxt">
                                    <asp:Label ID="lblTotalRegistroAgrupado" runat="server" Font-Bold="True" />
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="txtRealDolar" runat="server" />
                                    <asp:HiddenField ID="HiddenIndexador" runat="server" />
                                </div>
                            </div>
                        </div>
                        <div class="painelright bordagrid" id="pnlReprogramaVencimentos" visible="False"
                            runat="server" style="width: 28.5%; height: auto; padding: 4px 2px; background-image: linear-gradient(to bottom, transparent, rgba(0,0,0,.3)); -webkit-box-shadow: -5px 6px 6px 2px rgba(0,0,0,0.85); -moz-box-shadow: -5px 6px 6px 2px rgba(0,0,0,0.85); box-shadow: -5px 6px 6px 2px rgba(0,0,0,0.85);">
                            <div class="subtitulodiv">
                                Novo Vencimento:
                            </div>
                            <div class="row" runat="server">
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNovoVencimento" runat="server" CssClass="calendario" Width="86px" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnNovoVencimento" runat="server" CssClass="botao" OnClick="btnNovoVencimento_Click"
                                        Text="Confirmar" data-ToolTip="default" ToolTip="Confirmar Reprogramação dos Vencimentos"
                                        UseSubmitBehavior="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnCancelarNonoVencimento" runat="server" CssClass="botao" Text="Cancelar"
                                        UseSubmitBehavior="False" />
                                </div>
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="GridConsultaTitulos" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridConsultaTitulos_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"
                                        HeaderText="Sel"></asp:CommandField>
                                    <asp:TemplateField>
                                        <HeaderTemplate>
                                            <asp:CheckBox ID="chkAllTitulos" runat="server" AutoPostBack="True" OnCheckedChanged="chkAllTitulos_CheckedChanged"
                                                data-ToolTip="default" ToolTip="Seleciona todos os títulos de mesma moeda e indexador." />
                                        </HeaderTemplate>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="ChkGridTitulos" runat="server" AutoPostBack="True" OnCheckedChanged="ChkGridTitulos_CheckedChanged"></asp:CheckBox>
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle Width="30px"></ItemStyle>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Registro" HeaderText="Titulo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencto"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="50px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle Width="250px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Historico" HeaderText="Hist&#243;rico" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Dolar" DataFormatString="{0:N}" HeaderText="Dólares">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right" Width="100px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Reais" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right" Width="100px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Moeda">
                                        <HeaderStyle Width="30px"></HeaderStyle>
                                        <ItemStyle Width="30px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Indexador" HeaderText="Ind">
                                        <HeaderStyle HorizontalAlign="Right" Width="20px"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right" Width="20px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Grupado" HeaderText="Agrupado">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" Width="30px" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <HeaderTemplate>
                                            <label>
                                                RP</label>
                                            <input id="chkAlls" onclick="selectAll(this);" type="checkbox" data-tooltip="default"
                                                title="Seleciona todos os títulos para efetuar uma reprogramação de vencimentos." />
                                        </HeaderTemplate>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkReprogramar" CssClass="chk" runat="server" data-tooltip="default"
                                                ToolTip="Reprogramar Vencimento" />
                                        </ItemTemplate>
                                        <HeaderStyle Height="23px" Width="23px" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel2" ID="TabPanel2">
                    <HeaderTemplate>
                        Contabilização
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="gridRazao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
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
                                    <asp:BoundField HeaderText="Movimento" DataField="Movimento">
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
                                    <asp:BoundField HeaderText="D&#233;bito" DataField="Debito">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Cr&#233;dito" DataField="Credito">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Saldo" DataField="Saldo">
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
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:ConsultaAdiantamentos ID="ucConsultaAdiantamentos" runat="server" />
    <uc:DestinoContabil ID="ucDestinoContabil" runat="server" />
</asp:Content>
