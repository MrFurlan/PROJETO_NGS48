<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioVendaPorTroca.aspx.vb" Inherits="NGS.Web.UI.RelatorioVendaPorTroca" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioVendaPorTroca" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioVendaPorTroca" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório de Troca Analítico
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
                        </li>

                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="subtitulodiv">
                VENDA
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEmpresaOrigem" runat="server" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsClienteVenda" runat="server" Text="Cliente:" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClienteOrigem" runat="server" Width="581px" Enabled="false" />
                    <asp:HiddenField ID="txtCodigoClienteOrigem" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaClienteOrigem" OnClick="cmdBuscaClienteOrigem_Click" runat="server"
                        CssClass="btn" Text=">" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafraOrigem" runat="server" Width="300px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button CssClass="btn" ID="cmdBuscaPedido" OnClick="cmdBuscaPedido_Click" runat="server"
                        Text=">" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Localizar o número do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMoedaVenda" runat="server" Width="300px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProdutoVenda" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <div class="subtitulodiv">
                        COMPRA
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEmpresaDestino" runat="server" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsClienteCompra" runat="server" Text="Cliente:" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClienteDestino" runat="server" Enabled="false" Width="585px" />
                    <asp:HiddenField ID="txtCodigoClienteDestino" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaClienteDestino" runat="server" CausesValidation="False" OnClick="cmdBuscaClienteDestino_Click"
                        Text=">" UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafraDestino" runat="server" Width="300px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedidoCompra" runat="server" Enabled="False" />
                    <asp:Button ID="cmdBuscaPedidoCompra" runat="server" OnClick="cmdBuscaPedidoCompra_Click"
                        CssClass="btn" Text=" > " data-ToolTip="default" ToolTip="Localizar o número do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMoedaCompra" runat="server" Width="300px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProdutoCompra" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <div class="subtitulodiv">
                        Venda / Compra
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdOficial" runat="server" Checked="True" GroupName="Moeda" Text="Moeda Oficial" data-ToolTip="default" ToolTip="Selecionar o tipo de moeda." />
                    <asp:RadioButton ID="rdMoeda" runat="server" GroupName="Moeda" Text="Moeda Extrangeira" data-ToolTip="default" ToolTip="Selecionar o tipo de moeda." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Encargo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RdLiquido" runat="server" Checked="True" GroupName="encargo"
                        Text="Liquido" data-ToolTip="default" ToolTip="Selecionar se é encargo bruto ou líquido." />
                    <asp:RadioButton ID="rdBruto" runat="server" GroupName="encargo" Text="Bruto" data-ToolTip="default" ToolTip="Selecionar se é encargo bruto ou líquido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="ChkPeriodo" runat="server" AutoPostBack="True" Text="Usar Período:" />
                </div>
                <div class="coltxt">
                    <asp:Panel ID="pnlData" runat="server" Visible="False">
                        <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="112px" data-ToolTip="default" ToolTip="Marcar para inserir o período." />
                        &nbsp;à
                        <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="112px" data-ToolTip="default" ToolTip="Marcar para inserir o período." />
                    </asp:Panel>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo do Resumo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbResumoCompleto" runat="server" Checked="True" GroupName="grbresumo"
                        Text="Completo" data-ToolTip="default" ToolTip="Selecionar para resumir parcial ou completo." />
                    <asp:RadioButton ID="rbResumoParcial" runat="server" GroupName="grbresumo" Text="Parcial" data-ToolTip="default" ToolTip="Selecionar para resumir parcial ou completo." />
                </div>
            </div>
            <div class="subtitulodiv">
                <asp:CheckBox ID="chkValorizacao" runat="server" />
                <label>
                    Valorizacao
                </label>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Ref.:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataRef" runat="server" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Informar a data de referência." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Valor Ref.:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValorRef" runat="server" data-ToolTip="default"
                        ToolTip="Infomar o valor de referência." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Base De Calculo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdEntregue" runat="server" Checked="True" GroupName="base" Text="Entregue" data-ToolTip="default" ToolTip="Selecionar a base de cálculo entregue ou contratado." />
                    <asp:RadioButton ID="rdContratado" runat="server" GroupName="base" Text="Contratado" data-ToolTip="default" ToolTip="Selecionar a base de cálculo entregue ou contratado." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
