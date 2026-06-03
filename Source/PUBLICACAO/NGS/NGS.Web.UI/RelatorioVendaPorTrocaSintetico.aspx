<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioVendaPorTrocaSintetico.aspx.vb" Inherits="NGS.Web.UI.RelatorioVendaPorTrocaSintetico" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioVendaPorTrocaSintetico" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioVendaPorTrocaSintetico" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdnControlePopup" runat="server" />
            <div class="titulodiv">
                Relatório de Troca Sintético
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
                    <asp:CheckBox ID="chkConsolidaClienteVenda" runat="server" Text="Cliente:" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClienteOrigem" runat="server" Width="585px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoClienteOrigem" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaClienteOrigem" OnClick="cmdBuscaClienteOrigem_Click" runat="server"
                        Text=">" CausesValidation="False" UseSubmitBehavior="False" CssClass="btn"
                        data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
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
                    <asp:CheckBox ID="chkConsolidarClienteCompra" runat="server" Text="Cliente:" />
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoClienteDestino" runat="server" />
                    <asp:TextBox ID="txtClienteDestino" runat="server" Enabled="false" Width="585px" />
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
                <div class="collbl" style="width: 170px;">
                    Moeda Exibição Relatório:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdOficial" runat="server" Checked="True" GroupName="Moeda" Text="Moeda Oficial" data-ToolTip="default" ToolTip="Selecionar o tipo de moeda." />
                    <asp:RadioButton ID="rdMoeda" runat="server" GroupName="Moeda" Text="Moeda Estrangeira" data-ToolTip="default" ToolTip="Selecionar o tipo de moeda." />
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 170px;">
                    Modelo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbCliente" runat="server" Checked="True" GroupName="tipo" Text="Por Cliente" data-ToolTip="default" ToolTip="Selecionar a opção desejada." />
                    &nbsp;
                    <asp:RadioButton ID="rbGrupo" runat="server" GroupName="tipo" Text="Por Grupo Produto" data-ToolTip="default" ToolTip="Selecionar a opção desejada." />
                    &nbsp;
                    <asp:RadioButton ID="rbPedido" runat="server" GroupName="tipo" Text="Por Pedido" data-ToolTip="default" ToolTip="Selecionar a opção desejada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 170px;">
                    Encargo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdLiquido" runat="server" Checked="True" GroupName="Encargo"
                        Text="Liquido" data-ToolTip="default" ToolTip="Selecionar se é encargo bruto ou líquido." />
                    <asp:RadioButton ID="rdBruto" runat="server" GroupName="Encargo" Text="Bruto"
                        data-ToolTip="default" ToolTip="Selecionar se é encargo bruto ou líquido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 170px;">
                    <asp:CheckBox ID="chkPeriodo" runat="server" AutoPostBack="True" Text="Usar Período" />
                </div>
                <div class="coltxt">
                    <asp:Panel ID="pnlData" runat="server" BorderStyle="None" Font-Bold="True" HorizontalAlign="Left"
                        Visible="False">
                        <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="112px"
                            data-ToolTip="default" ToolTip="Marcar para inserir o período." />
                        &nbsp;&nbsp;à&nbsp;
                        <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="112px"
                            data-ToolTip="default" ToolTip="Marcar para inserir o período." />
                    </asp:Panel>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
