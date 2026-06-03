<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CompensacaoPedidoTroca.aspx.vb"
    MasterPageFile="~/Principal.Master" Inherits="NGS.Web.UI.CompensacaoPedidoTroca" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl
        {
            width: 145px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngCompensacaoPedidoTroca" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioVendaPorTroca" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Compensação Pedido Troca
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="Consultar">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
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
                        <div class="subtitulodiv">
                            VENDA
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbEmpresaOrigem" runat="server" Width="618px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                <asp:CheckBox ID="chkConsClienteVenda" runat="server" Text="Consolidar Cliente:" />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtClienteOrigem" runat="server" Width="581px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdBuscaClienteOrigem" OnClick="cmdBuscaClienteOrigem_Click" runat="server"
                                    CssClass="btn" Text=">" CausesValidation="False" UseSubmitBehavior="False" />
                                <asp:HiddenField ID="txtCodigoClienteOrigem" runat="server" />
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
                                <asp:Button ID="cmdBuscaPedido" OnClick="cmdBuscaPedido_Click" runat="server" Text=">"
                                    CssClass="btn" CausesValidation="False" UseSubmitBehavior="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt lg">
                                <uc:SelecaoProduto ID="ucSelecaoProdutoVenda" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="subtitulodiv">
                                COMPRA
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbEmpresaDestino" runat="server" Width="618px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                <asp:CheckBox ID="chkConsClienteCompra" runat="server" Text="Consolidar Cliente:" />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtClienteDestino" runat="server" ReadOnly="True" Width="585px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdBuscaClienteDestino" runat="server" CausesValidation="False" OnClick="cmdBuscaClienteDestino_Click"
                                    CssClass="btn" Text=">" UseSubmitBehavior="False" />
                                <asp:HiddenField ID="txtCodigoClienteDestino" runat="server" />
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
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdBuscaPedidoCompra" runat="server" OnClick="cmdBuscaPedidoCompra_Click"
                                    CssClass="btn" Text=">" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt lg">
                                <uc:SelecaoProduto ID="ucSelecaoProdutoCompra" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="subtitulodiv">
                                Venda / Compra
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Moeda Exibição Relatório:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdOficial" runat="server" Checked="True" GroupName="Moeda" Text="Moeda Oficial" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdMoeda" runat="server" GroupName="Moeda" Text="Moeda Extrangeira" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                <asp:CheckBox ID="ChkPeriodo" runat="server" AutoPostBack="True" Text="Usar Período:" />
                            </div>
                            <div class="coltxt">
                                <asp:Panel ID="pnlData" runat="server" Visible="False">
                                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="112px" />
                                    à
                                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="112px" />
                                </asp:Panel>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="Compensar">
                    <ContentTemplate>
                        <div class="titulodiv">
                            Compensação
                        </div>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="LnkCompensar" Text="Consultar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gdvPedido" runat="server" CellPadding="4" ForeColor="#333333" GridLines="None"
                                AutoGenerateColumns="false">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <HeaderTemplate>
                                            <asp:CheckBox ID="chkAll" runat="server" />
                                        </HeaderTemplate>
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkPedido" runat="server" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="EmpresaVenda" HeaderText="Empresa Venda" />
                                    <asp:BoundField DataField="EndEmpresaVenda" HeaderText="End" />
                                    <asp:BoundField DataField="PedidoVenda" HeaderText="Pedido Venda" />
                                    <asp:BoundField DataField="EmpresaCompra" HeaderText="Empresa Compra" />
                                    <asp:BoundField DataField="EndEmpresaCompra" HeaderText="End" />
                                    <asp:BoundField DataField="PedidoCompra" HeaderText="Pedido Compra" />
                                    <asp:BoundField DataField="ValorOficialVenda" HeaderText="Valor Oficial Venda" />
                                    <asp:BoundField DataField="ValorMoedaVenda" HeaderText="Valor Moeda Venda" />
                                    <asp:BoundField DataField="ValorOficialCompra" HeaderText="Valor Oficial Compra" />
                                    <asp:BoundField DataField="ValorMoedaCompra" HeaderText="Valor Moeda Compra" />
                                    <asp:BoundField DataField="CompensarOficial" HeaderText="Valor Oficial Compensar" />
                                    <asp:BoundField DataField="CompensarMoeda" HeaderText="Valor Moeda Compensar" />
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
