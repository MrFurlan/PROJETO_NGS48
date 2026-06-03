<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucPedidoEncargo.ascx.vb"
    Inherits="NGS.Web.UI.ucPedidoEncargo" %>
<div id="divPedidoEncargo" class="uc" title="Encargos do Item do Pedido" style="display: none;">
    <asp:UpdatePanel ID="updpnlPedidoEncargo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="HIDLinhaProduto" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton CssClass="iconConsultar" ID="lnkRecarregarEncargos" runat="server" Text="Recarregar" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton CssClass="iconFechar" ID="LnkFechar" runat="server" Text="Fechar" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProdutoEncargo" runat="server" Width="500px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataEncargos" runat="server" Width="75px" CssClass="calendario"
                        Enabled="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conf. Operação:
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblConfigOperacao" runat="server" />
                </div>

                <div class="coltxt">
                    <asp:Button ID="btnAlterarOP" runat="server" Text="Alterar operação" UseSubmitBehavior="False" CssClass="btn"
                        data-ToolTip="default" ToolTip="click para editar a operação." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkAjustarReais" runat="server" AutoPostBack="True" Font-Bold="True"
                        Visible="False" OnCheckedChanged="chkAjustarReais_CheckedChanged" ToolTip="Mostrar valor em reais." Text="Ajustar valor em reais" />
                </div>
            </div>

            <div class="bordagrid" style="height: auto;">
                <asp:GridView ID="GridEncargos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="CodigoEncargo" HeaderText="Encargo" ReadOnly="True">
                            <FooterStyle HorizontalAlign="Left"></FooterStyle>
                            <HeaderStyle HorizontalAlign="Left" Width="120px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="120px"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Base">
                            <ItemTemplate>
                                <asp:TextBox ID="txtBase" runat="server" Enabled="False" CssClass="txtDecimal" Text='<%# Bind("Base", "{0:N2}") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" Width="120px" />
                            <ItemStyle HorizontalAlign="Right" Width="120px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Perc %">
                            <ItemTemplate>
                                <asp:TextBox ID="txtPercentual" runat="server" Enabled="False" CssClass="txtDecimal9" Text='<%# Bind("Percentual", "{0:N9}") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" Width="120px" />
                            <ItemStyle HorizontalAlign="Right" Width="120px" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Valor">
                            <ItemTemplate>
                                <asp:TextBox ID="txtValor" runat="server" Enabled="False" CssClass="txtDecimal" Text='<%# Bind("Valor", "{0:N2}") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" Width="120px" />
                            <ItemStyle HorizontalAlign="Right" Width="120px" />
                        </asp:TemplateField>
                        <asp:TemplateField ShowHeader="False">
                            <ItemTemplate>
                                <asp:Button ID="btnAlterarEncargo" runat="server" OnClick="btnAlterarEncargo_Click" Style="width: auto !important;"
                                    Text="Alterar" />
                                <asp:Button ID="btnSalvarEncargo" runat="server" CausesValidation="True" OnClientClick=" if(!confirm('Deseja realmente alterar o encargo?')) return false;"
                                    Text="OK" Visible="False" OnClick="btnSalvarEncargo_Click" Style="width: auto !important;" />
                                <asp:Button ID="btnCancelarEncargo" runat="server" Text="Cancelar" Visible="False" Style="width: auto !important;"
                                    OnClick="btnCancelarEncargo_Click" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" Width="100px" />
                            <ItemStyle HorizontalAlign="Center" Width="100px" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
