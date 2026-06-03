<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaPedido.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaPedido" %>
<div id="divConsultaPedido" class="uc" title="Consulta de Pedido" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaPedido" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:Panel ID="Panel1" runat="server" Height="135px" ScrollBars="Vertical">
                <asp:GridView ID="dgPedido" runat="server" CellPadding="4" ForeColor="#333333" GridLines="None"
                    AutoGenerateColumns="False" OnSelectedIndexChanged="dgPedido_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Empresa_Id" HeaderText="Empresa" />
                        <asp:BoundField DataField="EndEmpresa_Id" HeaderText=" " />
                        <asp:BoundField DataField="Nome" HeaderText=" " />
                        <asp:BoundField DataField="Safra" HeaderText="Safra" />
                        <asp:BoundField DataField="DataPedido" HeaderText="DataPedido" />
                        <asp:BoundField DataField="Pedido_Id" HeaderText="Pedido" />
                        <asp:BoundField DataField="PedidoEfetivo" HeaderText="Pedido(CN)" />
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente" />
                        <asp:BoundField DataField="EndCliente" HeaderText=" " />
                        <asp:BoundField DataField="ClienteNome" HeaderText=" " />
                        <asp:BoundField DataField="Complemento" HeaderText="Comp." />
                        <asp:BoundField DataField="DescMoeda" HeaderText="Moeda" />
                        <asp:BoundField DataField="Origem" HeaderText="Vlr Venda" DataFormatString="{0:N2}" />
                        <asp:BoundField DataField="Destino" HeaderText="Vlr Compra" DataFormatString="{0:N2}" />
                        <asp:BoundField DataField="Saldo" HeaderText="Saldo" DataFormatString="{0:N2}" />
                    </Columns>
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                </asp:GridView>
            </asp:Panel>
            <asp:Panel ID="Panel2" runat="server" ScrollBars="Vertical">
                <asp:GridView ID="dgItens" runat="server" CellPadding="4" ForeColor="#333333" GridLines="None"
                    AutoGenerateColumns="False">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <Columns>
                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" />
                        <asp:BoundField DataField="Descricao" HeaderText=" ">
                            <FooterStyle Width="150px" />
                            <HeaderStyle Width="150px" />
                            <ItemStyle Width="150px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="QuantidadePedido" HeaderText="Qtde" DataFormatString="{0:N4}" />
                        <asp:BoundField DataField="UnitarioOficial" HeaderText="Unit. Oficial" DataFormatString="{0:N2}" />
                        <asp:BoundField DataField="TotalOficial" HeaderText="Total Oficial" DataFormatString="{0:N2}" />
                        <asp:BoundField DataField="UnitarioMoeda" HeaderText="Unit. Moeda" DataFormatString="{0:N2}" />
                        <asp:BoundField DataField="TotalMoeda" HeaderText="Total Moeda" DataFormatString="{0:N2}" />
                    </Columns>
                </asp:GridView>
            </asp:Panel>
            <asp:Button ID="btnOK" runat="server" OnClick="btnOK_Click" Text="OK" Width="65px" />
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
