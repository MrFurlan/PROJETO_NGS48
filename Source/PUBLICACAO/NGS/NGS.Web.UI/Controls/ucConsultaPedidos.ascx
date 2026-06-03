<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaPedidos.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaPedidos" %>
 
<div id="divConsultaPedidos" class="uc" title="Consulta de Pedidos" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaPedidos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconSair" runat="server">
                            <asp:LinkButton ID="lnkExcluir" Text="Fechar" runat="server" />
                        </li>
                        <li runat="server" style="float: right;">
                            <div id="c_search_psb_f" class="c_search_mc c_mf" style="width: 175px; border: 1px solid black;
                                padding-right: 2px; margin-top: 2px;">
                                <div class="row" style="margin-top: 0;">
                                    <div class="coltxt" style="width: 80%;">
                                        <asp:TextBox ID="txtConsultaPedido" runat="server" AutoPostBack="true" MaxLength="20"
                                            class="c_search_box c_ml" />
                                    </div>
                                    <div class="coltxt" style="float: right; width: 9%;">
                                        <input type="button" value="" style="background-image: url(Images/search.png); border: 0 none;
                                            width: 15px;" id="c_search_psb_go" class="c_search_go" data-tooltip="default"
                                            title="Pesquisa por Pedido" />
                                    </div>
                                </div>
                            </div>
                        </li>
                        <li runat="server" style="float: right;">
                            <div class="row" style="margin-top: 0;">
                                <div class="coltxt">
                                    <label>
                                        Safra:
                                    </label>
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlSafra" runat="server" Width="200px" AutoPostBack="true"
                                        OnSelectedIndexChanged="ddlSafra_SelectedIndexChanged" />
                                </div>
                            </div>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="bordagrid" style="height:442px;">
                <asp:GridView ID="GridPedidos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridPedidos_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                        <asp:BoundField DataField="CodigoEmpresa" HeaderText="Empresa" />
                        <asp:BoundField DataField="EnderecoEmpresa" HeaderText="End" />
                        <asp:BoundField DataField="Pedido" HeaderText="Pedido" />
                        <asp:BoundField DataField="DataPedido" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Dt.Pedido"
                            HtmlEncode="False" />
                        <asp:BoundField DataField="DataEntrega" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Dt.Entrega" />
                        <asp:BoundField DataField="Produto" HeaderText="Produto" HtmlEncode="False" />
                        <asp:BoundField DataField="NomeDoProduto" HeaderText="Descri&#231;&#227;o" HtmlEncode="False" />
                        <asp:BoundField DataField="Operacao" HeaderText="OP" />
                        <asp:BoundField DataField="SubOperacao" HeaderText="SO" />
                        <asp:BoundField DataField="DescricaoOperacao" HeaderText="Descrição da Operação">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Safra" HeaderText="Safra" />
                        <asp:BoundField DataField="Contratada" HeaderText="Contratada" DataFormatString="{0:n0}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Unitario" HeaderText="Unitário" DataFormatString="{0:n10}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Entregue" HeaderText="Entregue" DataFormatString="{0:n0}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Saldo" HeaderText="Saldo" DataFormatString="{0:n0}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
                <asp:GridView ID="gridPedidos2" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" Visible="False" AllowPaging="True"
																								
                    OnPageIndexChanging="gridPedidos2_PageIndexChanging" PageSize="12" OnSelectedIndexChanged="GridPedidos_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                        <asp:BoundField DataField="CodigoEmpresa" HeaderText="Empresa">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="EnderecoEmpresa" HeaderText="End">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Pedido" DataField="Pedido">
                            <HeaderStyle HorizontalAlign="Left" Width="60px" />												  
                            <ItemStyle HorizontalAlign="Left" Width="60px" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Efetivo" DataField="PedidoEfetivo">
                            <HeaderStyle HorizontalAlign="Left" Width="70px"  />												  
                            <ItemStyle HorizontalAlign="Left" Width="70px" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Itens">
                            <ItemTemplate>
                                <asp:Image ID="imgItens" runat="server" ImageUrl="~/Images/detalhes.png" data-tooltip="default"
                                    ToolTip='<%# Bind("Itens") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:BoundField HeaderText="SubOp" DataField="SubOp">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Cliente" DataField="NomeCliente">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Data Pedido" DataField="DataPedido" DataFormatString="{0:dd/MM/yyyy}">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Moeda" DataField="Moeda">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Qtde" DataField="Quantidade" DataFormatString="{0:N4}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField HeaderText="Total" DataField="Total" DataFormatString="{0:N2}">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
