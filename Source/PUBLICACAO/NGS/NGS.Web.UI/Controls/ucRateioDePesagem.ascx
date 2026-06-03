<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucRateioDePesagem.ascx.vb"
    Inherits="NGS.Web.UI.ucRateioDePesagem" %>
<div id="divRateioDePesagem" class="uc" title="Informe os dados do rateio" style="display: none;">
    <asp:UpdatePanel ID="updpnlRateioDePesagem" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <table style="width: 100%;">
                <tr>
                    <td colspan="4" valign="middle">
                        <table class="actions" style="width: 100%;">
                            <tr>
                                <td id="tdBotaoNovo" class="iconIncluir" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkGravar" runat="server">
                                                                <span>Gravar</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconLimpar" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkLimpar" runat="server">
                                                                <span>Limpar</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconFechar" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkFechar" runat="server">
                                                                <span>Fechar</span>
                                    </asp:LinkButton>
                                </td>
                                <td style="display: block;">
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Cliente:</span>
                            <span style="margin-left: 51px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td colspan="3">
                        <asp:TextBox ID="txtCliente" runat="server" Width="585px" ReadOnly="True" />
                        <asp:Button ID="btnCliente" runat="server" Text=" > " Style="margin-left: 1px;" Enabled="False"
                            UseSubmitBehavior="False" /><asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Pedido:</span>
                            <span style="margin-left: 51px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:TextBox ID="txtPedido" runat="server" ReadOnly="True" />
                        <asp:Button ID="btnPedido" runat="server" Text=" > " Style="margin-left: 1px;" Enabled="False"
                            UseSubmitBehavior="False" />
                    </td>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Saldo:</span>
                        </div>
                    </td>
                    <td>
                        <asp:TextBox ID="txtSaldoPedido" runat="server" ReadOnly="True" />
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Depósito:</span>
                            <span style="margin-left: 51px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlDeposito" runat="server" Width="588px">
                        </asp:DropDownList>
                    </td>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Autorização:</span>
                        </div>
                    </td>
                    <td>
                        <asp:TextBox ID="txtAutorizacao" runat="server" ReadOnly="True" />
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Operação:</span>
                            <span style="margin-left: 51px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:DropDownList ID="ddlSubOperacao" runat="server" Width="588px" AutoPostBack="True">
                        </asp:DropDownList>
                    </td>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Quantidade:</span>
                            <span style="margin-left: 51px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:TextBox ID="txtQuantidade" runat="server" CssClass="txtNumerico" ReadOnly="True" />
                        <asp:Button ID="cmdQuantidade" runat="server" Text="Ok" UseSubmitBehavior="False" />
                    </td>
                </tr>
                <tr>
                    <td colspan="4">
                        <div class="bordasimples" style="min-height: 185px; max-height: 185px; overflow: auto;
                            width: 100%;">
                            <asp:GridView ID="gridRateio" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Codigo" HeaderText="Cliente">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nome" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="End" HeaderText="End">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Pedido" HeaderText="Pedido">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoDeposito" HeaderText="Deposito">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Produto" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Op" HeaderText="OP">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="SO" HeaderText="SO">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Autorizacao" HeaderText="Autorização">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Bruto" DataFormatString="{0:N0}" HeaderText="Bruto">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Desconto" DataFormatString="{0:N0}" HeaderText="Desconto">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Liquido" DataFormatString="{0:N0}" HeaderText="Líquido">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Principal">
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkPrincipal" runat="server" Checked='<%# Eval("Principal") %>'
                                                Enabled="False" />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Center" Width="60px" />
                                        <ItemStyle HorizontalAlign="Center" Width="60px" />
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluir" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente excluir o item?');"
                                                OnClick="imgExcluir_Click" Visible="False" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
