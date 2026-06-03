<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucEncargosNFG.ascx.vb"
    Inherits="NGS.Web.UI.ucEncargosNFG" %>
<div id="divEncargosNFG" class="uc" title="Informe os encargos do produto" style="display: none;">
    <asp:UpdatePanel ID="updpnlEncargosNFG" runat="server">
        <ContentTemplate>
            <table style="width: 100%;">
                <tr>
                    <td>
                        <asp:HiddenField ID="HID" runat="server" />
                        <asp:HiddenField ID="hdfIndex" runat="server" />
                        <table class="actions" style="width: 100%;">
                            <tr>
                                <td id="Td1" class="iconConsultar" runat="server" style="width: 20%;">
                                    <asp:LinkButton ID="lnkRecarregar" runat="server" CssClass="lnkMenu">
                                        <span>Recarregar Encargos</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconLimpar" runat="server" style="width: 20%;">
                                    <asp:LinkButton ID="lnkLimpar" runat="server" CssClass="lnkMenu">
                                        <span>Limpar</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconFechar" runat="server" style="width: 20%;">
                                    <asp:LinkButton ID="lnkFechar" runat="server" CssClass="lnkMenu">
                                        <span>Fechar</span>
                                    </asp:LinkButton>
                                </td>
                                <td style="width: 40%; display: block;">
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td>
                        <table width="100%">
                            <tr>
                                <td style="width: 100px; white-space: nowrap;" class="primario" valign="top">
                                    <div class="headerGray">
                                        <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Produto:</span>
                                        <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                    </div>
                                </td>
                                <td>
                                    <asp:DropDownList ID="ddlProdutoSel" runat="server" Width="595px" AutoPostBack="true"
                                        OnSelectedIndexChanged="ddlProdutoSel_SelectedIndexChanged" />
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <asp:Panel ID="pnlEncargos" runat="server" CssClass="bordasimples" Width="100%" Height="300px"
                                        ScrollBars="Vertical">
                                        <asp:GridView ID="DgEncargos" runat="server" Width="100%" ForeColor="#333333" OnRowCommand="DgEncargos_RowCommand"
                                            OnRowCreated="DgEncargos_RowCreated" GridLines="None" CellPadding="4" AutoGenerateColumns="False">
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></FooterStyle>
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333"></RowStyle>
                                            <Columns>
                                                <asp:BoundField DataField="Codigo" HeaderText="Encargo">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="SituacaoTributaria" HeaderText="ClaICMS">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="SituacaoTributariaPISCOFINS" HeaderText="PisCofins" />
                                                <asp:BoundField DataField="SituacaoTributariaIPI" HeaderText="IPI" />
                                                <asp:BoundField DataField="CodigoOperacao" HeaderText="Op">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="CodigoSubOperacao" HeaderText="So">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="CodigoGrupoProduto" HeaderText="Grupo">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="CodigoProduto" HeaderText="Produto">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="EstadoOrigem" HeaderText="Origem">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:BoundField DataField="EstadoDestino" HeaderText="Destino">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:TemplateField HeaderText="Base">
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtBaseEncargoItem" runat="server" CssClass="txtDecimal" BorderStyle="None"
                                                            Enabled="False" Text='<%# Eval("Base", "{0:N2}") %>' Width="100px" />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:TemplateField HeaderText="Percentual">
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtPercentualEncargoItem" runat="server" CssClass="txtDecimal9"
                                                            BorderStyle="None" Enabled="False" Text='<%# Eval("Percentual", "{0:N9}") %>'
                                                            Width="100px" />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="PercentualExibicao" DataFormatString="{0:N9}" HeaderText="% Exib." />
                                                <asp:TemplateField HeaderText="Valor">
                                                    <ItemTemplate>
                                                        <asp:TextBox ID="txtValorEncargoItem" runat="server" CssClass="txtDecimal" Enabled="False"
                                                            Text='<%# Eval("Valor", "{0:N2}") %>' BorderStyle="None" Width="100px" />
                                                    </ItemTemplate>
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:TemplateField>
                                                <asp:BoundField DataField="Sinal" HeaderText="Sinal">
                                                    <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundField>
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:Button ID="btnEncargoItem" runat="server" CommandName="OK" UseSubmitBehavior="False"
                                                            Enabled="False" Text="OK" />
                                                    </ItemTemplate>
                                                </asp:TemplateField>
                                            </Columns>
                                            <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White"></PagerStyle>
                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></HeaderStyle>
                                            <EditRowStyle BackColor="#999999"></EditRowStyle>
                                        </asp:GridView>
                                    </asp:Panel>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
