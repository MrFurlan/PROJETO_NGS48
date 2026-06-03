<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucFixacaoProcuracao.ascx.vb"
    Inherits="NGS.Web.UI.ucFixacaoProcuracao" %>
<div id="divFixacaoProcuracao" class="uc" title="Procurações À Fixar" style="display: none;">
    <asp:UpdatePanel ID="updpnlFixacaoProcuracao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <table style="width: 100%;">
                <tr>
                    <td>
                        <asp:Panel ID="pnlProcuracao" runat="server" ScrollBars="Vertical" Width="100%" Height="400px">
                            <asp:GridView ID="grdProcuracoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="98%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:LinkButton ID="lnkCessaoDeCredito" CssClass="lnk" 
                                                data-tooltip="default" ToolTip="Selecionar registro." runat="server" Text=" &gt; "
                                                OnClick="lnkCessaoDeCredito_Click">
                                                    <i class="fa fa-arrow-right seta"></i>
                                            </asp:LinkButton>
                                        </ItemTemplate>
                                        <HeaderStyle Width="50px" />
                                        <ItemStyle Width="50px" />
                                    </asp:TemplateField>
                                    <asp:BoundField HeaderText="Cessão Créd." DataField="Procuracao">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Cessionário" DataField="Cessionario">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Quantidade" DataField="Quantidade">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Fixado" DataField="Fixado">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Saldo" DataField="Saldo">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            </asp:GridView>
                        </asp:Panel>
                    </td>
                </tr>
                <tr>
                    <td align="right" style="padding: 10px 0px 0px 10px;">
                        <asp:Button ID="btnFechar" runat="server" CssClass="botao" Text="Fechar" />
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
