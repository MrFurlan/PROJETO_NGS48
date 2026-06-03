<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaEstados.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaEstados" %>
<div id="divConsultaEstados" class="uc" title="Consulta de Estados" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaEstados" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="bordagrid">
                <asp:GridView ID="GridEstados" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridEstados_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Width="100%" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True">
                            <HeaderStyle Width="20px" />
                            <ItemStyle Width="20px" />
                        </asp:CommandField>
                        <asp:BoundField DataField="Codigo" HeaderText="UF">
                            <HeaderStyle HorizontalAlign="Left" Width="30px" />
                            <ItemStyle HorizontalAlign="Left" Width="30px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                            <HeaderStyle HorizontalAlign="Left" Width="250px" />
                            <ItemStyle HorizontalAlign="Left" Width="250px" />
                        </asp:BoundField>
                    </Columns>
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" CssClass="DataGridFixedHeader" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
