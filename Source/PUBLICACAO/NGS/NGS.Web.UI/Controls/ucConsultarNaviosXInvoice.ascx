<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultarNaviosXInvoice.ascx.vb" Inherits="NGS.Web.UI.ucConsultarNaviosXInvoice" %>


<div id="divConsultarNaviosXInvoice" class="uc" title="Consultar Navios X Invoice" style="display: none;">
    <asp:UpdatePanel ID="updpnlucConsultarNaviosXInvoice" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconFechar" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Navio:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlNavios" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlNavios_SelectedIndexChanged"
                        Width="584px" />
                </div>
            </div>
            <div class="row">
                <div class="bordagrid">
                    <asp:GridView ID="gridNaviosXInvoice" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridNaviosXInvoice_SelectedIndexChanged"
                        Width="100%">
                        <EditRowStyle BackColor="#999999" />
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                <HeaderStyle Width="20px" />
                                <ItemStyle Width="20px" />
                            </asp:CommandField>
                            <asp:BoundField DataField="Codigo_Id" HeaderText="Invoice">
                                <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                <ItemStyle Width="100px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Navio_Id" HeaderText="Navio">
                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Left" Width="50px"></ItemStyle>
                            </asp:BoundField>
                            <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle Width="550px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="DataDeChegada" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data de chegada">
                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                            </asp:BoundField>
                            <asp:BoundField DataField="Observacao" HeaderText="Observac&#231;&#227;o">
                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
