<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaCep.ascx.vb" Inherits="NGS.Web.UI.ucConsultaCep" %>
<style type="text/css">
    .upper {
        text-transform: uppercase;
    }
</style>
<div id="divConsultaCep" class="uc" title="Consulta de CEP" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaCep" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Rua:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtRua" ClientIDMode="Static" runat="server" Width="519px" CssClass="upper" />
                </div>

            </div>
            <div class="row">
                <div class="collbluc">
                    Cidade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCidade" ClientIDMode="Static" runat="server" Width="280px" CssClass="upper" />
                </div>
                <div class="collbluc">
                    Estado:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEstado" ClientIDMode="Static" runat="server" CssClass="upper" />
                </div>
            </div>
            <div class="bordagrid" style="height: 300px;">
                <asp:GridView ID="gridCep" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True" />
                        <asp:BoundField DataField="cep" HeaderText="Cep">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="logradouro" HeaderText="Rua">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="bairro" HeaderText="Bairro">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="localidade" HeaderText="Cidade">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="uf" HeaderText="Estado">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ibge" HeaderText="Ibge">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>


