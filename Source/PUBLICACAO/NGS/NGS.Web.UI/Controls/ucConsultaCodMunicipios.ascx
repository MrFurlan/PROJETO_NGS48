<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaCodMunicipios.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaCodMunicipios" %>
<script type="text/javascript">
    function runScriptConsultaCodMunicipios(e) {
        if (e.keyCode == 13) {
            $("#BtnOK", "#divConsultaCodMunicipios").click();
            return false;
        }
    }
</script>
<div id="divConsultaCodMunicipios" class="uc" title="Consulta de Municípios" style="display: none;">
    <asp:UpdatePanel ID="updConsultaCodMunicipios" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconSair" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkFechar" runat="server">
                                <span>Fechar</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Procurar:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtMunicipio" runat="server" Width="300px" onkeypress="return runScriptConsultaCodMunicipios(event)" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnOK" runat="server" ClientIDMode="Static" CssClass="btn" OnClick="BtnOK_Click"
                        Text=">" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridCodMunicipio" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridCodMunicipio_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                            <HeaderStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Munic&#237;pio" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
