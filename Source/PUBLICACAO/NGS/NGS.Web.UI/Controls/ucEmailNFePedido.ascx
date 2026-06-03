<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucEmailNFePedido.ascx.vb"
    Inherits="NGS.Web.UI.ucEmailNFePedido" %>
<script type="text/javascript">
    function pageLoadEmailNFePedido() {
        $("#MainContent_ucEmailNFePedido_btnEnviar", "#divEmailNFePedido").button();
        $("#MainContent_ucEmailNFePedido_btnFechar", "#divEmailNFePedido").button();

        $("#MainContent_ucEmailNFePedido_txtMensagem", "#divEmailNFePedido").keydown(function (e) {
            if (e.keyCode == 13) {
                e.preventDefault();
                return false;
            }
        });
    }

    $(document).ready(function () {
        pageLoadEmailNFePedido();
    });

    var prmEmailNFePedido = Sys.WebForms.PageRequestManager.getInstance();
    prmEmailNFePedido.add_endRequest(pageLoadEmailNFePedido);
</script>
<div id="divEmailNFePedido" class="uc" title="Envio de e-mail" style="display: none;">
    <asp:UpdatePanel ID="updpnlEmailNFePedido" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkEnviar" runat="server" UseSubmitBehavior="False" Text="Enviar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkFechar" runat="server" UseSubmitBehavior="False" Text="Fechar" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Destinatário:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDestinatario" runat="server" Width="380px" />
                    <asp:Button ID="btnAdicionar" runat="server" UseSubmitBehavior="False" Text=" + " />
                </div>
            </div>
            <div class="row">
                <div class="bordagrid" style="height: 150px;">
                    <asp:GridView ID="grd" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                        <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:BoundField DataField="EmailNFE" HeaderText="E-mail">
                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                            </asp:BoundField>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:ImageButton ID="imgExcluir" runat="server" ImageUrl="~/images/deletar.gif" CommandArgument='<%# Eval("EmailNFE") %>'
                                        Style="border: 0;" OnClick="imgExcluir_Click" OnClientClick="return confirm('Deseja realmente excluir este registro?');" />
                                </ItemTemplate>
                                <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Assunto:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAssunto" runat="server" MaxLength="100" Width="410px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Mensagem:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtMensagem" runat="server" Width="410px" Height="250px" TextMode="MultiLine"
                        MaxLength="500" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
