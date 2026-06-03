<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaRomaneios.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaRomaneios" %>
<script type="text/javascript">
    function pageLoadConsultaRomaneios() {
        $("#btnFechar", "#divConsultaRomaneios").button();
    }

    $(document).ready(function () {
        pageLoadConsultaRomaneios();
    });

    var prmConsultaRomaneios = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaRomaneios.add_endRequest(pageLoadConsultaRomaneios);
</script>
<div id="divConsultaRomaneios" class="uc" title="Consulta de Romaneios" style="display: none;">
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
            <td>
                <asp:UpdatePanel ID="updConsultaRomaneios" runat="server" UpdateMode="Conditional">
                    <ContentTemplate>
                        <asp:HiddenField ID="HID" runat="server" />
                        <div class="bordagrid">
                            <asp:GridView ID="GridRomaneios" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:BoundField DataField="Romaneio" HeaderText="Romaneio" />
                                    <asp:BoundField DataField="ES" HeaderText="E/S" />
                                    <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data" />
                                    <asp:BoundField DataField="Pedido" HeaderText="Pedido" />
                                    <asp:BoundField DataField="Produto" HeaderText="Produto" />
                                    <asp:BoundField DataField="NomeDoProduto" HeaderText="Descri&#231;&#227;o" />
                                    <asp:BoundField DataField="Operacao" HeaderText="OP" />
                                    <asp:BoundField DataField="SubOperacao" HeaderText="SO" />
                                    <asp:BoundField DataField="Bruto" DataFormatString="{0:n0}" HeaderText="Bruto" />
                                    <asp:BoundField DataField="Desconto" DataFormatString="{0:n0}" HeaderText="Desconto" />
                                    <asp:BoundField DataField="Liquido" DataFormatString="{0:n0}" HeaderText="Liquido" />
                                    <asp:BoundField DataField="Laudo" HeaderText="Laudo" />
                                </Columns>
                            </asp:GridView>
                        </div>
<%--                        <table class="borda" width="100%" cellpadding="0" cellspacing="0">
                            <tr>
                                <td>
                                    <asp:Panel ID="pnlConsultaRomaneios" runat="server" Height="483px" ScrollBars="Vertical"
                                        Width="100%">
                                    </asp:Panel>
                                </td>
                            </tr>
                        </table>--%>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </td>
        </tr>
        <tr>
            <td align="right" style="padding-top: 10px;">
                <asp:UpdatePanel ID="updButton" runat="server">
                    <ContentTemplate>
                        <asp:Button ID="btnFechar" runat="server" ClientIDMode="Static" Text="Fechar" />
                    </ContentTemplate>
                </asp:UpdatePanel>
            </td>
        </tr>
    </table>
</div>
