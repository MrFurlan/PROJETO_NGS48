<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucNFProdutor.ascx.vb" Inherits="NGS.Web.UI.ucNFProdutor" %>
<script type="text/javascript">
    function pageLoadConsultaNFProdutor() {
        $("#btnFecharUCProdutor", "#divConsultaNFProdutor").button();
    }

    $(document).ready(function () {
        pageLoadConsultaNFProdutor();
    });

    var prmConsultaNFProdutor = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaNFProdutor.add_endRequest(pageLoadConsultaNFProdutor);
</script>
<div id="divConsultaNFProdutor" class="uc" title="Nota Fiscal de Produtor" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaNFProdutor" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="row">
                <div class="bordagrid">
                    <asp:GridView ID="gridConsultaNFProdutor" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridConsultaNFProdutor_SelectedIndexChanged"
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
                            <asp:BoundField DataField="Movimento" HeaderText="Data" DataFormatString="{0:dd/MM/yyyy}">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Serie" HeaderText="Série">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Nota" HeaderText="Nota">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Valor" HeaderText="Valor" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <div class="row">
                <div class="painelright">
                    <asp:Button ID="btnFecharUCProdutor" runat="server" ClientIDMode="Static" Text="Fechar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>