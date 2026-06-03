<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucNFReferencialSaida.ascx.vb" Inherits="NGS.Web.UI.ucNFReferencialSaida" %>
<script type="text/javascript">
    function pageLoadConsultaNFReferencialSaida() {
        $("#btnFecharUCNFReferencialSaida", "#divConsultaNFReferencialSaida").button();
    }

    $(document).ready(function () {
        pageLoadConsultaNFReferencialSaida();
    });

   <%-- function RadioCheck(rb) {
        var gv = document.getElementById("<%=gridConsultaNFReferencialSaida.ClientID%>");
        var rbs = gv.getElementsByTagName("input");
        for (var i = 0; i < rbs.length; i++) {
            if (rbs[i].type == "checked") {
                if (rbs[i].checked && rbs[i] != rb) {
                    rbs[i].checked = false;
                    break;
                }
            }
        }
    }--%>

    var prmConsultaNFReferencialSaida = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaNFReferencialSaida.add_endRequest(pageLoadConsultaNFReferencialSaida);
</script>
<div id="divConsultaNFReferencialSaida" class="uc" title="Nota Fiscal para Complementar/Referencial" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaNFReferencialSaida" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li id="liNovo" class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Confirmar" runat="server" />
                        </li>
                        <li class="iconFechar" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="bordagrid">
                    <asp:GridView ID="gridConsultaNFReferencialSaida" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="99%">
                        <EditRowStyle BackColor="#999999" />
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:TemplateField HeaderText="">
                                <ItemTemplate>
<%--                                    <asp:CheckBox ID="rdSelecionado" onclick="RadioCheck(this)" runat="server" />--%>
                                    <asp:CheckBox ID="rdSelecionado" runat="server" AutoPostBack="true" />

                                </ItemTemplate>
                                <HeaderStyle Width="30px" />
                                <ItemStyle Width="30px" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="Movimento" HeaderText="Data" DataFormatString="{0:dd/MM/yyyy}">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="EntradaSaida" HeaderText="E/S">
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
                            <asp:BoundField DataField="CFOP_Id" HeaderText="CFOP">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Produto_Id" HeaderText="Produto">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Sequencia_Id" HeaderText="Sequência">
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
        </ContentTemplate>
    </asp:UpdatePanel>
</div>