<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucSelecaoOperacoes.ascx.vb"
    Inherits="NGS.Web.UI.ucSelecaoOperacoes" %>
<script type="text/javascript">
    function pageLoadSelecaoOperacoes() {
        $("div.accordion").accordion({
            active: false,  
            collapsible: true,
            alwaysOpen: false,
            heightStyle: "content",
            autoHeight: false,
            clearStyle: true
        });
    }

    $(document).ready(function () {
        pageLoadSelecaoOperacoes();
    });

    var prmSelecaoOperacoes = Sys.WebForms.PageRequestManager.getInstance();
    prmSelecaoOperacoes.add_endRequest(pageLoadSelecaoOperacoes);
</script>
<asp:UpdatePanel ID="updpnlSelecaoOperacoes" runat="server">
    <ContentTemplate>
        <asp:HiddenField ID="HIDSelecaoOperacoes" runat="server" />
        <asp:HiddenField ID="HIDOperacoes" runat="server" />
        <asp:HiddenField ID="HWhereOperacoes" runat="server" />
        <div id="<%=Me.ID%>" class="accordion">
            <h3>
                <asp:Label ID="lblNome" runat="server" Text="Seleção de Operacoes" />
            </h3>
            <div style="height: 100%;">
                <div class="painelleft" style="width: 40.5%; margin-right: 4px;">
                    <div class="titulodiv">
                        Operações
                    </div>
                    <div class="bordagrid" style="height: 250px; width: 100%;">
                        <asp:GridView ID="gridOperacoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                            ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridOperacoes_SelectedIndexChanged"
                            Width="100%">
                            <Columns>
                                <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chkOperacoes" runat="server" AutoPostBack="True" Checked='<%# Eval("Selecionado") %>'
                                            OnCheckedChanged="chkOperacoes_CheckedChanged" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Codigo" HeaderText="Codigo" />
                                <asp:BoundField DataField="Descricao" HeaderText="Descrição" />
                            </Columns>
                            <EditRowStyle BackColor="#999999" />
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        </asp:GridView>
                    </div>
                </div>
                <div class="painelright" style="width: 58.5%;">
                    <asp:Panel ID="pnlSelecaoOperacoes" runat="server">
                        <div class="titulodiv">
                            SubOperações
                        </div>
                        <div class="bordagrid" style="height: 250px;">
                            <asp:Panel ID="Panel5" runat="server">
                                <asp:GridView ID="gridSubOperacoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    Width="100%" ForeColor="#333333" GridLines="None">
                                    <Columns>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:CheckBox ID="chkSubOperacoes" runat="server" AutoPostBack="True" Checked='<%# Eval("Selecionado") %>'
                                                    OnCheckedChanged="chkSubOperacoes_CheckedChanged" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Codigo" HeaderText="Código" />
                                        <asp:BoundField DataField="Descricao" HeaderText="Descrição" />
                                    </Columns>
                                    <EditRowStyle BackColor="#999999" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                </asp:GridView>
                            </asp:Panel>
                        </div>
                    </asp:Panel>
                </div>
            </div>
        </div>
    </ContentTemplate>
</asp:UpdatePanel>
