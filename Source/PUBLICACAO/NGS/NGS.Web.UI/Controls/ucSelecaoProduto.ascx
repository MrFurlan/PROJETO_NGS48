<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucSelecaoProduto.ascx.vb"
    Inherits="NGS.Web.UI.ucSelecaoProduto" %>
<script type="text/javascript">
    function pageLoadSelecaoProduto() {
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
        pageLoadSelecaoProduto();
    });

    var prmSelecaoProduto = Sys.WebForms.PageRequestManager.getInstance();
    prmSelecaoProduto.add_endRequest(pageLoadSelecaoProduto);
</script>
<asp:UpdatePanel ID="updpnlSelecaoProduto" runat="server">
    <ContentTemplate>
        <asp:HiddenField ID="HIDSelecaoProduto" runat="server" />
        <asp:HiddenField ID="HWhereProduto" runat="server" />
        <div id="<%=Me.ID%>" class="accordion">
            <h3>
                <asp:Label ID="lblNome" runat="server" Text="Seleção de Produtos" />
            </h3>
            <div style="height: 100%;">
                <div class="painelleft" style="width: 40.5%; margin-right: 4px;">
                    <div class="titulodiv">
                        Grupo
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:RadioButton ID="rdnivel1" runat="server" AutoPostBack="True" GroupName="nivel"
                                OnCheckedChanged="rdnivel1_CheckedChanged" Text="1 - Nivel" />
                        </div>
                        <div class="coltxt">
                            <asp:RadioButton ID="rdnivel10" runat="server" AutoPostBack="True" GroupName="nivel"
                                OnCheckedChanged="rdnivel10_CheckedChanged" Text="10 - Nivel" />
                        </div>
                        <div class="coltxt">
                            <asp:RadioButton ID="rdnivel100" runat="server" AutoPostBack="True" GroupName="nivel"
                                OnCheckedChanged="rdnivel100_CheckedChanged" Text="100 - Nivel" />
                        </div>
                        <div class="coltxt">
                            <asp:RadioButton ID="rdnivel10000" runat="server" AutoPostBack="True" GroupName="nivel"
                                OnCheckedChanged="rdnivel10000_CheckedChanged" Text="10000 - Nivel" />
                            <asp:HiddenField ID="HGrupo" runat="server" />
                        </div>
                    </div>
                    <div class="bordagrid" style="height: 250px; width: 100%;">
                        <asp:GridView ID="gridGrupo" runat="server" AutoGenerateColumns="False" CellPadding="4"
                            ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridGrupo_SelectedIndexChanged"
                            Width="100%">
                            <Columns>
                                <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chkGrupo" runat="server" AutoPostBack="True" Checked='<%# Eval("Selecionado") %>'
                                            OnCheckedChanged="chkGrupo_CheckedChanged" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Codigo" HeaderText="Grupo" />
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
                    <asp:Panel ID="pnlSelecaoProdutos" runat="server">
                        <div class="titulodiv">
                            Produto
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:Button ID="btnMarcarTodos" runat="server" OnClick="btnMarcarTodos_Click" Text="Marcar Todos"
                                    UseSubmitBehavior="False" CssClass="botao" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnDesmarcarTodos" runat="server" OnClick="btnDesmarcarTodos_Click"
                                    Text="Desmarcar Todos" UseSubmitBehavior="False" CssClass="botao" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnInverte" runat="server" OnClick="btnInverte_Click" Text="Inverter Seleção"
                                    CssClass="botao" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 250px;">
                            <asp:Panel ID="Panel5" runat="server">
                                <asp:GridView ID="gridProduto" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    Width="100%" ForeColor="#333333" GridLines="None">
                                    <Columns>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:CheckBox ID="chkProduto" runat="server" AutoPostBack="True" Checked='<%# Eval("Selecionado") %>'
                                                    OnCheckedChanged="chkProduto_CheckedChanged" />
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
