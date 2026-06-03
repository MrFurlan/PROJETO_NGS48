<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaAdiantamentos.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaAdiantamentos" %>
<script type="text/javascript">
    function pageLoadConsultaAdiantamento() {
        $("#<%=btnFechar.ClientID%>", "#divConsultaAdiantamento").button();
    }

    $(document).ready(function () {
        pageLoadConsultaAdiantamento();
    });

    var prmConsultaAdiantamento = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaAdiantamento.add_endRequest(pageLoadConsultaAdiantamento);
</script>
<div id="divConsultaAdiantamento" title="Consulta de Adiantamentos" style="display: none;">
    <asp:UpdatePanel ID="updConsultaAdiantamento" runat="server">
        <ContentTemplate>
            <div>
                <div class="painelleft borda">

                    <div class="row">
                        <div class="coltxt">
                            <asp:RadioButton ID="rdTodos" runat="server" GroupName="x1" AutoPostBack="True" Text="Todos" Checked="true" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:RadioButton ID="rdComSaldo" runat="server" GroupName="x1" AutoPostBack="True" Text="Com Saldo" />
                        </div>
                    </div>
                    <div class="row">
                        <asp:RadioButton ID="rdSelecionaveis" runat="server" GroupName="x1" AutoPostBack="True" Text="Selecionaveis" />
                    </div>


                </div>

                <div class="painelleft borda" style="width: 86%; height: 96px;">

                    <div class="row">
                        <asp:Literal ID="LitPedido" runat="server" Text=""></asp:Literal>
                    </div>
                    <div class="row">
                        <asp:Literal ID="LitConta" runat="server" Text=""></asp:Literal>
                    </div>
                    <div class="row">
                        <asp:Literal ID="LitMoeda" runat="server" Text=""></asp:Literal>
                    </div>
                </div>
            </div>
            <div class="row">
                <asp:HiddenField ID="HID" runat="server" />
                <div class="bordagrid">
                    <asp:GridView ID="GridAdiantamentos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridAdiantamentos_SelectedIndexChanged">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                            <asp:BoundField DataField="NumeroDoAdto" HeaderText="Código">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Titulo" HeaderText="Titulo" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Pedido" HeaderText="Pedido" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Conta" HeaderText="Conta" />
                            <asp:BoundField DataField="Safra" HeaderText="Safra">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Vencimento" HeaderText="Vencimento" DataFormatString="{0:dd/MM/yyyy}">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Taxa" HeaderText="Taxa" DataFormatString="{0:N2}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Cifrao" HeaderText="Moeda" />
                            <asp:BoundField DataField="Adiantamento" HeaderText="Vlr.Adto" DataFormatString="{0:N2}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Juros" HeaderText="Juros" DataFormatString="{0:N2}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Baixas" HeaderText="Baixas" DataFormatString="{0:N2}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Saldo" HeaderText="Saldo" DataFormatString="{0:N2}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:Button ID="btnFechar" runat="server" Text="Fechar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
