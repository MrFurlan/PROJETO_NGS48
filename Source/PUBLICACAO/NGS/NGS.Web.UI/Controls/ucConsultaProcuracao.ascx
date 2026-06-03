<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaProcuracao.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaProcuracao" %>
<script type="text/javascript">
    function pageLoadConsultaProcuracao() {
        $("#btnFechar", "#divCessaoDeCredito").button();
    }

    $(document).ready(function () {
        pageLoadConsultaProcuracao();
    });

    var prmConsultaProcuracao = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaProcuracao.add_endRequest(pageLoadConsultaProcuracao);
</script>
<div id="divCessaoDeCredito" class="uc" title="Consulta de Cessão de Crédito" style="display: none">
    <asp:UpdatePanel ID="updConsultaCessaoDeCredito" runat="server" UpdateMode="Conditional">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="HIDTipo" runat="server" />
            <div class="row" runat="server">
                <div class="bordagrid" style="height: 483px; width: 99.8%;" runat="server">
                    <asp:GridView ID="GridProcuracao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridProcuracao_SelectedIndexChanged">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                            <asp:BoundField DataField="Procuracao" HeaderText="Cessão Crédito" HtmlEncode="False" />
                            <asp:BoundField DataField="Pedido" HeaderText="Pedido" />
                            <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data" />
                            <asp:BoundField DataField="Cliente" HeaderText="Cessionario" HtmlEncode="False" />
                            <asp:BoundField DataField="EndCliente" HeaderText="E" />
                            <asp:BoundField DataField="Nome" HeaderText="Nome" HtmlEncode="False" />
                            <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" DataFormatString="{0:N4}"
                                HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Realizado" HeaderText="Realizada" DataFormatString="{0:N4}"
                                HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Saldo" HeaderText="Saldo" DataFormatString="{0:N4}" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                    <asp:GridView ID="GridProcuracoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridProcuracoes_SelectedIndexChanged">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                            <asp:BoundField DataField="Codigo" HeaderText="Cessão Crédito" HtmlEncode="False" />
                            <asp:BoundField DataField="CodigoPedidoCedente" HeaderText="Pedido" />
                            <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data" />
                            <asp:BoundField DataField="CodigoCessionario" HeaderText="Cessionario" HtmlEncode="False" />
                            <asp:BoundField DataField="EnderecoCessionario" HeaderText="E" />
                            <asp:BoundField DataField="NomeCessionario" HeaderText="Nome" HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" DataFormatString="{0:n0}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Realizado" HeaderText="Realizada" DataFormatString="{0:n0}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Saldo" HeaderText="Saldo" DataFormatString="{0:n0}">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <div class="row">
                <div class="painelright">
                    <asp:Button ID="btnFechar" runat="server" ClientIDMode="Static" Text="Fechar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
