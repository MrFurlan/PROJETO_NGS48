<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaAutorizacaoDeRetirada.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaAutorizacaoDeRetirada" %>
<script type="text/javascript">
    function pageLoadConsultaAutorizacaoDeRetirada() {
        $("#<%=btnBuscar.ClientID%>", "#divConsultaAutorizacaoDeRetirada").button();
        $("#<%=btnFechar.ClientID%>", "#divConsultaAutorizacaoDeRetirada").button();
    }

    $(document).ready(function () {
        pageLoadConsultaAutorizacaoDeRetirada();
    });

    var prmConsultaAutorizacaoDeRetirada = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaAutorizacaoDeRetirada.add_endRequest(pageLoadConsultaAutorizacaoDeRetirada);
</script>
<div id="divConsultaAutorizacaoDeRetirada" class="uc" title="Consulta de Autorização de Retirada"
    style="display: none;">
    <asp:UpdatePanel ID="updpnlAutorizacaoDeRetirada" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="row">
                <div class="collbluc">
                    Ano:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAno" runat="server" MaxLength="4" Width="100px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnBuscar" runat="server" Text="Buscar" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridAutorizacao" runat="server" AutoGenerateColumns="False" OnSelectedIndexChanged="gridAutorizacao_SelectedIndexChanged"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Font-Size="Smaller" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Autorizacao" HeaderText="ID" />
                        <asp:BoundField DataField="CodigoPedido" HeaderText="Pedido" />
                        <asp:BoundField DataField="NomeClientePedido" HeaderText="Cliente" />
                        <asp:BoundField DataField="NomeClienteRetirada" HeaderText="Retirante">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="NomeAutorizante" HeaderText="Autorizante" DataFormatString="{0:N4}"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Taxa" HeaderText="Taxa" DataFormatString="{0:N6}" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="QuantidadeAutorizadaFisica" HeaderText="Autor.F&#237;sico"
                            DataFormatString="{0:N0}" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="QuantidadeEntregueFisica" HeaderText="Entregue F&#237;sico"
                            DataFormatString="{0:N0}" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SaldoFisico" HeaderText="Saldo F&#237;sico" DataFormatString="{0:N0}"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="QuantidadeAutorizadaFiscal" DataFormatString="{0:N0}"
                            HeaderText="Autor.Fiscal" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="QuantidadeEntregueFiscal" DataFormatString="{0:N0}" HeaderText="Entregue Fiscal"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="SaldoFiscal" DataFormatString="{0:N0}" HeaderText="Saldo Fiscal"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="CodigoPedidoServico" HeaderText="Pedido Servico">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="row">
                <div class="coltxt" style="float: right;">
                    <asp:Button ID="btnFechar" runat="server" UseSubmitBehavior="False" Text="Fechar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
