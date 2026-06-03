<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaNotaVendaAOrdem.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaNotaVendaAOrdem" %>
<script type="text/javascript">
    function pageLoadConsultaNotaVendaAOrdem() {
        $("#MainContent_ucConsultaNotaVendaAOrdem_btnBuscar", "#divConsultaNotaVendaAOrdem").button();
        $("#MainContent_ucConsultaNotaVendaAOrdem_btnFechar", "#divConsultaNotaVendaAOrdem").button();
    }

    $(document).ready(function () {
        pageLoadConsultaNotaVendaAOrdem();
    });

    var prmConsultaNotaVendaAOrdem = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaNotaVendaAOrdem.add_endRequest(pageLoadConsultaNotaVendaAOrdem);
</script>
<div id="divConsultaNotaVendaAOrdem" class="uc" title="Consulta de Nota Venda A Ordem"
    style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaNotaVendaAOrdem" runat="server">
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
                    <asp:Button ID="btnBuscar" runat="server" Text="Buscar" OnClick="btnBuscar_Click" />
                </div>
            </div>
            <div class="bordagrid" style="height: 500px;">
                <asp:GridView ID="gridNotas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridNotas_SelectedIndexChanged">
                    <EditRowStyle BackColor="#999999" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Movimento" HeaderText="Movimento" />
                        <asp:BoundField DataField="Empresa_Id" HeaderText="Empresa" />
                        <asp:BoundField DataField="EndEmpresa_Id" HeaderText="End" />
                        <asp:BoundField DataField="EmpresaNome" HeaderText="Nome" />
                        <asp:BoundField DataField="EmpresaCidade" HeaderText="Cidade" />
                        <asp:BoundField DataField="EmpresaEstado" HeaderText="Estado" />
                        <asp:BoundField DataField="Cliente_ID" HeaderText="Cliente" />
                        <asp:BoundField DataField="EndCliente_Id" HeaderText="End" />
                        <asp:BoundField DataField="Nome" HeaderText="Nome" />
                        <asp:BoundField DataField="Cidade" HeaderText="Cidade" />
                        <asp:BoundField DataField="Estado" HeaderText="Estado" />
                        <asp:BoundField DataField="EntradaSaida_Id" HeaderText="E/S" />
                        <asp:BoundField DataField="Nota_Id" HeaderText="Nota" />
                        <asp:BoundField DataField="Serie_Id" HeaderText="Serie" />
                        <asp:BoundField DataField="Produto_Id" HeaderText="Produto" />
                        <asp:BoundField DataField="NomeProduto" HeaderText="Nome" />
                        <asp:BoundField DataField="QtdeNota" HeaderText="Qtde" DataFormatString="{0:N4}"
                            HtmlEncode="False" />
                        <asp:BoundField DataField="Entregue" HeaderText="Entregue" DataFormatString="{0:N4}"
                            HtmlEncode="False" />
                        <asp:BoundField DataField="Unidade" HeaderText="UN." />
                    </Columns>
                </asp:GridView>
            </div>
            <div class="row">
                <div class="coltxt" style="float: right;">
                    <asp:Button ID="btnFechar" runat="server" Text="Fechar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
