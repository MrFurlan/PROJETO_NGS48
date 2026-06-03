<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaNotasDeFrete.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaNotasDeFrete" %>
<script type="text/javascript">
    function pageLoadConsultaNotasDeFrete() {
        $("#btnFechar", "#divConsultaNotasDeFrete").button();
    }

    $(document).ready(function () {
        pageLoadConsultaNotasDeFrete();
    });

    var prmConsultaNotasDeFrete = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaNotasDeFrete.add_endRequest(pageLoadConsultaNotasDeFrete);
</script>
<div id="divConsultaNotasDeFrete" class="uc" title="Consulta de Notas de Frete" style="display: none;">
    <asp:UpdatePanel ID="updConsultaNotasDeFrete" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="bordagrid">
                <asp:GridView ID="gridNotasFrete" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridNotas_SelectedIndexChanged">
                    <EditRowStyle BackColor="#999999" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
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
                        <asp:BoundField DataField="ValorNotaOficial" HeaderText="Valor" DataFormatString="{0:N4}"
                            HtmlEncode="False" />
                        <asp:BoundField DataField="Unitario" HeaderText="Unitario" DataFormatString="{0:N4}"
                            HtmlEncode="False" />
                        <asp:BoundField DataField="Unidade" HeaderText="UN." />
                    </Columns>
                </asp:GridView>
            </div>
            <div class="row">
                <div class="bordagrid">
                    <asp:Button ID="btnFechar" runat="server" ClientIDMode="Static" Text="Fechar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
