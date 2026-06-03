<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaContaCliente.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaContaCliente" %>
<script type="text/javascript">
    function pageLoadConsultaContaCliente() {
        $("#btnFechar", "#divConsultaContaClientes").button();
    }

    $(document).ready(function () {
        pageLoadConsultaContaCliente();
    });

    var prmConsultaContaCliente = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaContaCliente.add_endRequest(pageLoadConsultaContaCliente);
</script>
<div id="divConsultaContaClientes" class="uc" title="Consulta de Conta Cliente" style="display: none;">
    <div class="bordagrid">
        <asp:HiddenField ID="HID" runat="server" />
        <asp:GridView ID="gdvConta" runat="server" AutoGenerateColumns="False" CellPadding="4"
            ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gdvConta_SelectedIndexChanged">
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Left" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle BackColor="#999999" />
            <Columns>
                <asp:CommandField ButtonType="button" SelectText=" &gt; " ShowSelectButton="True"
                    ItemStyle-HorizontalAlign="Center" ItemStyle-VerticalAlign="Middle" />
                <asp:BoundField DataField="Conta_Id" HeaderText="Código" HeaderStyle-HorizontalAlign="Left" />
                <asp:BoundField DataField="Titulo" HeaderText="Descri&#231;&#227;o" HeaderStyle-HorizontalAlign="Left" />
            </Columns>
        </asp:GridView>
    </div>
    <div class="row">
        <div class="painelright">
            <asp:Button ID="btnFechar" runat="server" ClientIDMode="Static" Text="Fechar" />
        </div>
    </div>
</div>
