<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaEncargosPlanoDeContas.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaEncargosPlanoDeContas" %>
<script type="text/javascript">
    function pageLoadConsultaEncargosPlanoDeContas() {
        $("#btnSelecionar", "#divConsultaEncargosPlanoDeContas").button();
        $("#btnFechar", "#divConsultaEncargosPlanoDeContas").button();
    }

    $(document).ready(function () {
        pageLoadConsultaEncargosPlanoDeContas();
    });

    var prmConsultaEncargosPlanoDeContas = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaEncargosPlanoDeContas.add_endRequest(pageLoadConsultaEncargosPlanoDeContas);
</script>
<div id="divConsultaEncargosPlanoDeContas" class="uc" title="Consulta de Encargos Plano de Contas"
    style="display: none;">
    <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
            <td>
                <asp:UpdatePanel ID="updpnlEncargosPlanoDeContas" runat="server" class="bordasimples"
                    style="overflow-x: auto; height: 450px;">
                    <ContentTemplate>
                        <asp:HiddenField ID="HID" runat="server" />
                        <asp:GridView ID="grd" runat="server" AutoGenerateColumns="False" CellPadding="4"
                            ForeColor="#333333" GridLines="None" Width="100%" DataKeyNames="CodigoContaEncargo">
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox runat="server" ID="chkCodigo" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="CodigoContaEncargo" HeaderText="Código" />
                                <asp:BoundField DataField="TituloEncargo" HeaderText="Título" />
                            </Columns>
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#999999" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        </asp:GridView>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </td>
        </tr>
        <tr>
            <td align="right" style="padding: 10px 0px 0px 10px;">
                <asp:UpdatePanel ID="updBotoes" runat="server">
                    <ContentTemplate>
                        <asp:Button ID="btnSelecionar" runat="server" ClientIDMode="Static" Text="Selecionar"
                            Width="80px" />
                        <asp:Button ID="btnFechar" runat="server" ClientIDMode="Static" Text="Fechar" Width="80px" />
                    </ContentTemplate>
                </asp:UpdatePanel>
            </td>
        </tr>
    </table>
</div>
