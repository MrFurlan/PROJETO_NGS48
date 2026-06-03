<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaClientesDireto.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaClientesDireto" %>
<script type="text/javascript">
    function pageLoadConsultaClienteDireto() {
        $("#btnFechar", "#divConsultaClienteDireto").button();
        $(".txtInteiro").setMask('integer');
        $(".txtDecimal").setMask('decimal');
    }

    $(document).ready(function () {
        pageLoadConsultaClienteDireto();
    });

    var prmConsultaClienteDireto = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaClienteDireto.add_endRequest(pageLoadConsultaClienteDireto);
</script>
<div id="divConsultaClienteDireto" class="uc" title="Consulta de Clientes" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaClientesDireto" runat="server">
        <ContentTemplate>
            <table style="width: 100%;">
                <tr>
                    <td>
                        <asp:Panel ID="pnlGridView" runat="server" Height="365px" Width="100%" ScrollBars="Vertical"
                            CssClass="bordasimples">
                            <asp:HiddenField ID="HID" runat="server" />
                            <asp:GridView ID="gdvClientes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gdvClientes_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:BoundField DataField="Nome" HeaderText="Nome" ReadOnly="True">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Complemento" HeaderText="Complemento" ReadOnly="True">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cidade" HeaderText="Cidade" ReadOnly="True">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="UF">
                                        <ItemTemplate>
                                            <%# Eval("Estado.Codigo")%>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoFormatado" HeaderText="C&#243;digo" ReadOnly="True">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoEndereco" HeaderText="End" ReadOnly="True">
                                        <HeaderStyle HorizontalAlign="Right" Width="30px" />
                                        <ItemStyle HorizontalAlign="Right" Width="30px" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </asp:Panel>
                    </td>
                </tr>
                <tr>
                    <td align="right">
                        <asp:Button ID="btnFechar" runat="server" Text="Fechar" UseSubmitBehavior="False"
                            ClientIDMode="Static" Width="80px" />
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
