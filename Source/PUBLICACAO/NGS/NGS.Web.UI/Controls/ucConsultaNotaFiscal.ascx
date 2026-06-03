<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaNotaFiscal.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaNotaFiscal" %>
<script type="text/javascript">
    function pageLoadConsultaNotaFiscal() {
        $("#btnSelecionar", "#divConsultaNotaFiscal").button();
        $("#btnFechar", "#divConsultaNotaFiscal").button();
    }

    $(document).ready(function () {
        pageLoadConsultaNotaFiscal();
    });

    var prmConsultaNotaFiscal = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaNotaFiscal.add_endRequest(pageLoadConsultaNotaFiscal);
</script>
<div id="divConsultaNotaFiscal" class="uc" title="Consulta de Nota Fiscal" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaNotaFiscal" runat="server">
        <ContentTemplate>
            <script type="text/javascript">
                function selectAllCheckboxes(chkAll) {
                    var chk = $('#' + chkAll.id);
                    var checked = chk.attr('checked') == "checked";
                    $("input[type='checkbox']", "#MainContent_ucConsultaNotaFiscal_grdNotasDoPedido").not("#chkAll").each(function () {
                        $(this).attr("checked", checked);
                    });
                }
            </script>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="bordagrid" style="height: 115px; width: 950px;">
                <asp:GridView ID="grdNotasDoPedido" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" EmptyDataText="NENHUM REGISTRO ENCONTRADO"
                    ShowHeaderWhenEmpty="True">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <EmptyDataRowStyle VerticalAlign="Bottom" HorizontalAlign="Center" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                <input id="chkAll" onclick="selectAllCheckboxes(this);" type="checkbox" />
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkSelecionado" runat="server" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="center" Width="90px" />
                            <ItemStyle HorizontalAlign="center" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                Nota
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:Label ID="lblNota" runat="server" Text='<%# Eval("Nota_Id") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                Empresa
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:Label ID="lblEmpresa" runat="server" Text='<%# Eval("Empresa_Id") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                Cliente
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:Label ID="lblCliente" runat="server" Text='<%# Eval("Cliente_Id") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                E/S
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:Label ID="lblEntradaSaida" runat="server" Text='<%# Eval("EntradaSaida_Id") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                Produto
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:Label ID="lblProduto" runat="server" Text='<%# Eval("Produto_Id") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                CFOP
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:Label ID="lblCFOP" runat="server" Text='<%# Eval("CFOP_Id") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                Qtde
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:Label ID="lblQtde" runat="server" Text='<%# Eval("Quantidade") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <HeaderTemplate>
                                Valor
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:Label ID="lblValor" runat="server" Text='<%# Eval("Valor") %>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="row">
                <div class="painelright">
                    <asp:Button ID="btnSelecionar" runat="server" CssClass="botao" Text="Selecionar"
                        OnClick="btnSelecionar_Click" />
                    <asp:Button ID="btnFechar" runat="server" CssClass="botao" Text="Fechar" OnClick="btnFechar_Click" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
