<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucMDFeXEstado.ascx.vb"
    Inherits="NGS.Web.UI.ucMDFeXEstado" %>
<script type="text/javascript">
    function pageLoadMDFeXEstado() {
        $("#MainContent_ucMDFeXEstado_btnSelecionar", "#divMDFeXEstado").button();
        $("#MainContent_ucMDFeXEstado_btnFechar", "#divMDFeXEstado").button();
        $("input[type='checkbox']:visible", "#MainContent_ucMDFeXEstado_grd").change(function () {
            if ($(this).is(":checked") == true || $(this).is(":checked") == "true") {
                $(this).parent().parent().find("input[type=text].txtInteiro").prop('disabled', false).setMask("integer");
            } else {
                $(this).parent().parent().find("input[type=text].txtInteiro").prop('disabled', true).setMask("integer").val("0");
            }
        });
    }

    $(document).ready(function () {
        pageLoadMDFeXEstado();
    });

    var prmMDFeXEstado = Sys.WebForms.PageRequestManager.getInstance();
    prmMDFeXEstado.add_endRequest(pageLoadMDFeXEstado);
</script>
<style type="text/css">

</style>
<div id="divMDFeXEstado" class="uc" title="MDF-e x Estado" style="display: none;">
    <asp:UpdatePanel ID="updpnlMDFeXEstado" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="bordagrid">
                <asp:GridView ID="grd" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" DataKeyNames="Codigo">
                    <Columns>
                        <asp:TemplateField>
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemTemplate>
                                <asp:CheckBox runat="server" ID="chkCodigo" Checked='<%# Eval("Selecionado") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Seq.">
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="50px" />
                            <ItemTemplate>
                                <asp:TextBox ID="txtSequencia" runat="server" CssClass="txtInteiro" Text='<%# Eval("Sequencia") %>'
                                    Width="50px" Enabled="false" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Codigo" HeaderText="Código" ItemStyle-HorizontalAlign="Center"
                            ItemStyle-VerticalAlign="Middle" />
                        <asp:BoundField DataField="Descricao" HeaderText="Descrição" ItemStyle-HorizontalAlign="Center"
                            ItemStyle-VerticalAlign="Middle" />
                        <asp:BoundField DataField="Regiao" HeaderText="Região" ItemStyle-HorizontalAlign="Center"
                            ItemStyle-VerticalAlign="Middle" />
                    </Columns>
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                </asp:GridView>
            </div>
            <div class="painelright">
                <div class="row">
                    <div class="coltxt">
                        <asp:Button ID="btnSelecionar" runat="server" Text="Selecionar" UseSubmitBehavior="False"
                            Width="80px" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnFechar" runat="server" Text="Fechar" UseSubmitBehavior="False"
                            Width="80px" />
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
