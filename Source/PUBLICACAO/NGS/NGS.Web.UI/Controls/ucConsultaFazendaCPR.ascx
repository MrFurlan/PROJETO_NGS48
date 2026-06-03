<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaFazendaCPR.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaFazendaCPR" %>
<script type="text/javascript">
    function pageLoadConsultaFazendaCPR() {
        $("#btnFazenda", "#divConsultaFazendaCPR").button();
        $("#btnConfirmar", "#divConsultaFazendaCPR").button();
        $("#btnSair", "#divConsultaFazendaCPR").button();

    }

    $(document).ready(function () {
        pageLoadConsultaFazendaCPR();
    });

    var prmConsultaFazendaCPR = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaFazendaCPR.add_endRequest(pageLoadConsultaFazendaCPR);



</script>
<style type="text/css">
    input[type='text']
    {
        height: 16px;
    }
    .Hide
    {
        display: none;
    }
</style>
<div id="divConsultaFazendaCPR" class="uc" title="Consulta de Fazendas" style="display: none;">
    <table class="bordasimples" width="100%">
        <tr>
            <td>
                <asp:UpdatePanel ID="upNewUpdatePanel" runat="server" UpdateMode="Conditional">
                    <ContentTemplate>
                        <asp:HiddenField ID="HID" runat="server" />
                        <table width="100%" class="borda">
                            <tr>
                                <td class="titulotabela" colspan="2">
                                    CPR x Fazenda
                                </td>
                            </tr>
                            <tr>
                                <td valign="top" style="width: 10%">
                                    <div class="headerGray">
                                        <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Cliente:</span>
                                        <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                    </div>
                                    <td>
                                        <asp:Label ID="lblCliente" runat="server" CssClass="primario" Text="Label" />
                                    </td>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <asp:GridView ID="gridMatricula" runat="server" CellPadding="4" ForeColor="#333333"
                                        GridLines="None" Width="450px" AutoGenerateColumns="False">
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <EditRowStyle BackColor="#999999" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:BoundField DataField="EndCliente" HeaderText="Endereço" ControlStyle-CssClass="Hide">
                                                <HeaderStyle CssClass="Hide" />
                                                <ItemStyle CssClass="Hide" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Fazenda" HeaderText="Fazenda">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Matricula_Id" HeaderText="Matrícula">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Area" HeaderText="Área">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:TemplateField HeaderText="Área CPR" HeaderStyle-HorizontalAlign="Right" ItemStyle-HorizontalAlign="Right">
                                                <ItemTemplate>
                                                    <asp:TextBox ID="txtArea" runat="server" onkeyup="validar_valor(this, 15);" value="0,00"
                                                        Style="text-align: right;" CssClass="txtDecimal" Text='<%# eval("AreaCpr") %>' />
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                        </Columns>
                                    </asp:GridView>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" style="height: 26px" colspan="2">
                                    <asp:Button ID="btnConfirmar" runat="server" Text="Confirmar" CssClass="botao" OnClick="btnConfirmar_Click" />
                                    <asp:Button ID="btnFechar" runat="server" CssClass="botao" Text="Fechar" ClientIDMode="Static"
                                        OnClick="btnFechar_Click" />&nbsp;
                                </td>
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </td>
        </tr>
    </table>
</div>
