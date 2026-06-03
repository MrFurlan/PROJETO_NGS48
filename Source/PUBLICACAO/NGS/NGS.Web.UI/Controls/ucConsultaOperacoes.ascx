<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaOperacoes.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaOperacoes" %>
<script type="text/javascript">
    function pageLoadConsultaOperacoes() {
        $("#btnFechar", "#divConsultaOperacoes").button();
    }

    $(document).ready(function () {
        pageLoadConsultaOperacoes();
    });

    var prmConsultaOperacoes = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaOperacoes.add_endRequest(pageLoadConsultaOperacoes);
</script>
<div id="divConsultaOperacoes" class="uc" title="Consulta de Operações" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaOperacoes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="painelleft" style="width: 49%">
                <asp:ListBox ID="lstOperacoes" runat="server" AutoPostBack="True" Font-Size="12px"
                    Height="200px" Width="550px" />
            </div>
            <div class="painelright" style="width: 49%;">
                <div class="bordagrid" style="height: 198px; margin-top: 0px;">
                    <asp:GridView ID="grdSubOperacoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                            <asp:BoundField DataField="SubOperacao" HeaderText="C&#243;digo" HtmlEncode="False" />
                            <asp:BoundField DataField="Descricao" HeaderText="Sub-Opera&#231;&#227;o" HtmlEncode="False">
                                <HeaderStyle Width="400px" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="E/S">
                                <ItemTemplate>
                                    <asp:Label ID="lblES" runat="server" Text='<%# Eval("EntradaSaida") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="CodigoFiscal" HeaderText="Cfop" HtmlEncode="False" />
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
