<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaProdutoCupomFiscal.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaProdutoCupomFiscal" %>
<script type="text/javascript">
    function pageLoadConsultaProduto() {
        $("#btnFechar", "#divConsultaProdutoCupomFiscal").button();
    }

    function buscarProduto(e) {
        if (e.keyCode == 13) {
            $("#btnProduto", "#divConsultaProdutoCupomFiscal").click();
            return false;
        }
    }

    $(document).ready(function () {
        pageLoadConsultaProduto();
    });

    var prmConsultaProduto = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaProduto.add_endRequest(pageLoadConsultaProduto);
</script>
<div id="divConsultaProdutoCupomFiscal" class="uc" title="Consulta de Produto" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaProduto" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNome" ClientIDMode="Static" runat="server" onkeypress="return buscarProduto(event);"
                        Width="550px" />
                </div>
<%--                <div class="coltxt">
                    <asp:Button ID="btnProduto" runat="server" ClientIDMode="Static" Text=">" CssClass="btn"
                        OnClick="btnProduto_Click" />
                </div>--%>
                <div class="coltxt">
                    <asp:LinkButton ID="lnkProduto" runat="server" Height="20px" Width="20px"><asp:Image ID="imgSearch" ImageUrl="~/Images/search.png" ImageAlign="Baseline" runat="server" data-ToolTip="default"
                            ToolTip="Consulta Produto" /></asp:LinkButton>
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodBarras" runat="server" MaxLength="80" Width="240px"
                        data-ToolTip="default" ToolTip="Código de barras" onkeypress="return buscarCodBarras(event);" />

                    <asp:LinkButton ID="lnkConsultaCodBarras" runat="server" Height="20px" Width="20px">
                        <asp:Image ID="Image1" ImageUrl="~/Images/search.png" ImageAlign="Baseline" runat="server" data-ToolTip="default"
                            ToolTip="Consulta Produto"  /></asp:LinkButton>
                </div>
            </div>
            <div class="row">
                <div class="bordagrid">
                    <asp:GridView ID="gridProduto" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridProduto_SelectedIndexChanged"
                        Width="100%">
                        <EditRowStyle BackColor="#999999" />
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                <HeaderStyle Width="20px" />
                                <ItemStyle Width="20px" />
                            </asp:CommandField>
                            <asp:BoundField DataField="Produto_Id" HeaderText="Produto">
                                <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                <ItemStyle Width="100px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Nome" HeaderText="Nome">
                                <HeaderStyle HorizontalAlign="Left" />
                            </asp:BoundField>
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
