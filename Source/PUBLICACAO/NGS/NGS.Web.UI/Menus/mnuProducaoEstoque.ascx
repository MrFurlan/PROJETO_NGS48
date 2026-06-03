<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuProducaoEstoque.ascx.vb"
    Inherits="NGS.Web.UI.mnuProducaoEstoque" %>
<asp:Menu ID="mnuProducaoEstoque" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapProducaoEstoque">
</asp:Menu>
<asp:SiteMapDataSource ID="smapProducaoEstoque" runat="server" SiteMapProvider="ProducaoEstoqueMapProvider"
    ShowStartingNode="False" />
