<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuRevenda.ascx.vb"
    Inherits="NGS.Web.UI.mnuRevenda" %>
<asp:Menu ID="mnuRevenda" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapRevenda">
</asp:Menu>
<asp:SiteMapDataSource ID="smapRevenda" runat="server" SiteMapProvider="RevendaMapProvider"
    ShowStartingNode="False" />
