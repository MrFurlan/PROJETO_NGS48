<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuComercial.ascx.vb"
    Inherits="NGS.Web.UI.mnuComercial" %>
<asp:Menu ID="mnuComercial" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapComercial">
</asp:Menu>
<asp:SiteMapDataSource ID="smapComercial" runat="server" SiteMapProvider="ComercialMapProvider"
    ShowStartingNode="False" />
