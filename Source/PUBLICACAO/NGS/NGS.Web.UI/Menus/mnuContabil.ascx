<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuContabil.ascx.vb"
    Inherits="NGS.Web.UI.mnuContabil" %>
<asp:Menu ID="mnuContabil" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapContabil">
</asp:Menu>
<asp:SiteMapDataSource ID="smapContabil" runat="server" SiteMapProvider="ContabilMapProvider"
    ShowStartingNode="False" />
