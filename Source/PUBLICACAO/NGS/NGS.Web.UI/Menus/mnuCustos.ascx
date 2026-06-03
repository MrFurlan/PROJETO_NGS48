<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuCustos.ascx.vb"
    Inherits="NGS.Web.UI.mnuCustos" %>
<asp:Menu ID="mnuCustos" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapCustos">
</asp:Menu>
<asp:SiteMapDataSource ID="smapCustos" runat="server" SiteMapProvider="CustosMapProvider"
    ShowStartingNode="False" />
