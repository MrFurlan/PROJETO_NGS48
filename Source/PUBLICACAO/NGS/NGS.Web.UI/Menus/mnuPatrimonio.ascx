<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuPatrimonio.ascx.vb"
    Inherits="NGS.Web.UI.mnuPatrimonio" %>
<asp:Menu ID="mnuPatrimonio" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapPatrimonio">
</asp:Menu>
<asp:SiteMapDataSource ID="smapPatrimonio" runat="server" SiteMapProvider="PatrimonioMapProvider"
    ShowStartingNode="False" />
