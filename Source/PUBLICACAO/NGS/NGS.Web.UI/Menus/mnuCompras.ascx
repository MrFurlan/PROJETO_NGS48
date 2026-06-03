<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuCompras.ascx.vb"
    Inherits="NGS.Web.UI.mnuCompras" %>
<asp:Menu ID="mnuCompras" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapCompras">
</asp:Menu>
<asp:SiteMapDataSource ID="smapCompras" runat="server" SiteMapProvider="ComprasMapProvider"
    ShowStartingNode="False" />
