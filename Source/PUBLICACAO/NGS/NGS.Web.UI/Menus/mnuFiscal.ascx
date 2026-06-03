<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuFiscal.ascx.vb"
    Inherits="NGS.Web.UI.mnuFiscal" %>
<asp:Menu ID="mnuFiscal" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapFiscal">
</asp:Menu>
<asp:SiteMapDataSource ID="smapFiscal" runat="server" SiteMapProvider="FiscalMapProvider"
    ShowStartingNode="False" />
