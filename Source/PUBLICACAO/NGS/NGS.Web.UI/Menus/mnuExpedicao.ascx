<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuExpedicao.ascx.vb"
    Inherits="NGS.Web.UI.mnuExpedicao" %>
<asp:Menu ID="mnuExpedicao" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapExpedicao">
</asp:Menu>
<asp:SiteMapDataSource ID="smapExpedicao" runat="server" SiteMapProvider="ExpedicaoMapProvider"
    ShowStartingNode="False" />
