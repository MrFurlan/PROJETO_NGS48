<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuGerencial.ascx.vb" Inherits="NGS.Web.UI.mnuGerencial" %>
<asp:Menu ID="mnuGerencial" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapGerencial">
</asp:Menu>
<asp:SiteMapDataSource ID="smapGerencial" runat="server" SiteMapProvider="GerencialMapProvider"
    ShowStartingNode="False" />