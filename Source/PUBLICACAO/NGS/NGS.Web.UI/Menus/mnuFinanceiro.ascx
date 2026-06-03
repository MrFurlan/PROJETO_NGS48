<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuFinanceiro.ascx.vb"
    Inherits="NGS.Web.UI.mnuFinanceiro" %>
<asp:Menu ID="mnuFinanceiro" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapFinanceiro">
</asp:Menu>
<asp:SiteMapDataSource ID="smapFinanceiro" runat="server" SiteMapProvider="FinanceiroMapProvider"
    ShowStartingNode="False" />
