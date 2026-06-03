<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuAdiantamento.ascx.vb"
    Inherits="NGS.Web.UI.mnuAdiantamento" %>
<asp:Menu ID="mnuAdiantamento" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapAdiantamento" StaticEnableDefaultPopOutImage="true"
    StaticPopOutImageUrl="~/Images/seta.gif" DynamicEnableDefaultPopOutImage="true"
    DynamicPopOutImageUrl="~/Images/seta.gif">
</asp:Menu>
<asp:SiteMapDataSource ID="smapAdiantamento" runat="server" SiteMapProvider="AdiantamentoMapProvider"
    ShowStartingNode="False" />
