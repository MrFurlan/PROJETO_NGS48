<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="mnuGestao.ascx.vb"
    Inherits="NGS.Web.UI.mnuGestao" %>
<asp:Menu ID="mnuGestao" runat="server" Orientation="Horizontal" CssClass="Menu"
    Width="100%" DataSourceID="smapGestao" StaticEnableDefaultPopOutImage="true"
    StaticPopOutImageUrl="~/Images/seta.gif" DynamicEnableDefaultPopOutImage="true"
    DynamicPopOutImageUrl="~/Images/seta.gif">
</asp:Menu>
<asp:SiteMapDataSource ID="smapGestao" runat="server" SiteMapProvider="GestaoMapProvider"
    ShowStartingNode="False" />
