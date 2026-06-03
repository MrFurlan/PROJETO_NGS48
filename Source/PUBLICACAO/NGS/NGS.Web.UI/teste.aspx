<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="teste.aspx.vb" Inherits="NGS.Web.UI.teste" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="screste" runat="server" />
    <asp:UpdatePanel runat="server">
        <ContentTemplate>
            <asp:ListBox ID="lst" runat="server"></asp:ListBox>
            <asp:Button ID="btnEnviar" runat="server" Text="enviar" />
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
