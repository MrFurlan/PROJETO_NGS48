<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="CodificaTexto.aspx.vb" Inherits="NGS.Web.UI.CodificaTexto" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngEstados" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlEstados" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Codifica Texto
            </div>
            <div class="row">
                <div class="collbl">
                    Texto:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTexto" Width="200px" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnGerar" runat="server" Text="Gerar" CssClass="botao" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnDecript" runat="server" Text="Decriptografar" CssClass="botao" />
                </div>

            </div>
            <div class="row">
                <div class="collbl">
                    resultado - 64Bits:
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblResult64Bits" runat="server" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
