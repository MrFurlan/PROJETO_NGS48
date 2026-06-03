<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Importa.aspx.vb" Inherits="NGS.Web.UI.Importa" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngImporta" EnablePartialRendering="true" runat="server"
        EnableScriptGlobalization="True" EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlImporta" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Importações
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:Button ID="cmdImportar" runat="server" Width="250px" Text="Importar Titulos do Sistema Atual"
                        CssClass="botao" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:Button ID="cmdExportar" runat="server" Width="250px" Text="Exportar Titulos Para Sistema Atual"
                        CssClass="botao" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:Button ID="cmdImportaClientes" OnClick="cmdImportaClientes_Click" runat="server"
                        Width="250px" Text="Importa Clientes de outro Banco" CssClass="botao" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:Button ID="cmdImportaSinco" OnClick="cmdImportaSinco_Click" runat="server" Width="250px"
                        Text="Importa Lanctos Sistema Sinco" CssClass="botao" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:Button ID="Button1" OnClick="Button1_Click" runat="server" Width="250px" Text="Gera Saldo Sinco"
                        CssClass="botao" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:Label ID="txtMensagem" runat="server" Font-Size="9pt" ForeColor="#FF3300" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
