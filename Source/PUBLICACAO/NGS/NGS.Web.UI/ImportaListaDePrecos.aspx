<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ImportaListaDePrecos.aspx.vb" Inherits="NGS.Web.UI.ImportaListaDePrecos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script src="<%# ResolveUrl("~/Scripts/App/ImportaListaDePrecos.js")%>" type="text/javascript"></script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmImportaListaDePrecos" runat="server" EnableScriptGlobalization="true" EnableScriptLocalization="true" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlImportaListaDePrecos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Importação da Lista de Precos
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkImportar" Text="Importar" runat="server" ClientIDMode="Static" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="575px" ClientIDMode="Static" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Movimento:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtMovimnto" CssClass="calendario" runat="server" Width="96px" data-ToolTip="default" ToolTip="Informar o data para a Lista de Preços." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Arquivo:
                </div>
                <div class="coltxt">
                    <input id="filUpload" style="width: 575px" type="file" runat="server" clientidmode="Static" />
                </div>
            </div>
        </ContentTemplate>
        <Triggers>
            <asp:PostBackTrigger ControlID="lnkImportar" />
        </Triggers>
    </asp:UpdatePanel>
</asp:Content>