<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ImportaFolha.aspx.vb" Inherits="NGS.Web.UI.ImportaFolha" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script src="<%# ResolveUrl("~/Scripts/App/importafolha.js")%>" type="text/javascript"></script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmFolha" runat="server" EnableScriptGlobalization="true" EnableScriptLocalization="true" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlFolha" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Importação da Folha de Pagamento
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
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="lstUnidadeNegocio" runat="server" Width="575px" AutoPostBack="True" ClientIDMode="Static"
                        OnSelectedIndexChanged="lstUnidadeNegocio_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="lstEmpresa" runat="server" Width="575px" ClientIDMode="Static" />
                </div>
            </div>

            <div class="row">
                <div class="collbl">
                    Sistema:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSistema" runat="server" Width="576px" ClientIDMode="Static">
                        <asp:ListItem>PADRÃO</asp:ListItem>
                        <asp:ListItem>GCI</asp:ListItem>
                        <asp:ListItem>DOMINIO</asp:ListItem>
                    </asp:DropDownList>
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
