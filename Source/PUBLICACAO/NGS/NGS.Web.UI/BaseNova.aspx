<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="BaseNova.aspx.vb" Inherits="NGS.Web.UI.BaseNova" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngBaseNova" runat="server" AsyncPostBackTimeout="50000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlBaseNova" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Ajustar Base Nova Nota Fiscal
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Planilha" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkGravar" runat="server" Text="Processar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="633px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="633px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt" style="width: 131px">
                    <asp:TextBox ID="txtDataInicial" runat="server" Width="100px" CssClass="calendario" />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" Width="100px" CssClass="calendario" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
