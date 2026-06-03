<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AtivosXProcessos.aspx.vb" Inherits="NGS.Web.UI.AtivosXProcessos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 130px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngAtivosXProcessos" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAtivosXProcessos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Processo de Depreciação e Contabilização
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkProcessar" runat="server" Text="Processar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Eliminar" />
                        </li>
<%--                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkSintese" runat="server" Text="Sintese" />
                        </li>--%>
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="595px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="595px" OnSelectedIndexChanged="DdlEmpresa_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Última Atualização:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataUltAtualizacao" CssClass="calendario" Width="100px" Enabled="false"
                        runat="server" data-ToolTip="default" ToolTip="Data da ultima atualização do bem." />
                </div>
                <div class="collbl">
                    Próxima Atualização:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataProxAtualizacao" CssClass="calendario" Width="100px" Enabled="false"
                        runat="server" data-ToolTip="default" ToolTip="Data da proxima atualização do bem." />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="chkSeguro" runat="server" Text="Seguro" AutoPostBack="true" data-ToolTip="default" ToolTip="" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
