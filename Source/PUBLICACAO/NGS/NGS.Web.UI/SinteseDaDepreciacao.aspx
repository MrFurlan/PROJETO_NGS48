<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="SinteseDaDepreciacao.aspx.vb" Inherits="NGS.Web.UI.SinteseDaDepreciacao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngSinteseDaDepreciacao" runat="server">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlSinteseDaDepreciacao" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Sintese da Depreciação
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" runat="server" Text="Relatório" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="590px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidado" runat="server" Text="Consolidar" data-ToolTip="default" ToolTip="Unificar os dados por empresa." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="590px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlEmpresa_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Depreciação:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" data-ToolTip="default" ToolTip="Data de inicio da depreciação." />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="chkSeguro" runat="server" Text="Seguro" data-ToolTip="default" ToolTip="Selecionar para listar somente os dados de seguros." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
