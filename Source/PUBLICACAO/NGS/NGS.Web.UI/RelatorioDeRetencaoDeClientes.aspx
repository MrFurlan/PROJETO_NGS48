<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeRetencaoDeClientes.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeRetencaoDeClientes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngEstados" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlTrait" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Relatório De Retenção de Clientes
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." runat="server" Text='Empresa:' />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="595px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cultura:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCultura" runat="server" Width="595px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra 01:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra01" runat="server" Width="595px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra 02:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra02" runat="server" Width="595px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra 03:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra03" runat="server" Width="595px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Marca do Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMarcaProd" runat="server" Width="595px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
