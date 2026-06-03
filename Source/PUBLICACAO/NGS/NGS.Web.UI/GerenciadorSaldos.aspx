<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="GerenciadorSaldos.aspx.vb" Inherits="NGS.Web.UI.GerenciadorSaldos" %>

<asp:Content ID="scriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngGerenciadorSaldos" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlGerenciadorSaldos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Gerenciador de Saldos
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkGerarLote" Text="Gerar Lote" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluirLote" Text="Excluir Lote" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
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
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="618px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="618px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Estorno:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" Width="80px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Data que ocorreu a anulação do lançamento." />
                </div>
                <div class="collbl">
                    Data Ajuste:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" Width="80px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Data que ocorreu o ajuste." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
