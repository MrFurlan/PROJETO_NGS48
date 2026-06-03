<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioMargemDeLucro.aspx.vb" Inherits="NGS.Web.UI.RelatorioMargemDeLucro" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioMargemDeLucro" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioMargemDeLucro" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório de margem de lucro bruto
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ckAgruparCliente" AutoPostBack="true" runat="server" Text="Agrupar Cliente" data-ToolTip="default" ToolTip="Marcar a opção desejada para gerar o relatório." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkVisualizarPedido" runat="server" Text="Visualizar Pedido" data-ToolTip="default" ToolTip="Marcar a opção desejada para gerar o relatório." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" ToolTip="Unidade de negócio empresarial." runat="server" Width="628px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" runat="server" Text="Cons. Empresa:"
                        data-ToolTip="default" ToolTip="Consolidar o cnpj da empresa." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="628px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" runat="server" Text=" Cons. Cliente:"
                        data-ToolTip="default" ToolTip="Consolidar o cpf/cnpj do cliente." />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="587px" ReadOnly="True" disabled="disabled" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button Style="margin-left: 1px;" ID="cmdBuscaCliente" OnClick="cmdBuscaCliente_Click"
                        CssClass="btn" runat="server" Text=">" UseSubmitBehavior="False" CausesValidation="False" data-ToolTip="default" ToolTip="Unificar as informações por cliente." />
                </div>
            </div>
            <div class="row">
                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="250px" />
                </div>
                <div class="collbl">
                    Marca:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMarca" runat="server" Width="250px" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
