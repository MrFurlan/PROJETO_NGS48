<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeEstoqueDeTerceiros.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeEstoqueDeTerceiros" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrptRET" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updRET" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Relatório de Estoque em Poder de Terceiros
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" runat="server" Text="Relatório" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" runat="server" Text="Cons. Empresa" ToolTip="Consolidar o cnpj da empresa." data-ToolTip="default" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente."
                        runat="server" Text="Cons. Cliente" data-ToolTip="default" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" Width="560px" runat="server" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnBuscarClientes" runat="server" CssClass="btn" Text=">" data-ToolTip="default"
                        ToolTip="Unificar as informações por cliente." />
                </div>
            </div>
            <div class="row">
                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
            </div>
            <div class="row">
                <div class="collbl">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmPoder" runat="server" Width="160px">
                        <asp:ListItem>Todos</asp:ListItem>
                        <asp:ListItem>Em Nosso Poder</asp:ListItem>
                        <asp:ListItem>Poder de Terceiros</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkMostrarPedidos" runat="server" Text="Mostrar Pedido" data-ToolTip="default"
                        ToolTip="Dados pertencentes à empresa ou a terceiros" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ChkSomenteValor" runat="server" Visible="false" Text="Exibir Somente Cliente/Valor"
                        data-ToolTip="default" ToolTip="Dados pertencentes à empresa ou a terceiros" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data até:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataAte" CssClass="calendario" runat="server" Width="86px" data-ToolTip="default"
                        ToolTip="Selecionar a data de consulta." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
