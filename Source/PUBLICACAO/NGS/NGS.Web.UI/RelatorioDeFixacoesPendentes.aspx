<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeFixacoesPendentes.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeFixacoesPendentes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 140px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioDeFixacoesPendentes" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioDeFixacoesPendentes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório De Fixações Pendentes
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
                    <asp:CheckBox ID="ChkConsEmpresa" runat="server" Text="Consolidar Empresa:" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="618px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsCliente" runat="server" Text="Consolidar Cliente:" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="578px" Enabled="false" />
                    <asp:HiddenField ID="HDCodigoCliente" runat="server"></asp:HiddenField>
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaCliente" runat="server" Text=">" CausesValidation="False"
                        UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Unificar as informações por cliente." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlClasse" runat="server" Width="160px">
                        <asp:ListItem Selected="True">AMBOS</asp:ListItem>
                        <asp:ListItem>COMPRAS</asp:ListItem>
                        <asp:ListItem>VENDAS</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="300px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" Width="118px" Style="margin-right: 2px;" runat="server" data-ToolTip="default" ToolTip="Número do pedido a ser consultado." />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaPedido" runat="server" Text=">" CssClass="btn" CausesValidation="False"
                        UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Número do pedido a ser consultado." />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="ChkPeriodo" runat="server" AutoPostBack="True" Text="Usar Período:" data-ToolTip="default" ToolTip="Informar o período a ser considerado." />
                </div>
                <div class="coltxt">
                    <asp:Panel ID="pnlData" runat="server" Visible="False">
                        <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="112px" />
                        &nbsp;à
                        <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="112px" />
                    </asp:Panel>
                </div>
            </div>
            <div class="row">
                <uc:SelecaoProduto ID="ucSelecaoProdutoVenda" runat="server" />
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
