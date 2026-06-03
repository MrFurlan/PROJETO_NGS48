<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PedidosPendentes.aspx.vb" Inherits="NGS.Web.UI.PedidosPendentes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 205px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngPedidosPendentesDeVendas" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPedidosPendentesDeVendas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Pedidos Pendentes - Fiscalmente
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" runat="server" Text="Relatorio" />
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
                    Relatorios:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkRelAnalitico" runat="server" Checked="True" Text="Analitico Pedido"
                        data-ToolTip="default" ToolTip="Selecionar o relatório desejado." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkRelResumoClienteProduto" runat="server" Text="Cliente/Produto"
                        data-ToolTip="default" ToolTip="Selecionar o relatório desejado." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkRelResumoCliente" runat="server" Text="Cliente" data-ToolTip="default"
                        ToolTip="Selecionar o relatório desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Apresentação:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkConsolidarEmpresasApresentacao" runat="server" Text="Consolidar Empresas"
                        data-ToolTip="default" ToolTip="Selecionar se irá ser unificado por cliente ou por empresas." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkConsolidarClientesApresentacao" runat="server" Text="Consolidar Clientes"
                        data-ToolTip="default" ToolTip="Selecionar se irá ser unificado por cliente ou por empresas." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" AutoPostBack="True" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" data-ToolTip="default" ToolTip="Consolidar o cnpj da empresa." runat="server"
                        Text="Consolidar Empresa:" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlSafra" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe da Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlClasse" runat="server">
                        <asp:ListItem>COMPRAS</asp:ListItem>
                        <asp:ListItem>VENDAS</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkPeriodo" runat="server" Text="Usar Periodo, Pedidos Abertos de: "
                        AutoPostBack="True" />
                </div>
                <div class="coltxt">
                    <asp:Panel ID="pnlDataMovimento" runat="server" Visible="False" HorizontalAlign="Left">
                        <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="112px"
                            data-ToolTip="default" ToolTip="Selecionar a data dos pedidos em aberto." />
                        à:
                        <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="112px"
                            data-ToolTip="default" ToolTip="Selecionar a data dos pedidos em aberto." />
                    </asp:Panel>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" data-ToolTip="default" ToolTip="Consolidar o cpf/cnpj do cliente."
                        runat="server" Text="Consolidar Cliente:" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="568px" data-ToolTip="default"
                        ToolTip="Unificar as informações por cliente." />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdCliente" OnClick="cmdCliente_Click" runat="server" Text=">" CssClass="btn"
                        data-ToolTip="default" ToolTip="Unificar as informações por cliente." />
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
