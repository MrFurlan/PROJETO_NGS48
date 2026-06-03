<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PosicaoDePedido.aspx.vb" Inherits="NGS.Web.UI.PosicaoDePedido" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngPosicaoDePedido" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelacaoDeNotasPisCofins" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Posição de Pedido - Financeiro
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" Text="Processar" runat="server" />
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
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged1"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" Width="96px" CssClass="calendario" data-ToolTip="default" ToolTip="Informar o período a ser considerado." />
                </div>
                <div class="collbl">
                    à:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" Width="96px" CssClass="calendario" data-ToolTip="default" ToolTip="Informar o período a ser considerado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="cbConsolidado" runat="server" Text="Consolidado" data-ToolTip="default" ToolTip="Filtrar por Consolidado ou por Vencimento." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="cbVencimento" Text="Vencimento" runat="server" data-ToolTip="default" ToolTip="Filtrar por Consolidado ou por Vencimento." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo Relatório:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoPosicao" Checked="true" GroupName="Relatorio" Text=" Posição "
                        runat="server" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoPlanilha" GroupName="Relatorio" Text=" Planilha " runat="server" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedidos:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdAbertos" Checked="true" GroupName="pedidos" Text="Abertos"
                        runat="server" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdTodos" GroupName="pedidos" Text="Todos" runat="server" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
