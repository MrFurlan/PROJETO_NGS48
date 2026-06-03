<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeAnaliseDeProducao.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeAnaliseDeProducao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngRelatorioDeAnaliseDeProducao" runat="server" />
    <asp:UpdatePanel ID="updpnlRelatorioDeAnaliseDeProducao" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Análises de Produção
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
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlUnidadeDeNegocioOrigem" runat="server" Width="600px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlUnidadeDeNegocioOrigem_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresaClienteOrigem" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Fábrica:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlFabrica" runat="server" Width="600px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlGrupoDeProdutos" runat="server" Width="600px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlGrupoDeProdutos_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlProdutos" runat="server" Width="600px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" CssClass="calendario" TabIndex="5"
                        runat="server" Width="116px" data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" CssClass="calendario" TabIndex="6"
                        runat="server" Width="116px" CausesValidation="True" data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
