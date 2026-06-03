<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PosicaoDeTransferencias.aspx.vb" Inherits="NGS.Web.UI.PosicaoDeTransferencias" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngPosicaoDeTransferencias" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPosicaoDeTransferencias" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Posição de Transferências
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
                    Unidade Origem:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlUnidadeDeNegocioOrigem" runat="server" Width="600px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlUnidadeDeNegocioOrigem_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa Origem:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresaClienteOrigem" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade Destino:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlUnidadeDeNegocioDestino" runat="server" Width="600px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlUnidadeDeNegocioDestino_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa Destino:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresaClienteDestino" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" CssClass="calendario" TabIndex="5"
                        runat="server" Width="96px" data-ToolTip="default" ToolTip="Informar o período inicial de consulta." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" CssClass="calendario" TabIndex="6"
                        runat="server" Width="96px" CausesValidation="True" data-ToolTip="default" ToolTip="Informar o período final de consulta." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
