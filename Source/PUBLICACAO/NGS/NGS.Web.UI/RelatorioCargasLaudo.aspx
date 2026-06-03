<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioCargasLaudo.aspx.vb" Inherits="NGS.Web.UI.RelatorioCargasLaudo" %>

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
                Totalizador Diário de Cargas e Descargas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" runat="server">
                                <span>Relatório</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkHtml" runat="server">
                                <span>Html</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server">
                                <span>Limpar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server">
                                <span>Ajuda</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlUnidadeDeNegocio" runat="server" Width="600px" OnSelectedIndexChanged="DdlUnidadeDeNegocio_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" runat="server" Text="Consolidar Empresa"
                        data-ToolTip="default" ToolTip="Consolidar o cnpj da empresa." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresaCliente" runat="server" Width="600px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" runat="server" Text="Consolidar Cliente"
                        data-ToolTip="default" ToolTip="Consolidar o cpf/cnpj do cliente." />
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="565px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultarCliente" OnClick="btnConsultarCliente_Click" runat="server"
                        CssClass="btn" UseSubmitBehavior="False" Text=">" data-ToolTip="default" ToolTip="Unificar as informações por cliente." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlClasseOperacao" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe Sub-Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlClasseSuboperacao" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtDataInicial" runat="server" Width="100px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Informar o período inicial de consulta." />
                </div>
                <div class="collbl">
                    Data final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtDataFinal" runat="server" Width="100px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Informar o período final de consulta." />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 100%;">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
