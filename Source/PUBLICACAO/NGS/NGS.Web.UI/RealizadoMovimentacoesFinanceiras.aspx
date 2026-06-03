<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RealizadoMovimentacoesFinanceiras.aspx.vb" Inherits="NGS.Web.UI.RealizadoMovimentacoesFinanceiras" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <div class="titulodiv">
        Realizado Movimentações Financeiras
    </div>
    <asp:UpdatePanel ID="UpdatePanelPrincipal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="602px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="602px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" Enabled="false" runat="server" Width="562px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" CssClass="btn" OnClick="btnCliente_Click" runat="server"
                        Text=">" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Periodo Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" CssClass="calendario" runat="server"
                        Width="90px" data-ToolTip="default" ToolTip="Data inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Periodo Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" CssClass="calendario" runat="server"
                        Width="90px" CausesValidation="True" data-ToolTip="default" ToolTip="Data final da pesquisa." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
