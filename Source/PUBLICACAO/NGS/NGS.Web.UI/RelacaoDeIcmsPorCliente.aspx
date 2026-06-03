<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelacaoDeIcmsPorCliente.aspx.vb" Inherits="NGS.Web.UI.RelacaoDeIcmsPorCliente" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmRelacaoDeIcmsPorCliente" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelacaoDeIcmsPorCliente" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Relação de Icms Por Cliente
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
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
                    <asp:TextBox ID="txtDataInicial" runat="server" Width="96px" CssClass="calendario" data-ToolTip="default" ToolTip="Informar o período de apuração." />
                    &nbsp;à&nbsp;
                    <asp:TextBox ID="txtDataFinal" runat="server" Width="96px" CssClass="calendario" data-ToolTip="default" ToolTip="Informar o período de apuração." />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidado" runat="server" Text="Consolidado" data-ToolTip="default" ToolTip="Selecionar para consolidar as informações." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
