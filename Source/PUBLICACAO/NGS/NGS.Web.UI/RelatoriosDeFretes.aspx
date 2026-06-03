<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatoriosDeFretes.aspx.vb" Inherits="NGS.Web.UI.RelatoriosDeFretes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngDefault" runat="server" EnablePartialRendering="true"
        EnableScriptGlobalization="True" EnableScriptLocalization="True">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlDefault" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatórios De Fretes
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEmpresa" runat="server" Width="601px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtDeposito" runat="server" Enabled="False" Width="562px" />
                    <asp:HiddenField ID="TxtCodigoDeposito" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnDeposito" runat="server" OnClick="BtnDeposito_Click" CssClass="btn"
                        Text=">" UseSubmitBehavior="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtClientes" runat="server" Enabled="False" Width="562px" />
                </div>
                <div class="coltxt">
                    <asp:Button CssClass="btn" ID="cmdConsultaCliente" runat="server" OnClick="cmdConsultaCliente_Click"
                        Text=">" UseSubmitBehavior="False" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Transportador:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTransportador" runat="server" Width="562px" Enabled="False" /><asp:HiddenField
                        ID="txtCodigoTransportador" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnTransportador" runat="server" OnClick="BtnTransportador_Click"
                        Text=">" UseSubmitBehavior="False" CssClass="btn" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupoProduto" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlGrupoProduto_SelectedIndexChanged"
                        Width="601px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProduto" runat="server" Width="601px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="112px" />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="112px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CkFretesAPagar" runat="server" AutoPostBack="True" OnCheckedChanged="CkFretesAPagar_CheckedChanged"
                        Text="Fretes a Pagar" ValidationGroup="Tipo" />
                    <asp:CheckBox ID="CkPosicaoDeFretes" runat="server" AutoPostBack="True" OnCheckedChanged="CkPosicaoDeFretes_CheckedChanged"
                        Text="Posição de Fretes" ValidationGroup="Tipo" />
                    <asp:CheckBox ID="CkDolar" runat="server" Text="Dolar" Enabled="False" />
                    <asp:CheckBox ID="CkSaldoApagar" runat="server" Text="Com Saldo a Pagar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
