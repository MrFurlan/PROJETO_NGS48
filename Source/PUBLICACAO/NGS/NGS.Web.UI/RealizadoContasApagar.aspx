<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RealizadoContasApagar.aspx.vb" Inherits="NGS.Web.UI.RealizadoContasApagar" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRealizadoContasPagar" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRealizadoContasPagar" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Realizado Contas à Pagar
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
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
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="604px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlUnidadeDeNegocio_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="604px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="hdfConta" runat="server" />
                    <asp:TextBox ID="txtConta" runat="server" Width="565px" ReadOnly="true" data-ToolTip="default"
                        ToolTip="Selecionar a conta desejada." />
                    <asp:Button ID="btnConta" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Selecionar a conta desejada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente."
                        runat="server" Text="Cliente" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Enabled="False" Width="565px" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" CssClass="btn" runat="server" Text=">" OnClick="btnCliente_Click"
                        data-ToolTip="default" ToolTip="Selecionar o fornecedor desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Carteira:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCarteira" runat="server" Width="604px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoInicial" CssClass="calendario" runat="server" Width="100px"
                        data-ToolTip="default" ToolTip="Data inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoFinal" CssClass="calendario" runat="server" Width="100px"
                        CausesValidation="True" data-ToolTip="default" ToolTip="Data final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="radMovimento" runat="server" Text="Movimento" GroupName="radPeriodo" Checked="true" />
                    <asp:RadioButton ID="radVencimento" runat="server" Text="Vencimento" GroupName="radPeriodo" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
