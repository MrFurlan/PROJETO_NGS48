<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeLaudosSemNotaFiscal.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeLaudosSemNotaFiscal" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 135px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngLaudoSemNotaFiscal" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlLaudoSemNotaFiscal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório de Laudos sem Nota Fiscal
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
                    Unidade
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlUnidadeDeNegocio" runat="server" Width="606px" OnSelectedIndexChanged="DdlUnidadeDeNegocio_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="606px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente."
                        runat="server" Text="Consolidar Cliente:" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="569px" data-ToolTip="default"
                        ToolTip="Unificar as informações por cliente." />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Unificar as informações por cliente." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Deposito:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDeposito" runat="server" Width="569px" data-ToolTip="default"
                        ToolTip="Local de armazenamento da mercadoria." />
                    <asp:HiddenField ID="txtCodigoDeposito" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnDeposito" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Local de armazenamento da mercadoria." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoInicial" CssClass="calendario" runat="server" Width="116px"
                        data-ToolTip="default" ToolTip="Informar o data inicial de consulta." />
                </div>
                <div class="collbl">
                    Período final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoFinal" CssClass="calendario" runat="server" Width="116px"
                        CausesValidation="True" data-ToolTip="default" ToolTip="Informar o data final de consulta." />
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
