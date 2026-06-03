<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="NovoRelatorioDeTitulos.aspx.vb" Inherits="NGS.Web.UI.NovoRelatorioDeTitulos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            min-width: 133px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioDeTitulosAReceber" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioDeTitulos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdfCodigoCliente" runat="server" ClientIDMode="Static" />
            <div class="titulodiv">
                Relatório de Títulos
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
                    <asp:DropDownList ID="DdlUnidadeConsultaTitulos" runat="server" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlUnidadeConsultaTitulos_SelectedIndexChanged" Width="618px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." runat="server" Text="Consolidar Empresa" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresaConsultaTitulos" runat="server" Width="618px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente." runat="server" Text="Consolidar Cliente" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeCliente" runat="server" Width="581px" Enabled="False" ClientIDMode="Static" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaCliente" OnClick="cmdBuscaCliente_Click" runat="server" UseSubmitBehavior="False"
                        CssClass="btn" Text=">" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Títulos:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RdReceber" runat="server" Checked="True" GroupName="titulo"
                        Text="Receber" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RdPagar" runat="server" GroupName="titulo" Text="Pagar" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdContabil" runat="server" GroupName="titulo" Text="Contábil" />
                </div>
                <div class="collbl" style="margin-left: 32px;">
                    Agrupamento Apurar:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdMestre" runat="server" GroupName="mestre" Text="Mestre" Checked="True" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdFilho" runat="server" GroupName="mestre" Text="Filho" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkBaixados" runat="server" Checked="True" Text="Baixados" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkProvisao" runat="server" Checked="True" Text="Provisao" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkPrevisao" runat="server" Checked="True" Text="Previsão" />
                </div>
                <div class="collbl">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMoeda" runat="server" Width="125px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Relatório:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="radData" runat="server" AutoPostBack="True" Checked="True" GroupName="Quebra"
                        OnCheckedChanged="radData_CheckedChanged" Text="Por Data" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="radClienteData" runat="server" AutoPostBack="True" GroupName="Quebra"
                        OnCheckedChanged="radClienteData_CheckedChanged" Text="Por Cliente / Data" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkListarProdutos" runat="server" Font-Bold="True" Text="Listar Produtos do Pedido"
                        Visible="False" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RdClientePedido" runat="server" AutoPostBack="True" GroupName="Quebra"
                        OnCheckedChanged="RdClientePedido_CheckedChanged" Text="Por Cliente / Pedido" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Representante:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlRepresentante" runat="server" Width="617px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Carteira:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCarteiraDoTitulo" runat="server" Width="617px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" OnSelectedIndexChanged="ddlSafra_SelectedIndexChanged"
                        Width="618px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkUsarPeriodo" runat="server" AutoPostBack="True" Text="Usar Periodo"
                        Checked="True" />
                </div>
                <div class="coltxt">
                    <asp:Panel ID="pnlData" runat="server" Width="400px">
                        <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" runat="server" data-ToolTip="default" ToolTip="Periodo Inicial"
                            Width="116px" CssClass="calendario" />
                        &nbsp;a
                        <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" runat="server" CausesValidation="True"
                            Width="116px" CssClass="calendario" />
                    </asp:Panel>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
