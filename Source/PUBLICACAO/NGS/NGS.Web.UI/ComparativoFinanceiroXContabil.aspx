<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ComparativoFinanceiroXContabil.aspx.vb" Inherits="NGS.Web.UI.ComparativoFinanceiroXContabil" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl
        {
            width: 140px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="UpdatePanelPrincipal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Comparativo Financeiro X Contábil
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Formato Saida:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="radPDF" runat="server" Checked="True" GroupName="RELATORIO"
                        Text="PDF" />
                    <asp:RadioButton ID="radHTML" runat="server" GroupName="RELATORIO" Text="HTML" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEmpresa" runat="server" Width="550px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnEmpresa" OnClick="btnEmpresa_Click" runat="server" Text=" > "
                        CssClass="btn" />
                    <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." runat="server" />
                    Consolidar Empresa:
                </div>
                <div class="collbl" style="margin-left: 130px;">
                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente." runat="server" />
                    Consolidar Cliente:
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="550px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" OnClick="btnCliente_Click" runat="server" Text=" > "
                        CssClass="btn" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Limite:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" CssClass="calendario" runat="server"
                        Width="96px" data-ToolTip="default" ToolTip="Periodo Inicial" />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="ChkIsola" runat="server" />
                    Isolar zerados:
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="line-height: 18px; margin-left: 155px; margin-right: 10px;">
                    <asp:RadioButton ID="RbReceber" runat="server" AutoPostBack="True" Font-Bold="True"
                        GroupName="RP" OnCheckedChanged="RbReceber_CheckedChanged" Text="Receber" Width="77px" />
                    <br />
                    <asp:CheckBox ID="Chk_1010301" runat="server" Text="1010301 - Clientes no Pais" />
                    <br />
                    <asp:CheckBox ID="Chk_1010302" runat="server" Text="1010302 - Clientes no Exterior" />
                    <br />
                    <asp:CheckBox ID="Chk_1010201" runat="server" Text="1010201 - Clientes no Pais" />
                    <br />
                    <asp:CheckBox ID="Chk_1010202" runat="server" Text="1010202 - Clientes no Exterior" />
                </div>
                <div class="coltxt" style="line-height: 18px;">
                    <asp:RadioButton ID="RbPagar" runat="server" Text="Pagar" AutoPostBack="True" GroupName="RP"
                        Font-Bold="True" OnCheckedChanged="RbPagar_CheckedChanged" />
                    <br />
                    <asp:CheckBox ID="Chk_2010101" runat="server" Text="2010101 - Fornecedores de Insumos e Produtos Rurais" />
                    <br />
                    <asp:CheckBox ID="Chk_2010102" runat="server" Text="2010102 - Fornecedores a Fixar Produtos Rurais" />
                    <br />
                    <asp:CheckBox ID="Chk_2010103" runat="server" Text="2010103 - Fornecedores de Fretes" />
                    <br />
                    <asp:CheckBox ID="Chk_2010104" runat="server" Text="2010104 - Fornecedores Comissões" />
                    <br />
                    <asp:CheckBox ID="Chk_2010105" runat="server" Text="2010105 - Fornecedores de Serviços" />
                    <br />
                    <asp:CheckBox ID="Chk_2010108" runat="server" Text="2010108 - Fornecedores Convenios " />
                    <br />
                    <asp:CheckBox ID="Chk_2010109" runat="server" Text="2010109 - Fornecedores Diversos" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
</asp:Content>
