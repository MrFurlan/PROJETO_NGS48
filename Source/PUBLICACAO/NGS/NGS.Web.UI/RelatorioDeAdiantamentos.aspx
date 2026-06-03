<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="RelatorioDeAdiantamentos.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeAdiantamentos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadAdiantamentos() {
            $("#<%=lstConta.ClientID%>").multiselect({
                header: "Escolha as Contas Contabeis de Adiantamento.",
                selectedList: 2
            }).multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadAdiantamentos();
            var prmAdiantamentos = Sys.WebForms.PageRequestManager.getInstance();
            prmAdiantamentos.add_endRequest(pageLoadAdiantamentos);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPlanoDeContas" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPlanoDeContas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório de Adiantamento(s)
                <img src="Images/info16x16-2.png" alt="I" style="margin-left: 4px; margin-bottom: 6px;" data-tooltip="default" title="Relatório de Adiantamento(s)." />
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton class="iconPdf" ID="lnkPdf" runat="server" Text="Pdf" data-tooltip="default" ToolTip="Emite Relatório em ''.pdf''" />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
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
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="590px" data-tooltip="default" ToolTip="Unidade de negócio empresarial." AutoPostBack="true" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsEmpresa" runat="server" Text="Cons. Empresa:" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="590px" data-tooltip="default" ToolTip="Empresa principal, onde foi realizado o adiantamento." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsCliente" runat="server" Text="Cons. Cliente:" />
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" Enabled="false" runat="server" Width="551px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnConsultaCliente" runat="server" CssClass="btn" Text=">" data-tooltip="default" ToolTip="Busca o cliente ou fornecedor no qual se destina o adiantamento." />
                </div>
            </div>


            <div class="row">
                <div class="collbl">
                    Data Base:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDtBase" runat="server" CssClass="calendario" data-tooltip="default" Width="273px" ToolTip="Data Base do Relatorio, adiantamentos e baixas após essa data serão ignorados." />
                </div>
            </div>

            <div class="row">
                <div class="collbl">
                    Pedido(s):
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" data-tooltip="default" Width="273px" ToolTip="Para consultar mais de um pedido, separe-os com virgula(,). Ex: 1,2,3." />
                </div>
            </div>

            <div class="row">
                <div class="collbl">
                    Exibição:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkEmpresa" runat="server" Text="Por Empresa." />
                </div>
            </div>


            <div class="row">
                <div class="collbl">
                    Filtro:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbAnalitico" runat="server" Text="Analítico" Checked="true" GroupName="Filtro" data-tooltip="default" ToolTip="Filtro utilizado para emissão em '.pdf', onde será possivel visualizar de forma analítica, as baixas do adiantamento." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbSintetico" runat="server" Text="Sintético" GroupName="Filtro" data-tooltip="default" ToolTip="Filtro utilizado para emissão em '.pdf', onde será possivel visualizar, os valores totais do adiantamento." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkApenasComSaldo" runat="server" Text="Apenas com Saldo a baixar" Checked="true" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo Adto:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbCliente" runat="server" Text="Cliente" Checked="true" GroupName="TipoAdto" data-tooltip="default" ToolTip="Encontrar as contas de adiantamento de Clientes." />
                </div>
                <div class="coltxt" style="margin-left: 7px;">
                    <asp:RadioButton ID="rbFornecedor" runat="server" Text="Fornecedor" GroupName="TipoAdto" data-tooltip="default" ToolTip="Encontrar as contas de adiantamento de Fornecedores." />
                </div>
            </div>

            <div class="row">
                <div class="collbl">
                    Conta Contabil:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstConta" runat="server" data-tooltip="default" ToolTip="Selecione as Contas Contabeis de Adiantamento" CssClass="multiselect" />
                </div>
            </div>


        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
