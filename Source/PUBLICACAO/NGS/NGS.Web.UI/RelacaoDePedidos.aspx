<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelacaoDePedidos.aspx.vb" Inherits="NGS.Web.UI.RelacaoDePedidos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadRelatorioRelacaoDePedidos() {
            $("#<%=lstClasses.ClientID%>").multiselect({
                header: "Escolha as Classes!",
                selectedList: 20
            }).multiselectfilter();

            $("#<%=lstFinalidade.ClientID%>").multiselect().multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadRelatorioRelacaoDePedidos();
            var prmRelatorioRelacaoDePedidos = Sys.WebForms.PageRequestManager.getInstance();
            prmRelatorioRelacaoDePedidos.add_endRequest(pageLoadRelatorioRelacaoDePedidos);
        });
    </script>
    <style type="text/css">
        .collbl {
            width: 135px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmRelacaoDePedidos" runat="server" AsyncPostBackTimeout="10000"
        EnableScriptGlobalization="True" EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelacaoDePedidos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relação De Pedidos
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
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcelDados" runat="server" Text="Excel Dados" />
                                </li>
                            </ul>
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
                    Tipo de Relatório:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdPorCliente" runat="server" Checked="True" GroupName="rel"
                        Text="Por Cliente" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:RadioButton ID="rdPorPedido" runat="server" GroupName="rel" Text="Por Pedido" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:RadioButton ID="rdPorProduto" runat="server" GroupName="rel" Text="Por Produto" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlUnidadeConsultaTitulos" runat="server" AutoPostBack="True" CssClass="multiselect"
                        OnSelectedIndexChanged="DdlUnidadeConsultaTitulos_SelectedIndexChanged" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." CssClass="multiselect" Text="Consolidar Empresa:" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresaConsultaTitulos" runat="server" TabIndex="2" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente." runat="server" Text="Consolidar Clientes:" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClientes" runat="server" Enabled="false" Width="585px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaCliente" runat="server" OnClick="cmdConsultaCliente_Click"
                        Text=" &gt; " UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Unificar as informações por cliente." />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Representante:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlRepresentante" runat="server" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstClasses" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="624px" data-ToolTip="default" ToolTip="Selecionar a classe desejada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbOperacao" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbOperacao_SelectedIndexChanged"
                        Width="315px" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbSubOperacao" runat="server" Width="305px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbEntradas" runat="server" GroupName="Tipo" Text="Entradas" data-ToolTip="default" ToolTip="Selecionar se o pedido é uma entrada ou uma saída." />
                    &nbsp; &nbsp;<asp:RadioButton ID="rbSaidas" runat="server" GroupName="Tipo" Text="Saídas" data-ToolTip="default" ToolTip="Selecionar se o pedido é uma entrada ou uma saída." />
                    &nbsp; &nbsp;<asp:RadioButton ID="rbTodas" runat="server" Checked="True" GroupName="Tipo"
                        Text="Todos" data-ToolTip="default" ToolTip="Selecionar se o pedido é uma entrada ou uma saída." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Finalidade:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstFinalidade" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="370px" data-ToolTip="default" ToolTip="Objetivo da criação do pedido." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdComFinalidade" Text="Com Estas Finalidades" GroupName="fn"
                        Checked="true" runat="server" data-ToolTip="default" ToolTip="Objetivo da criação do pedido." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdSemFinalidade" Text="Sem Estas Finalidades" GroupName="fn"
                        runat="server" data-ToolTip="default" ToolTip="Objetivo da criação do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Marca:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMarca" runat="server" AutoPostBack="True" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 100%; margin-bottom: 4px;">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbSafra" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbSafra_SelectedIndexChanged"
                        Width="300px" />
                </div>
                <div class="collbl">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbMoeda" runat="server" AutoPostBack="True" Width="104px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Granéis:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbKg" runat="server" Checked="True" GroupName="Graneis" Text="Unitário por Kg" data-ToolTip="default" ToolTip="Selecionar se é por saca ou kg." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbSaca" runat="server" GroupName="Graneis" Text="Unitário por Saca" data-ToolTip="default" ToolTip="Selecionar se é por saca ou kg." />
                </div>
                <div class="collbl" style="margin-left: 83px;">
                    Pedidos por Troca:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdTrocaSim" runat="server" GroupName="troca" Text="Sim" data-ToolTip="default" ToolTip="Marcar se o pedido refere-se ou não a uma troca." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdTrocaNao" runat="server" GroupName="troca" Text="Não" data-ToolTip="default" ToolTip="Marcar se o pedido refere-se ou não a uma troca." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdTrocaTodos" runat="server" Checked="True" GroupName="troca"
                        Text="Todos" data-ToolTip="default" ToolTip="Marcar se o pedido refere-se ou não a uma troca." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="ckConsiderarPeriodo" runat="server" Checked="True" Enabled="False"
                        Text="Considerar Período:" AutoPostBack="True" data-ToolTip="default" ToolTip="Informar o período a ser considerado." />
                </div>
                <div class="coltxt" runat="server">
                    <asp:TextBox ID="txtPeriodoInicial" runat="server" CssClass="calendario" TabIndex="5"
                        data-ToolTip="default" ToolTip="Periodo Inicial" Width="100px" />&nbsp;á&nbsp;<asp:TextBox ID="txtPeriodoFinal"
                            runat="server" CausesValidation="True" CssClass="calendario" TabIndex="6" Width="100px" data-ToolTip="default" ToolTip="Informar o período a ser considerado." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkOmitirValores" runat="server" Text="Omitir Valores" data-ToolTip="default" ToolTip="Marcar a opção caso deseje omitir valores." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo do Relatório:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="radPeso" runat="server" Checked="True" GroupName="pesovalor"
                        Text="Peso" ValidationGroup="pesovalor" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:RadioButton ID="radValor" runat="server" GroupName="pesovalor" Text="Valor"
                        ValidationGroup="pesovalor" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Saldo:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkSaldoQuantidade" runat="server" Text="Com Saldo de Quantidade" data-ToolTip="default" ToolTip="Selecionar o tipo de saldo desejado a partir das opções." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkSaldoValor" runat="server" Text="Com Saldo de Valor" data-ToolTip="default" ToolTip="Selecionar o tipo de saldo desejado a partir das opções." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkFaturadoNaoEntregue" runat="server" Text="Com Saldo Faturado Nao Entregue" data-ToolTip="default" ToolTip="Selecionar o tipo de saldo desejado a partir das opções." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtro(s):
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkOmitirResumoProduto" runat="server" Text="Omitir Resumo Produto" Checked="true" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
