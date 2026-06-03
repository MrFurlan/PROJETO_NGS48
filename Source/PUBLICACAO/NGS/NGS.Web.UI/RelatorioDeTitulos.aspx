<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeTitulos.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeTitulos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadRelatorioDeTitulos() {
            $("#MainContent_lstFinalidade").multiselect().multiselectfilter();
            $("#MainContent_lstTipoPagRec").multiselect().multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadRelatorioDeTitulos();
            var prmRelatorioDeTitulos = Sys.WebForms.PageRequestManager.getInstance();
            prmRelatorioDeTitulos.add_endRequest(pageLoadRelatorioDeTitulos);
        });
    </script>
    <style type="text/css">
        .collbl {
            width: 120px;
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
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkPlanilhaExcel" Text="Planilha Excel" runat="server" />
                                </li>
                            </ul>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjudar" Text="Ajudar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlUnidadeConsultaTitulos" runat="server" Width="633px" OnSelectedIndexChanged="DdlUnidadeConsultaTitulos_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" runat="server" Text="Empresa:"
                        data-ToolTip="default" ToolTip="Consolidar o cnpj da empresa." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresaConsultaTitulos" runat="server" Width="633px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" runat="server" Text="Cliente:"
                        data-ToolTip="default" ToolTip="Consolidar o cpf/cnpj do cliente." />
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtClientes" runat="server" Width="593px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaCliente" OnClick="cmdBuscaCliente_Click" runat="server" UseSubmitBehavior="False"
                        CssClass="btn" Text=">" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnPedido" runat="server" OnClick="btnPedido_Click" Text=">" UseSubmitBehavior="False"
                        CssClass="btn" data-ToolTip="default" ToolTip="Número do pedido." />
                    <asp:HiddenField ID="txtCodigoPedido" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Representante:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlRepresentante" runat="server" Width="633px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Carteira:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCarteiraDoTitulo" runat="server" Width="633px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Finalidade Financeira:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlFinalidadeFinanceira" runat="server" Width="633px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Bancos:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlBanco" runat="server" Width="633px" />
                </div>
            </div>

            <div class="row">
                <div class="collbl">
                    Tipo Pgto/Rec:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstTipoPagRec" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="370px" data-ToolTip="default" ToolTip="Tipo de Pagamento do Titulo." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdComTpPagRec" Text="Com Estes Tipos" GroupName="tpg"
                        Checked="true" runat="server" data-ToolTip="default" ToolTip="Tipo de Pagamento do Titulo." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdSemTpPagRec" Text="Sem Estes Tipos" GroupName="tpg"
                        runat="server" data-ToolTip="default" ToolTip="Tipo de Pagamento do Titulo." />
                </div>
            </div>

            <div class="row">
                <div class="collbl">
                    Finalidade Pedido:
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
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" OnSelectedIndexChanged="ddlSafra_SelectedIndexChanged"
                        Width="633px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkPeriodo" runat="server" Text="Período:" Checked="True" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" runat="server" Width="116px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Data inicial à data final da consulta." />
                    &nbsp;&nbsp;à&nbsp;
                    <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" runat="server" Width="116px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Data inicial à data final da consulta." />
                    <asp:RadioButton ID="radVencimento" runat="server" AutoPostBack="True" Text="Vencimento" GroupName="radPeriodo" Checked="true" />
                    <asp:RadioButton ID="radMovimento" runat="server" AutoPostBack="True" Text="Movimento" GroupName="radPeriodo" />
                    <asp:RadioButton ID="radDataBase" runat="server" AutoPostBack="True" Text="Data Base" GroupName="radPeriodo" OnCheckedChanged="radDataBase_CheckedChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Títulos:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RdReceber" runat="server" AutoPostBack="True" Text="A Receber"
                        OnCheckedChanged="RdReceber_CheckedChanged" Checked="True" GroupName="titulo"
                        data-ToolTip="default" ToolTip="Selecionar o tipo de títulos a serem consultados." />
                    <asp:RadioButton ID="RdPagar" runat="server" AutoPostBack="True" Text="A Pagar" OnCheckedChanged="RdPagar_CheckedChanged"
                        GroupName="titulo" data-ToolTip="default" ToolTip="Selecionar o tipo de títulos a serem consultados." />
                    <asp:RadioButton ID="RdReceberPagar" runat="server" GroupName="titulo" Text="Todos"
                        data-ToolTip="default" ToolTip="Selecionar o tipo de títulos a serem consultados." />

                    <asp:CheckBox ID="chkProjecao" runat="server" AutoPostBack="true" Text="Projetar Titulos de Contratos Futuros" />

                    <asp:RadioButton ID="rdFinalidadeFinanceira" runat="server" Checked="True" GroupName="carteira"
                        Text="Finalidade Financeira" data-ToolTip="default" ToolTip="Selecionar o tipo de títulos a serem consultados." />

                    <asp:RadioButton ID="rdCarteira" runat="server" GroupName="carteira" Text="Carteira"
                        data-ToolTip="default" ToolTip="Selecionar o tipo de títulos a serem consultados." />
                </div>

            </div>
            <div class="row">
                <div class="collbl">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMoeda" runat="server" Width="125px" />
                </div>
            </div>




            <div class="row">
                <div class="collbl">
                    Quebra:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="radData" runat="server" Text="Por Data" Checked="True" GroupName="Quebra"
                        AutoPostBack="True" OnCheckedChanged="radData_CheckedChanged" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de quebra do relatório." />
                    <asp:RadioButton ID="radClienteData" runat="server" Text="Por Cliente / Data" GroupName="Quebra"
                        AutoPostBack="True" OnCheckedChanged="radClienteData_CheckedChanged" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de quebra do relatório." />
                    <asp:RadioButton ID="RdClientePedido" runat="server" AutoPostBack="True" GroupName="Quebra"
                        OnCheckedChanged="RdClientePedido_CheckedChanged" Text="Por Cliente / Pedido"
                        data-ToolTip="default" ToolTip="Selecionar o tipo de quebra do relatório." />
                    <asp:CheckBox ID="chkListarProdutos" runat="server" Font-Bold="True" Text="Listar Produtos do Pedido"
                        Visible="False" data-ToolTip="default" ToolTip="Selecionar o tipo de quebra do relatório." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadTodos" runat="server" AutoPostBack="True" Text="Todos" OnCheckedChanged="RadTodos_CheckedChanged"
                        GroupName="Selecao" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:RadioButton ID="RadBaixados" runat="server" AutoPostBack="True" Text="Baixados"
                        OnCheckedChanged="RadBaixados_CheckedChanged" GroupName="Selecao" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:RadioButton ID="RadAtivos" runat="server" AutoPostBack="True" Text="Ativos   "
                        OnCheckedChanged="RadAtivos_CheckedChanged" GroupName="Selecao" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:CheckBox ID="chkPrevisao" runat="server" Text="Previsão" Checked="True" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:CheckBox ID="chkProvisao" runat="server" Text="Provisão" Checked="True" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:CheckBox ID="chkMovtoBancario" runat="server" AutoPostBack="true" Text="Movimento Bancário" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:CheckBox ID="ChkApenasNF" runat="server" Text="Apenas com Nota Fiscal" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Valor:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbLiquido" Text="Líquido" GroupName="valor" data-ToolTip="default"
                        ToolTip="Valor Líquido do Título. Parâmetro usado apenas para Títulos Baixados em PDF"
                        Checked="true" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbDocumento" Text="Documento" GroupName="valor" data-ToolTip="default"
                        ToolTip="Valor do Documento do Título. Parâmetro usado apenas para Títulos Baixados em PDF"
                        runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Parametros:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkObservacao" runat="server" Text="Emitir Historico com Observações"
                        data-ToolTip="default" ToolTip="Selecionar para gerar relatório com observação." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
