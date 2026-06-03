<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioAnaliticoDeFretes.aspx.vb" Inherits="NGS.Web.UI.RelatorioAnaliticoDeFretes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadRelatorioAnaliticoDeFretes() {
            $("#MainContent_lstClasseOp").multiselect().multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadRelatorioAnaliticoDeFretes();
            var prmRelatorioAnaliticoDeFretes = Sys.WebForms.PageRequestManager.getInstance();
            prmRelatorioAnaliticoDeFretes.add_endRequest(pageLoadRelatorioAnaliticoDeFretes);
        });
    </script>
    <style type="text/css">
        .collbl {
            width: 170px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngRelatorioAnaliticoDeFretes" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioAnaliticoDeFretes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório Analítico de Fretes
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf" id="divPdf" runat="server">
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
                    Relatório:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoPorPedido" runat="server" AutoPostBack="true" Checked="true"
                        GroupName="TipoRelatorio" Text="Por Pedido" data-ToolTip="default" ToolTip="Selecionar se o relatório será por pedido, transportador ou pagamento" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoPorTransportador" runat="server" AutoPostBack="true" GroupName="TipoRelatorio"
                        Text="Por Transportador" data-ToolTip="default" ToolTip="Selecionar se o relatório será por pedido, transportador ou pagamento" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbPlanilhaDePagamento" runat="server" AutoPostBack="True" GroupName="TipoRelatorio"
                        Text="Planilha de Pagamento" data-ToolTip="default" ToolTip="Selecionar se o relatório será por pedido, transportador ou pagamento" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="598px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Transportador:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoTransportador" runat="server" />
                    <asp:TextBox ID="txtTransportador" runat="server" Width="558px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnTransportador" runat="server" OnClick="BtnTransportador_Click"
                        CssClass="btn" Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Nome da pessoa/empresa responsável pelo transporte da mercadoria." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="558px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaCliente" runat="server" Text=">" CssClass="btn" UseSubmitBehavior="False"
                        data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Width="90px" Enabled="false" CssClass="txtNumerico"
                        Style="text-align: right" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnPedido" runat="server" Text=">" UseSubmitBehavior="False" CssClass="btn"
                        data-ToolTip="default" ToolTip="Informar o número do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe Operacao:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstClasseOp" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="370px" data-ToolTip="default" ToolTip="Selecionar a classe de opeação desejada." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdComClasse" Text="Com Estas Classes" GroupName="cl" Checked="true"
                        runat="server" data-ToolTip="default" ToolTip="Selecionar a classe de opeação desejada." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdSemClasse" Text="Sem Estas Classes" GroupName="cl" runat="server"
                        data-ToolTip="default" ToolTip="Selecionar a classe de opeação desejada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Frete Pela Empresa No(a):
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdFreteNaNota" runat="server" Text="Nota" GroupName="Frete"
                        Checked="true" data-ToolTip="default" ToolTip=" Serão mostrados todos os pedidos que tenham pelo menos uma nota cujo frete seja por conta da empresa" />
                    <asp:RadioButton ID="rdFreteNoPedido" runat="server" Text="Pedido" GroupName="Frete"
                        data-ToolTip="default" ToolTip="Serão mostrados apenas os pedidos cujo frete seja por conta da empresa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="300px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkPeriodo" runat="server" AutoPostBack="true" />
                    Período:
                </div>

                <div class="coltxt" id="divTipoDePeriodo" runat="server">
                    <div class="coltxt">
                        <asp:TextBox ID="txtData1" CssClass="calendario" runat="server" Width="100px" data-ToolTip="default"
                            ToolTip="Informar a data inicial e fnal da consulta." />
                        &nbsp;&nbsp;á&nbsp;
                    <asp:TextBox ID="txtData2" CssClass="calendario" runat="server" Width="100px" data-ToolTip="default"
                        ToolTip="Informar a data inicial e fnal da consulta." />
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rdMovimentoNota" runat="server" Text="Movimento NF" GroupName="Per"
                            Checked="true" data-ToolTip="default" ToolTip="" />
                        <asp:RadioButton ID="rdMovimentoPedido" runat="server" Text="Data do Pedido" GroupName="Per"
                            data-ToolTip="default" ToolTip="" />
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Apenas com peso de chegada?
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoSim" runat="server" GroupName="grpPesoDeChegada" Text="Sim"
                        data-ToolTip="default" ToolTip="Informar se é para ser considerado ou não o peso de chegada." />
                    <asp:RadioButton ID="rdoNao" runat="server" GroupName="grpPesoDeChegada" Text="Não"
                        Checked="true" data-ToolTip="default" ToolTip="Informar se é para ser considerado ou não o peso de chegada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Apenas com financeiro?
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoFinanceiroSim" runat="server" GroupName="grpFinanceiro" Text="Sim"
                        AutoPostBack="true" data-ToolTip="default" ToolTip="Informar se é para ser considerado ou não apenas o financeiro." />
                    <asp:RadioButton ID="rdoFinanceiroNao" runat="server" GroupName="grpFinanceiro" Text="Não"
                        Checked="true" AutoPostBack="true" data-ToolTip="default" ToolTip="Informar se é para ser considerado ou não apenas o financeiro." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Financeiro em:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkProvisao" runat="server" Text="Provisão (3)" data-ToolTip="default"
                        ToolTip="Selecionar se a consulta será por provisão, previsão ou baixado." />
                    <asp:CheckBox ID="chkPrevisao" runat="server" Text="Previsão (2)" data-ToolTip="default"
                        ToolTip="Selecionar se a consulta será por provisão, previsão ou baixado." />
                    <asp:CheckBox ID="chkBaixa" runat="server" Text="Baixado (1)" data-ToolTip="default"
                        ToolTip="Selecionar se a consulta será por provisão, previsão ou baixado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    NF:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNota" runat="server" Width="100px" CssClass="txtNumerico" data-ToolTip="default"
                        ToolTip="Inserir o número da Nota Fiscal." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Ctrc:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDACTE" runat="server" Width="100px" CssClass="txtNumerico" data-ToolTip="default"
                        ToolTip="Inserir o número do CTRC." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
