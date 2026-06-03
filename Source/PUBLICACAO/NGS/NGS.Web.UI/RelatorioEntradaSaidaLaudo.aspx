<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioEntradaSaidaLaudo.aspx.vb" Inherits="NGS.Web.UI.RelatorioEntradaSaidaLaudo" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .wit {
            width: 160px;
        }
    </style>

    <script type="text/javascript">
        function pageLoadRelatorioEntradaSaidaLaudo() {
            $("#MainContent_lstClasseOp").multiselect().multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadRelatorioEntradaSaidaLaudo();
            var prmRelatorioEntradaSaidaLaudo = Sys.WebForms.PageRequestManager.getInstance();
            prmRelatorioEntradaSaidaLaudo.add_endRequest(pageLoadRelatorioEntradaSaidaLaudo);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioEntradaSaidaLaudo" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioEntradaSaidaLaudo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Posição de Cargas e Descargas
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
                    <asp:DropDownList ID="cmbUnidadeNegocio" runat="server" Width="624px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbUnidadeNegocio_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEmpresa" runat="server" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbDeposito" runat="server" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkClientes" runat="server" Text="Cliente:" />
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtClientes" runat="server" Width="585px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaCliente" OnClick="cmdConsultaCliente_Click" CssClass="btn"
                        runat="server" Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtrar por:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ChkDepositante" runat="server" CssClass="rotulo" Text="Depositante"
                        AutoPostBack="True" OnCheckedChanged="ChkDepositante_CheckedChanged" data-ToolTip="default"
                        ToolTip="Selecionar para gerar relatório por depositante." />
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
                    <asp:Button ID="cmdBuscaPedido" runat="server" CssClass="btn" CausesValidation="False"
                        OnClick="cmdBuscaPedido_Click" Text=">" UseSubmitBehavior="False" data-ToolTip="default"
                        ToolTip="Informar o número do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depositante:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDepositante" runat="server" Width="585px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoDepositante" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaDepositante" CssClass="btn" OnClick="cmdConsultaDepositante_Click"
                        runat="server" Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o depositante desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbOperacao" runat="server" Width="304px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbOperacao_SelectedIndexChanged" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbSubOperacao" runat="server" Width="315px" />
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
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlGrupoDeProdutos" runat="server" Width="624px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlGrupoDeProdutos_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbProdutos" runat="server" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt wit">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="70px"
                        data-ToolTip="default" ToolTip="Informar o período inicial de consulta." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="70px"
                        data-ToolTip="default" ToolTip="Informar o período final de consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Entrada/Saída:
                </div>
                <div class="coltxt wit">
                    <asp:RadioButton ID="optEntrada" runat="server" Text="Entrada" GroupName="EntSai"
                        Checked="True" data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                    <asp:RadioButton ID="optSaida" runat="server" Text="Saida" GroupName="EntSai" data-ToolTip="default"
                        ToolTip="Informar se a nota é de entrada ou saída." />
                </div>
                <div class="collbl">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbSituacao" runat="server" Width="150px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Média:
                </div>
                <div class="coltxt wit">
                    <asp:RadioButton ID="optAritmetica" runat="server" Text="Aritmetica" GroupName="Media"
                        data-ToolTip="default" ToolTip="Informar o tipo de média." />
                    <asp:RadioButton ID="optPonderada" runat="server" Text="Ponderada" GroupName="Media"
                        Checked="True" data-ToolTip="default" ToolTip="Informar o tipo de média." />
                </div>
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="optAberto" runat="server" Text="Aberto(s)" GroupName="Laudos"
                        data-ToolTip="default" ToolTip="Informar se é aberto ou fechado." />
                    <asp:RadioButton ID="optFechado" runat="server" Text="Fechado(s)" GroupName="Laudos"
                        Checked="True" data-ToolTip="default" ToolTip="Informar se é aberto ou fechado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Físico:
                </div>
                <div class="coltxt wit">
                    <asp:RadioButton ID="optSim" runat="server" Text="Sim" GroupName="Fisico" data-ToolTip="default"
                        ToolTip="Informar se a operação é física ou não." />
                    <asp:RadioButton ID="optNao" runat="server" Text="Não" GroupName="Fisico" data-ToolTip="default"
                        ToolTip="Informar se a operação é física ou não." />
                    <asp:RadioButton ID="optGeral" runat="server" Text="Todos" GroupName="Fisico" Checked="True"
                        data-ToolTip="default" ToolTip="Informar se a operação é física ou não." />
                </div>
                <div class="collbl">
                    Transporte:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbTransporte" runat="server" Width="150px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Agrupar:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ckAgruparDeposito" runat="server" Text="Deposito" AutoPostBack="True"
                        OnCheckedChanged="ckAgruparDeposito_CheckedChanged" data-ToolTip="default" ToolTip="Marcar a opção quanto ao agrupamento." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdListaNormal" runat="server" Checked="True" GroupName="rdN"
                        Text="Relatório Normal" data-ToolTip="default" ToolTip="Ao invés de aparecer o Laudo agrupado, mostra os que compõe o Agrupamento" />
                    &nbsp; &nbsp;
                    <asp:RadioButton ID="rdListaSemAgrupar" runat="server" GroupName="rdN" Text="Relatório sem Agrupados"
                        data-ToolTip="default" ToolTip="Ao invés de aparecer o Laudo agrupado, mostra os que compõe o Agrupamento" />
                </div>
            </div>
            <div id="INTACTA" runat="server" class="row" visible="False">
                <div class="collbl">
                    Analise:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlAnalise" runat="server" />
                    <asp:TextBox ID="txtParametrosAnalise" runat="server" data-ToolTip="default" ToolTip="Valor simples ex: 14 analise que deu exatamente 14,  para intervalos use - ex: 10-14, será analise informada de 10 a 14 %,  para multiplios 2;3 ex: na analise intacta seria todos que dera 2 teste positivo e 3 participante."></asp:TextBox>
                    <%--<asp:RadioButton ID="rdITodos" runat="server" Text="Todos" data-ToolTip="default"
                        ToolTip="Listar Todos" Checked="True" GroupName="intacta" />
                    &nbsp; &nbsp;
                    <asp:RadioButton ID="rdISim" runat="server" Text="Apenas Intacta" data-ToolTip="default"
                        ToolTip="Lista apenas Laudos Intacta" GroupName="intacta" />
                    &nbsp; &nbsp;
                    <asp:RadioButton ID="rdPositivo" runat="server" Text="Apenas Teste Positivo" data-ToolTip="default"
                        ToolTip="Lista apenas Laudos Teste Positivo" GroupName="intacta" />
                    &nbsp; &nbsp;
                    <asp:RadioButton ID="rdPSim" runat="server" Text="Intacta e Teste Positivo" data-ToolTip="default"
                        ToolTip="Lista apenas Laudos Intacta" GroupName="intacta" />
                    &nbsp; &nbsp;
                    <asp:RadioButton ID="rdINao" runat="server" Text="Sem Intacta" data-ToolTip="default"
                        ToolTip="Lista apenas Laudos sem Intacta" GroupName="intacta" />--%>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
