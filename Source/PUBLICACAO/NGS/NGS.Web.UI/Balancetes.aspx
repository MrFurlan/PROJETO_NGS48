<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Balancetes.aspx.vb" Inherits="NGS.Web.UI.Balancetes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadBalancetes() {
            $("#MainContent_lstCentroDeCusto").multiselect({
                header: "Escolha os centro de custos!",
                selectedList: 2
            }).multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadBalancetes();
            var prmBalancetes = Sys.WebForms.PageRequestManager.getInstance();
            prmBalancetes.add_endRequest(pageLoadBalancetes);
        });
    </script>
    <style type="text/css">
        .painelleft {
            width: 200px;
            margin-right: 4px;
        }

            .painelleft .row .coltxt {
                line-height: 10px;
            }

            .painelleft .subtitulodiv {
                margin-top: 2px;
            }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngBalancetes" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlBalancetes" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Balancetes Auxiliares
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
                                    <asp:LinkButton ID="lnkExcelDados" runat="server" Text="Excel Dados" ToolTip="Disponível apenas para Balancete Aux. Completo." data-ToolTip="default" />
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
                    <asp:CheckBox ID="chkConsolidarUnidade" runat="server" Text="Cons. Unidade:"
                        ToolTip="Consolida os gastos da unidade na empresa selecionada." /><br />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" AutoPostBack="True" Enabled="False"
                        OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged" Width="672px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="line-height: 14px; padding-left: 7px; padding-bottom: 2px; text-indent: 0; height: auto; width: 106px; margin-bottom: 4px;">
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." runat="server"
                        Text="Cons. Empresa:" data-ToolTip="default" /><br />
                    <asp:CheckBox ID="chkConsolidarSede" runat="server" Text="Cons. Filial:" data-ToolTip="default"
                        ToolTip="Consolida a Empresa, independente de endereçamento." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="672px" Enabled="False" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Periodo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Enabled="False"
                        Width="88px" data-ToolTip="default" ToolTip="Data inicial a finaldo relatório." />
                    &nbsp;à&nbsp;
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Enabled="False"
                        Width="88px" data-ToolTip="default" ToolTip="Data inicial a finaldo relatório." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Emissão:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataDeEmissao" runat="server" Width="88px" Enabled="False" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Dia da consulta." />
                </div>
                <div class="collbl">
                    Folha Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtIniciarNaFolha" runat="server" Width="56px" Enabled="False" data-ToolTip="default"
                        ToolTip="Referente a contagem das páginas." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtro Adicional:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkTitulos" runat="server" Text="Executar Comp. Títulos" data-ToolTip="default"
                        ToolTip="" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Contas:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdTodasAsContas" runat="server" Checked="True" GroupName="exibir"
                        Text="Todas" data-ToolTip="default" ToolTip="Selecionar quais contas serão consultadas quanto ao saldo." />
                    <asp:RadioButton ID="rdSomenteComSaldoAtual" runat="server" GroupName="exibir" Text="Somente Com Saldo Atual"
                        data-ToolTip="default" ToolTip="Selecionar quais contas serão consultadas quanto ao saldo." />
                    <asp:RadioButton ID="rdComMovimentoNoPeriodo" runat="server" GroupName="exibir" Text="Com Movimento No Periodo / Saldo"
                        data-ToolTip="default" ToolTip="Selecionar quais contas serão consultadas quanto ao saldo." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Seleção:
                </div>
                <div class="painelleft">
                    <div class="subtitulodiv">
                        Modelos:
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:RadioButtonList ID="Modelo" runat="server" AutoPostBack="True" Enabled="False"
                                OnSelectedIndexChanged="Modelo_SelectedIndexChanged" Width="163px" data-ToolTip="default"
                                ToolTip="Selecionar os campos desejados conforme necessidade.">
                                <asp:ListItem Selected="True" Value="1">Balancete Auxiliar</asp:ListItem>
                                <asp:ListItem Value="5">Balancete Aux. Completo</asp:ListItem>
                                <asp:ListItem Value="3">Anexo por Cliente</asp:ListItem>
                                <asp:ListItem Value="4">Anexo por Conta/Produto</asp:ListItem>
                                <asp:ListItem Value="6">Anexo por Produto/Conta</asp:ListItem>
                                <asp:ListItem Value="2">Centros de Custos</asp:ListItem>
                            </asp:RadioButtonList>
                        </div>
                    </div>
                </div>
                <div class="painelleft">
                    <div class="subtitulodiv">
                        Grupos de Contas
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBoxList ID="Grupos" runat="server" Enabled="False" Width="150px" data-ToolTip="default"
                                ToolTip="Selecionar os grupos desejados conforme necessidade.">
                                <asp:ListItem Selected="True" Value="1">1 - Ativo</asp:ListItem>
                                <asp:ListItem Selected="True" Value="2">2 - Passivo</asp:ListItem>
                                <asp:ListItem Selected="True" Value="3">3 - Receitas/Despesas</asp:ListItem>
                                <asp:ListItem Selected="True" Value="4">4 - ...</asp:ListItem>
                                <asp:ListItem Selected="True" Value="5">5 - Resultado</asp:ListItem>
                            </asp:CheckBoxList>
                        </div>
                    </div>
                </div>
                <div class="painelleft">
                    <div class="subtitulodiv">
                        Níveis de Contas
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBoxList ID="Niveis" runat="server" Enabled="False" Width="200px" data-ToolTip="default"
                                ToolTip="Selecionar os campos desejados conforme necessidade.">
                                <asp:ListItem Selected="True" Value="1">1 - Ativo</asp:ListItem>
                                <asp:ListItem Selected="True" Value="2">101 - Circulante</asp:ListItem>
                                <asp:ListItem Selected="True" Value="3">10101 - Disponivel</asp:ListItem>
                                <asp:ListItem Selected="True" Value="4">1010101 - Caixa</asp:ListItem>
                                <asp:ListItem Selected="True" Value="5">101010101 - Caixa</asp:ListItem>
                            </asp:CheckBoxList>
                        </div>
                    </div>
                </div>
                <asp:Panel ID="pnlNivelDeProduto" runat="server">
                    <div class="painelleft">
                        <div class="subtitulodiv">
                            Nivel Produto
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:RadioButtonList ID="rdGrupoProduto" runat="server" data-ToolTip="default" ToolTip="Selecionar os campos desejados conforme necessidade.">
                                    <asp:ListItem>1</asp:ListItem>
                                    <asp:ListItem Value="2">10</asp:ListItem>
                                    <asp:ListItem Value="3">101</asp:ListItem>
                                    <asp:ListItem Value="5">10101</asp:ListItem>
                                    <asp:ListItem Selected="True" Value="0">Produto</asp:ListItem>
                                </asp:RadioButtonList>
                            </div>
                        </div>
                    </div>
                </asp:Panel>
            </div>
            <div class="row" runat="server">
                <div class="coltxt" style="margin-left: 125px; width: 85.5%">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row" runat="server">
                <asp:Panel ID="pnlCentroDeCusto" runat="server" Style="margin-left: 125px">
                    <div class="painelleft" style="width: 150px">
                        <div class="subtitulodiv">
                            Níveis de Custos
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:RadioButtonList ID="NiveisDeCusto" runat="server" AutoPostBack="True" OnSelectedIndexChanged="NiveisDeCusto_SelectedIndexChanged"
                                    Width="150px" data-ToolTip="default" ToolTip="Selecionar os campos desejados conforme necessidade.">
                                    <asp:ListItem Value="1">1 - Nivel 1</asp:ListItem>
                                    <asp:ListItem Value="3">101 - Nivel 2</asp:ListItem>
                                    <asp:ListItem Selected="True" Value="5">10101 - Nivel 3</asp:ListItem>
                                </asp:RadioButtonList>
                            </div>
                        </div>
                    </div>
                    <div class="painelleft" style="width: 325px;">
                        <div class="subtitulodiv">
                            Opções Centro de Custo
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:ListBox ID="lstCentroCusto" runat="server" Width="150px" Height="70px" data-ToolTip="default"
                                    ToolTip="Selecionar os campos desejados conforme necessidade.">
                                    <asp:ListItem Value="0">Safra</asp:ListItem>
                                    <asp:ListItem Value="1">Unidade de Negocio</asp:ListItem>
                                    <asp:ListItem Value="2">Empresa</asp:ListItem>
                                    <asp:ListItem Value="3">Centro De Custo</asp:ListItem>
                                </asp:ListBox>
                                <br />
                                <asp:CheckBox ID="chkAnaliticoPorNota" runat="server" Font-Bold="true" Text="Analitico Por Nota" />
                            </div>
                            <div class="coltxt" style="top: 20px;">
                                <asp:ImageButton ID="imgAdiciona" runat="server" ImageUrl="~/images/ico-mais.gif"
                                    OnClick="imgAdiciona_Click" Style="width: 12px; margin-bottom: 2px;" />
                                <br />
                                <asp:ImageButton ID="imgRemove" runat="server" ImageUrl="~/images/ico-menos.gif"
                                    OnClick="imgRemove_Click" />
                            </div>
                            <div class="coltxt">
                                <asp:ListBox ID="lstCentroCustoSelecionados" runat="server" Width="150px" Height="70px" />
                            </div>
                        </div>
                    </div>
                    <div class="painelleft" style="width: 325px;">
                        <div class="subtitulodiv">
                            Centro de Custo
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:ListBox ID="lstCentroDeCusto" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                                    Width="325px" />
                            </div>
                        </div>
                    </div>
                </asp:Panel>
            </div>
            <div class="row">
                <div class="painelleft" style="margin-left: 125px">
                    <div class="subtitulodiv">
                        Isolar Lotes
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBoxList ID="IsolarLotes" runat="server" Enabled="False" Width="200px" data-ToolTip="default"
                                ToolTip="Selecionar os campos desejados conforme necessidade.">
                                <asp:ListItem Value="7000">7000 - Apuração De Custos</asp:ListItem>
                                <asp:ListItem Value="7001">7001 - Avaliação de Estoques</asp:ListItem>
                                <asp:ListItem Value="7500">7500 - Zeramento de Resultados</asp:ListItem>
                                <asp:ListItem Value="7600">7600 - Zeramento dos Estoques</asp:ListItem>
                            </asp:CheckBoxList>
                        </div>
                    </div>
                </div>
                <div class="painelleft">
                    <div class="subtitulodiv">
                        Isolar Compensação
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBoxList ID="IsolarCompensacao" runat="server" Enabled="False" Width="150px"
                                data-ToolTip="default" ToolTip="Selecionar os campos desejados conforme necessidade.">
                                <asp:ListItem Selected="True" Value="105">105 - Ativos</asp:ListItem>
                                <asp:ListItem Selected="True" Value="205">205 - Passivos</asp:ListItem>
                                <asp:ListItem>...</asp:ListItem>
                            </asp:CheckBoxList>
                        </div>
                    </div>
                </div>
                <div class="painelleft">
                    <div class="subtitulodiv">
                        Moeda
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlMoeda" runat="server" Height="19px" Width="135px" />
                        </div>
                    </div>
                </div>
                <div class="painelleft">
                    <div class="subtitulodiv">
                        Outros Parametros
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chkZeraContaDeResultado" runat="server" Style="line-height: 15px"
                                Text="Trazer Grupo de Contas 3 Saldo Inicial Zerado no Inicio do Periodo" />
                            <br />
                            <asp:CheckBox ID="chkpiscofins56" runat="server" Text="Situação Pis/Cofins" />
                            <br />
                            <asp:TextBox ID="txtSituacaoPISCOFINS" runat="server" Width="200px" />
                        </div>
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo Contas:
                </div>
                <div class="coltxt" style="width: 86%">
                    <div class="subtitulodiv">
                        Seleção de Grupos de Contas
                    </div>
                </div>
            </div>
            <div class="row" style="line-height: 12px;">
                <div class="bordagrid" style="height: 200px; margin-left: 125px; width: 85%;">
                    <div class="painelleft" style="width: 49%;">
                        <asp:CheckBoxList ID="SelecaoDeGrupos1" runat="server" />
                    </div>
                    <div class="painelleft" style="width: 49%;">
                        <asp:CheckBoxList ID="SelecaoDeGrupos2" runat="server" />
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
