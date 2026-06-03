<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="EstoquesPorOperacoes.aspx.vb" Inherits="NGS.Web.UI.EstoquesPorOperacoes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngEstoquesPorOperacoes" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlEstoquesPorOperacoes" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Estoques Por Operações
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcelDados" runat="server" Text="Excel" />
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
                    <asp:DropDownList ID="ddlUnidade" ToolTip="Unidade de negócio empresarial." runat="server"
                        Width="600px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlDeposito" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
            </div>
            <div class="row">
                <div class="collbl">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdAtivo" runat="server" Text="Apenas Produtos Ativos" Checked="True" GroupName="Situacao" data-ToolTip="default" ToolTip="Listar apenas os Produtos com Situação Normal." />
                    <asp:RadioButton ID="rdInativo" runat="server" Text="Apenas Produto(s) Inativo(s)/Bloqueado(s)" GroupName="Situacao" data-ToolTip="default" ToolTip="Listar apenas os Produtos com a Situação diferente da Normal." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtro(s)
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbPorNome" runat="server" Checked="true" GroupName="rbfiltro" Text="Por Nome" data-ToolTip="default" ToolTip="" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbPorDescricaoMapa" runat="server" GroupName="rbfiltro" Text="Por Descrição Mapa" data-ToolTip="default" ToolTip="" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbTodos" runat="server" GroupName="rbfiltro" Text="Todos" data-ToolTip="default" ToolTip="" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Estoque:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadFisico" runat="server" Text="Fisico" GroupName="Estoque" data-ToolTip="default" ToolTip="Selecionar estoque físico ou fiscal." />
                    <asp:RadioButton ID="RadFiscal" runat="server" Text="Fiscal" Checked="True" GroupName="Estoque" data-ToolTip="default" ToolTip="Selecionar estoque físico ou fiscal." />
                </div>
                <div class="collbl" style="margin-left: 75px;">
                    Exercicio:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlExercicio" runat="server" Width="112px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="margin-left: 125px;">
                    <asp:CheckBox ID="CkCDeposito" runat="server" Text="Consolidar Depósito" data-ToolTip="default" ToolTip="" />
                    <asp:CheckBox ID="chkContraPartida" runat="server" Text="Gerar Contra Partida dos Depósitos de Terceiros" data-ToolTip="default" ToolTip="" />
                    <asp:CheckBox ID="chkOperacaoCusto" runat="server" Text="Operações que Afetam o Custo" data-ToolTip="default" ToolTip="" />
                    <asp:CheckBox ID="chkConsNotasComSerieEspecificas" runat="server" Text="Considerar Notas C/ Séries 101, 102, 103, 104"
                        Checked="True" data-ToolTip="default" ToolTip="" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Mês Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtDiaMesInicial" runat="server" Width="32px" MaxLength="2" data-ToolTip="default" ToolTip="Data inicial da consulta." />
                    <asp:DropDownList ID="ddlMesInicial" runat="server" Width="112px">
                        <asp:ListItem Value="01">Janeiro</asp:ListItem>
                        <asp:ListItem Value="02">Fevereiro</asp:ListItem>
                        <asp:ListItem Value="03">Mar&#231;o</asp:ListItem>
                        <asp:ListItem Value="04">Abril</asp:ListItem>
                        <asp:ListItem Value="05">Maio</asp:ListItem>
                        <asp:ListItem Value="06">Junho</asp:ListItem>
                        <asp:ListItem Value="07">Julho</asp:ListItem>
                        <asp:ListItem Value="08">Agosto</asp:ListItem>
                        <asp:ListItem Value="09">Setembro</asp:ListItem>
                        <asp:ListItem Value="10">Outubro</asp:ListItem>
                        <asp:ListItem Value="11">Novembro</asp:ListItem>
                        <asp:ListItem Value="12">Dezembro</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="collbl">
                    Mês Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtDiaMesFinal" runat="server" Width="32px" MaxLength="2" data-ToolTip="default" ToolTip="Data final da consulta." />
                    <asp:DropDownList ID="ddlMesFinal" runat="server" Width="112px">
                        <asp:ListItem Value="01">Janeiro</asp:ListItem>
                        <asp:ListItem Value="02">Fevereiro</asp:ListItem>
                        <asp:ListItem Value="03">Mar&#231;o</asp:ListItem>
                        <asp:ListItem Value="04">Abril</asp:ListItem>
                        <asp:ListItem Value="05">Maio</asp:ListItem>
                        <asp:ListItem Value="06">Junho</asp:ListItem>
                        <asp:ListItem Value="07">Julho</asp:ListItem>
                        <asp:ListItem Value="08">Agosto</asp:ListItem>
                        <asp:ListItem Value="09">Setembro</asp:ListItem>
                        <asp:ListItem Value="10">Outubro</asp:ListItem>
                        <asp:ListItem Value="11">Novembro</asp:ListItem>
                        <asp:ListItem Value="12">Dezembro</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
