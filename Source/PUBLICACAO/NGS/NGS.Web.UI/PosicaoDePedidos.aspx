<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PosicaoDePedidos.aspx.vb" Inherits="NGS.Web.UI.PosicaoDePedidos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 140px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngESNotas" runat="server" AsyncPostBackTimeout="10000" />
    <asp:HiddenField ID="HID" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPosicaoDePedidos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Posição de Pedidos
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
                    Posição no dia:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" runat="server" CssClass="calendario" Width="96px" data-ToolTip="default"
                        ToolTip="Selecionar a data desejada para consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbUnidadeNegocio" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbUnidadeNegocio_SelectedIndexChanged"
                        Width="672px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." Text="Consolidar Empresa:"
                        runat="server" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEmpresa" runat="server" Width="672px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente."
                        runat="server" Text="Consolidar Cliente:" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" TabIndex="3" runat="server" Width="632px" Font-Names="monospace"
                        Enabled="False" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaClientes" TabIndex="4" OnClick="cmdConsultaClientes_Click"
                        runat="server" Text=">" UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Unificar as informações por cliente." />
                </div>
            </div>
            <div>
                <div class="collbl">
                    Representante:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtRepresentante" TabIndex="3" runat="server" Width="632px" Font-Names="monospace"
                        Enabled="False" />
                    <asp:HiddenField ID="txtCodigoRepresentante" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaRepresentante" TabIndex="5" OnClick="cmdConsultaRepresentante_Click"
                        runat="server" Text=">" UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Consulta de Representante." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe Operacao:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbClasse" runat="server" AutoPostBack="True" DataMember="ClassesDeOperacoes"
                        DataTextField="Descricao" DataValueField="Classe_Id" Width="672px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe SubOperação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlClasseSubOperacao" runat="server" Width="672px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbSafra" runat="server" Width="672px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 100%; margin-bottom: 7px; margin-top: 1px;">
                    <uc:SelecaoOperacoes ID="ucSelecaoOperacoes" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 100%; margin-bottom: 7px; margin-top: 1px;">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedidos:
                </div>
                <div class="coltxt" style="width: 347px;">
                    <asp:RadioButton ID="rdComSaldo" runat="server" Checked="True" GroupName="TipoRelatorio"
                        Text="Com Saldo" data-ToolTip="default" ToolTip="Status dos pedidos para listagem." />
                    <asp:RadioButton ID="rdLiquidados" runat="server" GroupName="TipoRelatorio" Text="Liquidados" data-ToolTip="default" ToolTip="Status dos pedidos para listagem." />
                    <asp:RadioButton ID="rdTodos" runat="server" GroupName="TipoRelatorio" Text="Todos" data-ToolTip="default" ToolTip="Status dos pedidos para listagem." />
                </div>
                <div class="collbl" style="width: 160px; margin-left: 2px;">
                    Listar Cessionario:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadCessionarioSim" runat="server" GroupName="cessionario" Text="Sim" data-ToolTip="default" ToolTip="Selecionar se o pedido é cessionário." />
                    <asp:RadioButton ID="RadCessionarioNao" runat="server" Checked="True" GroupName="cessionario"
                        Text="Não" data-ToolTip="default" ToolTip="Selecionar se o pedido é cessionário." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdAberto" runat="server" Checked="True" GroupName="pedAF" Text="Aberto (Fiscal/Financ.)" data-ToolTip="default" ToolTip="Situação dos pedidos para listagem." />
                    <asp:RadioButton ID="rdFechado" runat="server" GroupName="pedAF" Text="Fechado (Fiscal/Financ.)" data-ToolTip="default" ToolTip="Situação dos pedidos para listagem." />
                    <asp:RadioButton ID="rdTodosAF" runat="server" GroupName="pedAF" Text="Todos" data-ToolTip="default" ToolTip="Situação dos pedidos para listagem." />
                </div>
                <div class="collbl" style="width: 160px; margin-left: 18px;">
                    Fretes:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlFretes" runat="server">
                        <asp:ListItem Value="0">TODOS</asp:ListItem>
                        <asp:ListItem Value="1">CIF</asp:ListItem>
                        <asp:ListItem Value="2">FOB</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedidos por Troca:
                </div>
                <div class="coltxt" style="width: 348px;">
                    <asp:RadioButton ID="rdTrocaSim" runat="server" GroupName="Ant" Text="Sim" data-ToolTip="default" ToolTip="Selecionar se o pedido é de troca ou não." />
                    <asp:RadioButton ID="rdTrocaNao" runat="server" GroupName="Ant" Text="Não" data-ToolTip="default" ToolTip="Selecionar se o pedido é de troca ou não." />
                    <asp:RadioButton ID="rdTrocaTodos" runat="server" Checked="True" GroupName="Ant"
                        Text="Todos" data-ToolTip="default" ToolTip="Selecionar se o pedido é de troca ou não." />
                </div>
                <div class="collbl" style="width: 160px;">
                    <asp:CheckBox ID="chkDataEntregaPedido" runat="server" AutoPostBack="True" Text="Data Entrega Pedido: " data-ToolTip="default" ToolTip="Marcar se desejar inserir data de entrega do pedido para consulta." />
                </div>
                <div class="coltxt">
                    <asp:Panel ID="pnlDataEntregaPedido" runat="server" Visible="False">
                        <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="96px" />
                        &nbsp;a
                        <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="96px" />
                    </asp:Panel>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação Antecipada:
                </div>
                <div class="coltxt" style="width: 348px;">
                    <asp:RadioButton ID="rdAntecipadaSim" runat="server" GroupName="troca" Text="Sim" data-ToolTip="default" ToolTip="Selecionar se o pedido é uma operação antecipada ou não." />
                    <asp:RadioButton ID="rdAntecipadaNao" runat="server" GroupName="troca" Text="Não" data-ToolTip="default" ToolTip="Selecionar se o pedido é uma operação antecipada ou não." />
                    <asp:RadioButton ID="rdAntecipadaTodos" runat="server" Checked="True" GroupName="troca"
                        Text="Todos" data-ToolTip="default" ToolTip="Selecionar se o pedido é uma operação antecipada ou não." />
                </div>
                <div class="collbl" style="width: 160px;">
                    <asp:CheckBox ID="ckDataDeAbertura" runat="server" AutoPostBack="True" Text="Data Lançamento Pedido:" data-ToolTip="default" ToolTip="Marcar se desejar inserir data de lançamento do pedido para consulta." />
                </div>
                <div class="coltxt">
                    <asp:Panel ID="pnlDataAberturaPedido" runat="server" Visible="False">
                        <asp:TextBox ID="txtDataInicialAbertura" runat="server" CssClass="calendario" Width="96px" />
                        &nbsp;a
                        <asp:TextBox ID="txtDataFinalAbertura" runat="server" CssClass="calendario" Width="96px" />
                    </asp:Panel>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Recompra:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdReSim" runat="server" GroupName="Re" Text="Sim" data-ToolTip="default" ToolTip="Selecionar se o pedido é uma recompra ou não." />
                    <asp:RadioButton ID="rdReNao" runat="server" GroupName="Re" Text="Não" data-ToolTip="default" ToolTip="Selecionar se o pedido é uma recompra ou não." />
                    <asp:RadioButton ID="rdReTodos" runat="server" GroupName="Re" Text="Todos" Checked="True" data-ToolTip="default" ToolTip="Selecionar se o pedido é uma recompra ou não." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
