<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ExtratoDePedido.aspx.vb" Inherits="NGS.Web.UI.ExtratoDePedido" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngExtratoDePedido" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlExtratoDePedido" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Extrato de Pedido
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server"
                                OnClick="lnkConsultar_Click" Text="Consultar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server"
                                OnClick="lnkLimpar_Click" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server"
                                OnClick="lnkAjuda_Click" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEmpresa" runat="server" Width="638px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClientes" runat="server" Width="598px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button CssClass="btn" ID="cmdBuscaCliente" OnClick="cmdBuscaCliente_Click" runat="server"
                        Text=">" CausesValidation="False" UseSubmitBehavior="False" data-tooltip="default"
                        ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbGrupoProduto" runat="server" Width="638px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbProduto" runat="server" Width="638px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Parâmetros:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkNomeProduto" runat="server" Text="Mostrar Descrição do Produto:" data-ToolTip="default" ToolTip="Selecionar os parâmetros a serem exibidos no relatório." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkOcultarFinanceiro" runat="server" Text="Ocultar Financeiro" data-ToolTip="default" ToolTip="Selecionar os parâmetros a serem exibidos no relatório." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkOcultarFrete" runat="server" Text="Ocultar Frete" data-ToolTip="default" ToolTip="Selecionar os parâmetros a serem exibidos no relatório." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkOcultarLancamentoContabil" runat="server" Text="Ocultar Lançamento Contábil" data-ToolTip="default" ToolTip="Selecionar os parâmetros a serem exibidos no relatório." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkSintetico" runat="server" Text="Relatório Sintético" data-ToolTip="default" ToolTip="Selecionar os parâmetros a serem exibidos no relatório." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkPesagem" runat="server" Text="Pesagem" data-ToolTip="default" ToolTip="Selecionar os parâmetros a serem exibidos no relatório." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Limite:
                </div>
                <div class="coltxt" style="width: 140px;">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="100px" data-ToolTip="default" ToolTip="Informar a data limite que deseja consultar." />
                </div>
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="300px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Width="100px" CssClass="txtNumerico" data-ToolTip="default" ToolTip="Número do pedido a ser consultado." />
                </div>
                <div class="coltxt">
                    <asp:Button CssClass="btn" ID="cmdBuscaPedido" OnClick="cmdBuscaPedido_Click" runat="server"
                        Text="&gt;" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Número do pedido a ser consultado." />
                </div>
                <div class="collbl">
                    Pedido Efetivo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedidoEfetivo" runat="server" Width="100px" data-ToolTip="default" ToolTip="Número do pedido efetivo a ser consultado." />
                </div>
                <div class="coltxt">
                    <asp:Button CssClass="btn" ID="cmdBuscaPedidoEfetivo" OnClick="cmdBuscaPedidoEfetivo_Click" runat="server"
                        Text="&gt;" CausesValidation="False" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Número do pedido efetivo a ser consultado." />
                </div>
                <div class="collbl">
                    Entrada/Saída:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkEntrada" runat="server" Text="Entrada" Checked="True" data-ToolTip="default" ToolTip="Selecionar se o pedido é uma entrada ou uma saída." />
                    <asp:CheckBox ID="chkSaida" runat="server" Text="Saída" Checked="True" data-ToolTip="default" ToolTip="Selecionar se o pedido é uma entrada ou uma saída." />
                </div>
            </div>
            <div class="row">
                 <div class="collbl">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdAberto" runat="server" Checked="True" GroupName="pedAF" Text="Aberto (Fiscal/Financ.)" data-ToolTip="default" ToolTip="Situação dos pedidos para listagem." />
                    <asp:RadioButton ID="rdFechado" runat="server" GroupName="pedAF" Text="Fechado (Fiscal/Financ.)" data-ToolTip="default" ToolTip="Situação dos pedidos para listagem." />
                    <asp:RadioButton ID="rdTodos" runat="server" GroupName="pedAF" Text="Todos" data-ToolTip="default" ToolTip="Situação dos pedidos para listagem." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Emissão em:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdPdf" runat="server" Text="PDF" Checked="True" GroupName="extrato" data-ToolTip="default" ToolTip="Formato desejado para impressão do relatório." />
                    <asp:RadioButton ID="rdHtml" runat="server" Text="HTML" GroupName="extrato" data-ToolTip="default" ToolTip="Formato desejado para impressão do relatório." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
