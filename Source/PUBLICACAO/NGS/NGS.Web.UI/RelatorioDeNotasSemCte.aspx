<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeNotasSemCte.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeNotasSemCte" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioDeNotasSemCte" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioDeNotasSemCte" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório De Notas Sem Cte
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
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="622px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Width="582px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button CssClass="btn" ID="btnBuscaCliente" OnClick="btnBuscaCliente_Click" runat="server"
                        Text=">" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Width="100" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoPedido" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnPedido" runat="server" Text=">" UseSubmitBehavior="False" OnClick="btnPedido_Click"
                        CssClass="btn" data-ToolTip="default" ToolTip="Informar o número do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Transportador:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoTransportador" runat="server" />
                    <asp:TextBox ID="TxtTransportador" runat="server" Enabled="False" Width="582px" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnConsultarTransportador" runat="server" OnClick="BtnConsultarTransportador_Click"
                        Text=">" UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Nome da pessoa/empresa responsável pelo transporte da mercadoria." />
                </div>
            </div>
            <div class="row">
                <div class="collbl colw">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="86px"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="86px"
                        data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Entrada/Saída:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbEntrada" GroupName="EntSai" runat="server" Text="Entrada"
                        Checked="True" data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                    <asp:RadioButton ID="rbSaida" GroupName="EntSai" runat="server" Text="Saída" data-ToolTip="default"
                        ToolTip="Informar se a nota é de entrada ou saída." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
