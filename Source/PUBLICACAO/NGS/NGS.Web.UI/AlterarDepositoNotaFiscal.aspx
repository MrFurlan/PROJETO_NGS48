<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="AlterarDepositoNotaFiscal.aspx.vb" Inherits="NGS.Web.UI.AlterarDepositoNotaFiscal" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngAlterarDepositoNotaFiscal" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlAlterarDepositoNotaFiscal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Alterar Depósito
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
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
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtClientes" runat="server" Width="613px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaCliente" runat="server" Text=" > " OnClick="cmdConsultaCliente_Click"
                        CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" Width="86px" CssClass="calendario" runat="server"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" Width="86px" runat="server"
                        data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nota Fiscal:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNotaFiscal" CssClass="txtNumerico" runat="server" MaxLength="9"
                        Width="86px" data-ToolTip="default" ToolTip="Número da Nota Fiscal." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDepositos" Enabled="false" runat="server" Width="613px" />
                    <asp:HiddenField ID="txtCodigoDeposito" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnDeposito" OnClick="btnDeposito_Click" runat="server" UseSubmitBehavior="False"
                        Text=">" CssClass="btn" data-ToolTip="default" ToolTip="Depósito da própria empresa onde a mercadiria está entrando ou saindo." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Origem/Destino:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtOriDess" Enabled="false" runat="server" Width="613px" />
                    <asp:HiddenField ID="txtCodigoOriDes" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnOriDes" CssClass="btn" runat="server" UseSubmitBehavior="False"
                        Text=" > " OnClick="btnOriDes_Click" data-ToolTip="default" ToolTip="Depósito do cliente/fornecedor." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtNomeDoCliente" runat="server" data-ToolTip="default" ToolTip="Nome do Cliente." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    E/S:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtEntSai" runat="server" data-ToolTip="default" ToolTip="" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Série:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtSerie" runat="server" data-ToolTip="default" ToolTip="Série da Nota." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Nota:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtNota" runat="server" data-ToolTip="default" ToolTip="Número da Nota Fiscal." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
</asp:Content>
