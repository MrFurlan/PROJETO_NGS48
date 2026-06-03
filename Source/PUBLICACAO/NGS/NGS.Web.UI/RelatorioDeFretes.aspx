<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeFretes.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeFretes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioDeFretes" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioDeFretes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório CTRC / Recibo de Frete
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
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEmpresa" runat="server" Width="598px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClientes" runat="server" Width="558px" Enabled="False" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaCliente" OnClick="cmdConsultaCliente_Click" CssClass="btn"
                        runat="server" Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Transportador:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTransportador" runat="server" Width="558px" Enabled="false" />
                    <asp:HiddenField ID="txtCodigoTransportador" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnTransportador" runat="server" CssClass="btn" OnClick="BtnTransportador_Click"
                        Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Nome da pessoa/empresa responsável pelo transporte da mercadoria." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupoProduto" runat="server" Width="598px" OnSelectedIndexChanged="ddlGrupoProduto_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProduto" runat="server" Width="598px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt" style="width: 124px;">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="90px"
                        data-ToolTip="default" ToolTip="Informar o período inicial de consulta." />
                </div>
                <div class="collbl" style="margin-left: 5px;">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="90px"
                        data-ToolTip="default" ToolTip="Informar o período final de consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Width="90px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnPedido" OnClick="btnPedido_Click" runat="server" Text=">" UseSubmitBehavior="False"
                        CssClass="btn" data-ToolTip="default" ToolTip="Localizar o pedido." />
                </div>
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdCtrc" runat="server" Text="CTRC" GroupName="frete" data-ToolTip="default"
                        ToolTip="Selecionar se é CTRC ou Pedido." />
                    <asp:RadioButton ID="rdRecibo" runat="server" Text="Recibo" GroupName="frete" Checked="True"
                        data-ToolTip="default" ToolTip="Selecionar se é CTRC ou Pedido." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
