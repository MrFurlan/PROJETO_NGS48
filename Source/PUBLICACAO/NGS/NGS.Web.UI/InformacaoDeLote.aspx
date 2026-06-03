<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="InformacaoDeLote.aspx.vb" Inherits="NGS.Web.UI.InformacaoDeLote" %>


<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngInformacaoDeLote" runat="server" AsyncPostBackTimeout="1000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlInformacaoDeLote" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Informações de Lote
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server"
                                Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server"
                                Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="600px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdCliente" CssClass="btn" OnClick="cmdCliente_Click" runat="server"
                        Text=">" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbOperacao" runat="server" Width="294px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbOperacao_SelectedIndexChanged" />
                    <asp:DropDownList ID="cmbSubOperacao" runat="server" Width="295px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="96px" data-ToolTip="default" ToolTip="Informar o data inicial de consulta." />
                </div>
                <div class="collbl">
                    Período Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="96px" data-ToolTip="default" ToolTip="Informar o data final de consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Ordem:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdValidade" runat="server" Text="Validade" Checked="True" GroupName="Ordem" data-ToolTip="default" ToolTip="Ordena relatório por Validade do Lote." />
                    <asp:RadioButton ID="rdMovimento" runat="server" Text="Movimento" GroupName="Ordem" data-ToolTip="default" ToolTip="Ordena relatório por Movimento." />
                    <asp:RadioButton ID="rdProduto" runat="server" Text="Produto" GroupName="Ordem" data-ToolTip="default" ToolTip="Ordena o relatório por Produto." />
                    <asp:RadioButton ID="rdCliente" runat="server" Text="Cliente" GroupName="Ordem" data-ToolTip="default" ToolTip="Ordena o relatório por Cliente." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>