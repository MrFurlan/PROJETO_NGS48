<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeComissoes.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeComissoes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioDeComissoes" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioDeComissoes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório De Comissões
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
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" OnClick="lnkLimpar_Click" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Relatorio:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbRepresentante" runat="server" Checked="True" GroupName="agrupamento"
                        Text="Por Representante / Cliente / Pedido" data-ToolTip="default" ToolTip="Selecionar de acordo com o relatório desejado." />
                    <asp:RadioButton ID="rbProduto" runat="server" GroupName="agrupamento" Text="Por Produto" 
                        data-ToolTip="default" ToolTip="Selecionar de acordo com o relatório desejado." />
                    <asp:RadioButton ID="rdPorFaixa" runat="server" GroupName="agrupamento" Text="Por Faixa de Comissao" 
                        data-ToolTip="default" ToolTip="Selecionar de acordo com o relatório desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidaEmpresa" runat="server" Text="Consol. Empresa:" data-ToolTip="default" ToolTip="Empresa Para negociação/consulta." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="628px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Representante:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtRepresentante" runat="server" Enabled="false" Width="589px" />
                    <asp:HiddenField ID="txtCodigoRepresentante" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaRepresentante" runat="server" CausesValidation="False" OnClick="cmdBuscaRepresentante_Click"
                        Text=">" UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o representante para consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClientes" runat="server" Width="589px" Enabled="false" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaCliente" OnClick="cmdBuscaCliente_Click" runat="server" Text=">"
                        CausesValidation="False" UseSubmitBehavior="False" CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="255px" />
                </div>
                <div class="collbl">
                    Marca:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMarca" runat="server" Width="245px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMoeda" runat="server" Width="255px" />
                </div>
                <div class="collbl">
                    Troca:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlTroca" runat="server" Width="245px">
                        <asp:ListItem></asp:ListItem>
                        <asp:ListItem Value="1">Com Troca</asp:ListItem>
                        <asp:ListItem Value="0">Sem Troca</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 100%;">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 100%;">
                    <div class="subtitulodiv" style="margin-top: 9px;">
                        Colunas
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:CheckBox ID="chkCompraReais" runat="server" Checked="True" Text="Compra R$" data-ToolTip="default" ToolTip="" />&nbsp;
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="chkVendaReais" runat="server" Checked="True" Text="Venda R$" data-ToolTip="default" ToolTip="" />&nbsp;
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="chkMargem" runat="server" Checked="True" Text="Margem" data-ToolTip="default" ToolTip="" />&nbsp;
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="chkIndice" runat="server" Checked="True" Text="Indice" data-ToolTip="default" ToolTip="" />&nbsp;
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="chkCompraDolar" runat="server" Checked="True" Text="Compra U$" data-ToolTip="default" ToolTip="" />&nbsp;
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="chkVendaDolar" runat="server" Checked="True" Text="Venda U$" data-ToolTip="default" ToolTip="" />&nbsp;
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="chkQuantidade" runat="server" Checked="True" Text="Quantidade" data-ToolTip="default" ToolTip="" />&nbsp;
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="chkTotalReais" runat="server" Checked="True" Text="Total R$" data-ToolTip="default" ToolTip="" />&nbsp;
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="chkTotalDolar" runat="server" Checked="True" Text="Total U$" data-ToolTip="default" ToolTip="" />&nbsp;
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="chkComissaoReais" runat="server" Checked="True" Text="Comissao R$" data-ToolTip="default" ToolTip="" />&nbsp;
                    &nbsp; &nbsp;
                    <asp:CheckBox ID="chkComissaoDolar" runat="server" Checked="True" Text="Comissao U$" data-ToolTip="default" ToolTip="" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
