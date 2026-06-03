<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="RelatorioDeVendasPorRepresentante.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeVendasPorRepresentante" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script src="<%# ResolveUrl("~/Scripts/App/relatoriodevenda.js")%>" type="text/javascript"></script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">'
    <asp:ScriptManager ID="scmngBaixasFinanceirasIndividuais" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <div class="titulodiv">
        Relatório de Vendas.
    </div>
    <asp:UpdatePanel ID="updpnlVendasPorRepresentante" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf" >
                                    <asp:LinkButton ID="lnkRelatorio" runat="server" Text="Relatório Pdf" />
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
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="598px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkEmpresaCons" runat="server" Text="Empresa Cons." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="598px" />
                </div>
            </div>
            <div class="row paramexercicio">
                <div class="collbl">
                    Mês Base:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMesBase" runat="server" Width="110px">
                        <asp:ListItem Value="1" Text="Janeiro" />
                        <asp:ListItem Value="2" Text="Fevereiro" />
                        <asp:ListItem Value="3" Text="Março" />
                        <asp:ListItem Value="4" Text="Abril" />
                        <asp:ListItem Value="5" Text="Maio" />
                        <asp:ListItem Value="6" Text="Junho" />
                        <asp:ListItem Value="7" Text="Julho" />
                        <asp:ListItem Value="8" Text="Agosto" />
                        <asp:ListItem Value="9" Text="Setembro" />
                        <asp:ListItem Value="10" Text="Outubro" />
                        <asp:ListItem Value="11" Text="Novembro" />
                        <asp:ListItem Value="12" Text="Dezembro" />
                    </asp:DropDownList>
                </div>
                <div class="collbl">
                    Exercício:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlAno" runat="server" Width="120px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo Emissão:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbPorRepresentante" runat="server" Checked="true" GroupName="TipoEmissao" Text="Por Representante" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbPorCliente" runat="server" GroupName="TipoEmissao" Text="Por Cliente" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbPorPedido" runat="server" GroupName="TipoEmissao" Text="Por Pedido" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:Label ID="lblCliente" runat="server" Text="Representante:" />
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
            <div class="row paramporpedido">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt ">
                    <asp:TextBox ID="txtData1" runat="server" Width="86px" CssClass="calendario" />
                </div>
                <div class="coltxt">
                    Á
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData2" runat="server" Width="86px" CssClass="calendario" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkRepresentanteNaoInformado" runat="server" Text="Emitir apenas pedido sem representante." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
