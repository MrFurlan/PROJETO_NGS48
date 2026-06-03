<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CurvaABC.aspx.vb" Inherits="NGS.Web.UI.CurvaABC" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style>
        .collbl {
            width: 140px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCurvaABC" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCurvaABC" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Curva ABC
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton class="iconPdf" ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel" />
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
                    <asp:CheckBox ID="chkConsolidar" runat="server" CssClass="rotulo" Text="Consolidar Empresas:" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="561px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:RadioButton ID="rdPeriodo" runat="server" Checked="True" GroupName="periodo" Text="Período:" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" data-ToolTip="default"
                        ToolTip="Data inicial a final do pedido." />
                </div>
                <div class="collbl">
                    à:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" data-ToolTip="default"
                        ToolTip="Data inicial a final do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:RadioButton ID="rdAnoSafra" runat="server" GroupName="periodo"
                        Text="Ano Safra:" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlAno" runat="server" Width="77px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Posição:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdProduto" runat="server" GroupName="posicao" Text="Por Produto"
                        data-ToolTip="default" ToolTip="Selecionar se o relatório será gerado por produto, cliente ou consolidado por cliente." />
                    &nbsp;<asp:RadioButton ID="rdCliente" runat="server" Checked="True" GroupName="posicao"
                        Text="Por Cliente" data-ToolTip="default" ToolTip="Selecionar se o relatório será gerado por produto, cliente ou consolidado por cliente." />
                    &nbsp;<asp:CheckBox ID="chkConsolidarCliente" data-ToolTip="default" ToolTip="Selecionar se o relatório será gerado por produto, cliente ou consolidado por cliente." runat="server"
                        Text="Consolidar Cliente" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operações:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdCompras" runat="server" GroupName="operacoes" Text="Compras"
                        data-ToolTip="default" ToolTip="Indormar se é uma compra ou venda." />
                    <asp:RadioButton ID="rdVendas" runat="server" Checked="True" GroupName="operacoes"
                        Text="Vendas" data-ToolTip="default" ToolTip="Indormar se é uma compra ou venda." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Valor:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbBruto" runat="server" Checked="True" GroupName="Valor" Text="Bruto"
                        data-ToolTip="default" ToolTip="Selecionar como o valor será apreentado no relatório." />
                    <asp:RadioButton ID="rbLiquido" runat="server" GroupName="Valor" Text="Liquido" data-ToolTip="default"
                        ToolTip="Selecionar como o valor será apreentado no relatório." />
                    &nbsp;<asp:RadioButton ID="rbBrutoMenosPisCofins" runat="server" GroupName="Valor"
                        Text="Bruto - PIS - COFINS - ISS + IPI (- ICMS Na Venda)" data-ToolTip="default"
                        ToolTip="Selecionar como o valor será apreentado no relatório." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Mercado:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdInterno" runat="server" GroupName="mercado" Text="Interno"
                        data-ToolTip="default" ToolTip="Selecionar se o mercado é interno ou externo." />&nbsp;
                    <asp:RadioButton ID="rdExterno" runat="server" GroupName="mercado" Text="Externo"
                        data-ToolTip="default" ToolTip="Selecionar se o mercado é interno ou externo." />
                    <asp:RadioButton ID="rdTodosMercados" runat="server" Checked="True" GroupName="mercado"
                        Text="Todos" data-ToolTip="default" ToolTip="Selecionar se o mercado é interno ou externo." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Curva:
                </div>
                <div class="coltxt">
                    &nbsp; A &nbsp;<asp:DropDownList ID="ddlCurvaA" runat="server">
                        <asp:ListItem>80</asp:ListItem>
                        <asp:ListItem>75</asp:ListItem>
                        <asp:ListItem>70</asp:ListItem>
                        <asp:ListItem>65</asp:ListItem>
                        <asp:ListItem Selected="True">60</asp:ListItem>
                        <asp:ListItem>55</asp:ListItem>
                        <asp:ListItem>50</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp; B&nbsp;
                    <asp:DropDownList ID="ddlCurvaB" runat="server">
                        <asp:ListItem>40</asp:ListItem>
                        <asp:ListItem>35</asp:ListItem>
                        <asp:ListItem Selected="True">30</asp:ListItem>
                        <asp:ListItem>25</asp:ListItem>
                        <asp:ListItem>20</asp:ListItem>
                        <asp:ListItem>15</asp:ListItem>
                        <asp:ListItem>10</asp:ListItem>
                        <asp:ListItem>5</asp:ListItem>
                        <asp:ListItem>0</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp; C&nbsp;
                    <asp:DropDownList ID="ddlCurvaC" runat="server">
                        <asp:ListItem>30</asp:ListItem>
                        <asp:ListItem>25</asp:ListItem>
                        <asp:ListItem>20</asp:ListItem>
                        <asp:ListItem>15</asp:ListItem>
                        <asp:ListItem Selected="True">10</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Marca:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMarca" runat="server" AutoPostBack="True" Width="250px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbReal" runat="server" Checked="True" GroupName="Moeda" Text="Real" />
                    <asp:RadioButton ID="rbDolar" runat="server" GroupName="Moeda" Text="Dolar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:SelecaoProduto ID="ucSelecaoProdutoABC" runat="server" />
</asp:Content>
