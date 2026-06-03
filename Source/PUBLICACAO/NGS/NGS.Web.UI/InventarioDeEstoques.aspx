<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="InventarioDeEstoques.aspx.vb" Inherits="NGS.Web.UI.InventarioDeEstoques" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngInventarioDeEstoques" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlInventarioDeEstoques" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Inventário de Estoque
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
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcelDados" runat="server" Text="Excel Dados" />
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="600px" OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" />
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
                    Data:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlMes" runat="server" Width="104px" OnSelectedIndexChanged="DdlMes_SelectedIndexChanged"
                        AutoPostBack="True">
                        <asp:ListItem Value="1">Janeiro</asp:ListItem>
                        <asp:ListItem Value="2">Fevereiro</asp:ListItem>
                        <asp:ListItem Value="3">Mar&#231;o</asp:ListItem>
                        <asp:ListItem Value="4">Abril</asp:ListItem>
                        <asp:ListItem Value="5">Maio</asp:ListItem>
                        <asp:ListItem Value="6">Junho</asp:ListItem>
                        <asp:ListItem Value="7">Julho</asp:ListItem>
                        <asp:ListItem Value="8">Agosto</asp:ListItem>
                        <asp:ListItem Value="9">Setembro</asp:ListItem>
                        <asp:ListItem Value="10">Outubro</asp:ListItem>
                        <asp:ListItem Value="11">Novembro</asp:ListItem>
                        <asp:ListItem Value="12">Dezembro</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp; / &nbsp;
                    <asp:DropDownList ID="DdlAno" runat="server" Width="120px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadReal" runat="server" Text="Real" Checked="True" GroupName="Moeda"
                        data-ToolTip="default" ToolTip="Refere-se a moeda de negociação." />
                    <asp:RadioButton ID="RadDolar" runat="server" Text="Dolar" GroupName="Moeda" data-ToolTip="default"
                        ToolTip="Refere-se a moeda de negociação." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Valor:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadValorReal" runat="server" Text="Real" Checked="True" GroupName="Valor"
                        data-ToolTip="default" ToolTip="Selecionar se o valor é real ou arbitrado." />
                    <asp:RadioButton ID="RadValorAvaliado" runat="server" Text="Avaliado" GroupName="Valor"
                        data-ToolTip="default" ToolTip="Selecionar se o valor é real ou arbitrado." />
                    <asp:TextBox ID="txtPorcentagem" runat="server" Width="48px" data-ToolTip="default"
                        ToolTip="Selecionar se o valor é real ou arbitrado." />%
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Relatório:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadPorProduto" runat="server" Text="Por Produto" Checked="True"
                        GroupName="Relatorio" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório por produto, filial ou todas as filiais." />
                    <asp:RadioButton ID="RadPorFilial" runat="server" Text="Por Filial" GroupName="Relatorio"
                        data-ToolTip="default" ToolTip="Selecionar o tipo de relatório por produto, filial ou todas as filiais." />
                    <asp:CheckBox ID="ChkTodasAsFiliais" runat="server" Text="Todas as Filiais" data-ToolTip="default"
                        ToolTip="Selecionar o tipo de relatório por produto, filial ou todas as filiais." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
