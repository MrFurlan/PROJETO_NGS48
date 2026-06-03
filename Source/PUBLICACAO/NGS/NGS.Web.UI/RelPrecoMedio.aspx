<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelPrecoMedio.aspx.vb" Inherits="NGS.Web.UI.RelPrecoMedio" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngRelPrecoMedio" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelPrecoMedio" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Preço Médio
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Formato Saida:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadPDF" runat="server" Checked="True" GroupName="RD" Text="PDF"
                        data-ToolTip="default" ToolTip="Formato da emissão do relatório." />
                    <asp:RadioButton ID="RadHTML" runat="server" GroupName="RD" Text="HTML" data-ToolTip="default"
                        ToolTip="Formato da emissão do relatório." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="lstEmpresa" runat="server" Width="575px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="lstGrupoProduto" runat="server" Width="575px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="lstProduto" runat="server" Width="576px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicio" runat="server" CssClass="calendario" Width="120px"
                        data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFim" runat="server" CssClass="calendario" Width="120px" data-ToolTip="default"
                        ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Desdobra Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DropOperacao" runat="server" Width="151px">
                        <asp:ListItem Value="S">SIM</asp:ListItem>
                        <asp:ListItem Selected="True" Value="N">NÃO</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="collbl">
                    Tipo Relatório:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="dpRelPor" runat="server" Width="152px">
                        <asp:ListItem>Fixa&#231;&#245;es</asp:ListItem>
                        <asp:ListItem>Movimento</asp:ListItem>
                        <asp:ListItem>Saldo</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadioReais" runat="server" Checked="True" GroupName="Moeda"
                        Text="Reais" data-ToolTip="default" ToolTip="Refere-se a moeda de negociação." />
                    <asp:RadioButton ID="RadioDolar" runat="server" GroupName="Moeda" Text="Dolar" data-ToolTip="default"
                        ToolTip="Refere-se a moeda de negociação." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
