<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PrecoMedioV20.aspx.vb" Inherits="NGS.Web.UI.PrecoMedioV20" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPrecoMedio" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPrecoMedio" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Preço Médio V 2.00
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton Text="Relatório" ID="lnkRelatorio" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton Text="HTML" ID="lnkHTML" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton Text="Limpar" ID="lnkLimpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton Text="Ajuda" ID="lnkAjuda" runat="server" />
                        </li>
                    </ul>
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
                    <asp:DropDownList ID="lstProduto" runat="server" Width="575px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicio" CssClass="calendario" runat="server" Width="100px" data-ToolTip="default" ToolTip="Período inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFim" CssClass="calendario" runat="server" Width="100px" data-ToolTip="default" ToolTip="Período final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Desdobra Oper.:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DropOperacao" runat="server" Width="126px">
                        <asp:ListItem>S</asp:ListItem>
                        <asp:ListItem Selected="True" Value="N"></asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="collbl" style="margin-left: 5px;">
                    Tipo Relatório:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="dpRelPor" runat="server" Width="126px">
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
                    <asp:RadioButton ID="RadioDolar" runat="server" GroupName="Moeda" Text="Dolar" data-ToolTip="default" ToolTip="Refere-se a moeda de negociação." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
