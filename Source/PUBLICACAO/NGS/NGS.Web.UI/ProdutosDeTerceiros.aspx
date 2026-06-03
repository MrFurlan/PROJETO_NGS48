<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ProdutosDeTerceiros.aspx.vb" Inherits="NGS.Web.UI.ProdutosDeTerceiros" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngProdutosDeTerceiros" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlProdutosDeTerceiros" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Produtos De Terceiros
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
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
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlGrupo" runat="server" Width="600px" OnSelectedIndexChanged="DdlGrupo_SelectedIndexChanged1"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlProduto" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlMes" runat="server" Width="120px" OnSelectedIndexChanged="DdlMes_SelectedIndexChanged"
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
                </div>
                <div class="coltxt">
                    &nbsp; / &nbsp;
                    <asp:DropDownList ID="DdlAno" runat="server" Width="120px" OnSelectedIndexChanged="DdlAno_SelectedIndexChanged"
                        AutoPostBack="True">
                    </asp:DropDownList>
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
