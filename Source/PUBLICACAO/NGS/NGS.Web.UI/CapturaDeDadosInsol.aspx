<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CapturaDeDadosInsol.aspx.vb" Inherits="NGS.Web.UI.CapturaDeDadosInsol" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngCapturaDeDados" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCapturaDeDados" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Captura Dados Indústria De Óleo de Soja
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkRelatorio" runat="server" Text="Processar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
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
                    Mes Inicial:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlMesInicial" runat="server" Width="120px" AutoPostBack="True">
                        <asp:ListItem Value="1">Janeiro</asp:ListItem>
                        <asp:ListItem Value="2">Fevereiro</asp:ListItem>
                        <asp:ListItem Value="3">Marco</asp:ListItem>
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
                <div class="collbl">
                    Mes Final:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlMesFinal" runat="server" Width="120px">
                        <asp:ListItem Value="1">Janeiro</asp:ListItem>
                        <asp:ListItem Value="2">Fevereiro</asp:ListItem>
                        <asp:ListItem Value="3">Marco</asp:ListItem>
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
                <div class="collbl">
                    Dia:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDia" runat="server" Text="0" Width="29px" data-ToolTip="default"
                        ToolTip="Informar o dia do mês." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Exercicio:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlAno" runat="server" Width="120px" AutoPostBack="True" />
                </div>
                <div class="collbl">
                    Preço:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlPreco" runat="server" Width="120px">
                        <asp:ListItem>Avaliado</asp:ListItem>
                        <asp:ListItem>Real</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Mes Ciclos:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCiclos" runat="server" Width="27px" Text="5" data-ToolTip="default"
                        ToolTip="Informar o mês e selecionar uma das opções ao lado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Processos:
                </div>
                <div class="coltxt">
                    <asp:CheckBoxList ID="ChkResultado" runat="server" Enabled="False" Font-Bold="True"
                        data-ToolTip="default" ToolTip="Informações referente a execução do processo de captura." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
