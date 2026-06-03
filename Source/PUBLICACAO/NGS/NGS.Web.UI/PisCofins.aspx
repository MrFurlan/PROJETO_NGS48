<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PisCofins.aspx.vb" Inherits="NGS.Web.UI.PisCofins" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngPisCofins" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPisCofins" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                PisCofins
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
                    <asp:DropDownList ID="ddlUnidade" ToolTip="Unidade de negócio empresarial." runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged1"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="634px" OnSelectedIndexChanged="DdlEmpresa_SelectedIndexChanged1"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    ICMS:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlProcesso" runat="server" Width="634px" OnSelectedIndexChanged="DdlProcesso_SelectedIndexChanged1"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    IPI:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlProcessoIPI" runat="server" Width="634px" OnSelectedIndexChanged="DdlProcesso_SelectedIndexChanged1"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="96px" />
                </div>
                <div class="collbl">
                    à:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="96px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Livro:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtLivro" CssClass="txtNumerico" runat="server" Width="64px" />
                </div>
                <div class="collbl" style="margin-left: 53px">
                    Folha:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFolha" CssClass="txtNumerico" runat="server" Width="64px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Finalidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlFinalidade" runat="server" Width="344px">
                        <asp:ListItem Value="0">0 - Remessa do Arquivo Original</asp:ListItem>
                        <asp:ListItem Value="1">1 - Remessa do Arquivo Substituto</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Perfil:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlPerfil" runat="server" Width="344px">
                        <asp:ListItem Value="A">A - Perfil A</asp:ListItem>
                        <asp:ListItem Value="B">B - Perfil B</asp:ListItem>
                        <asp:ListItem Value="C">C - Perfil C</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Atividade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlAtividade" runat="server" Width="344px">
                        <asp:ListItem Value="0">0 - Industrial ou equiparado a industria</asp:ListItem>
                        <asp:ListItem Value="1">1 - Outros</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Arq.Saida
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtArquivoDeSaida" runat="server" Width="336px" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
