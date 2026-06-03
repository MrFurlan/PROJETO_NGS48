<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RegistrosFiscaisIN86.aspx.vb" Inherits="NGS.Web.UI.RegistrosFiscaisIN86" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .ident {
            margin-left: 125px;
        }

        .right {
            float: right;
            margin-right: 190px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRegistrosFiscaisIN86" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRegistrosFiscaisIN86" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Registros Fiscais IN86 - SVA
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
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
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Ano:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlAno" runat="server">
                        <asp:ListItem>2009</asp:ListItem>
                        <asp:ListItem>2010</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="collbl">
                    Mes Inicial:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlMesInicial" runat="server">
                        <asp:ListItem Value="1">Jan</asp:ListItem>
                        <asp:ListItem Value="2">Fev</asp:ListItem>
                        <asp:ListItem Value="3">Mar</asp:ListItem>
                        <asp:ListItem Value="4">Abr</asp:ListItem>
                        <asp:ListItem Value="5">Mai</asp:ListItem>
                        <asp:ListItem Value="6">Jun</asp:ListItem>
                        <asp:ListItem Value="7">Jul</asp:ListItem>
                        <asp:ListItem Value="8">Ago</asp:ListItem>
                        <asp:ListItem Value="9">Set</asp:ListItem>
                        <asp:ListItem Value="10">Out</asp:ListItem>
                        <asp:ListItem Value="11">Nov</asp:ListItem>
                        <asp:ListItem Value="12">Dez</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="collbl">
                    Mes Final:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlMesFinal" runat="server">
                        <asp:ListItem Value="1">Jan</asp:ListItem>
                        <asp:ListItem Value="2">Fev</asp:ListItem>
                        <asp:ListItem Value="3">Mar</asp:ListItem>
                        <asp:ListItem Value="4">Abr</asp:ListItem>
                        <asp:ListItem Value="5">Mai</asp:ListItem>
                        <asp:ListItem Value="6">Jun</asp:ListItem>
                        <asp:ListItem Value="7">Jul</asp:ListItem>
                        <asp:ListItem Value="8">Ago</asp:ListItem>
                        <asp:ListItem Value="9">Set</asp:ListItem>
                        <asp:ListItem Value="10">Out</asp:ListItem>
                        <asp:ListItem Value="11">Nov</asp:ListItem>
                        <asp:ListItem Value="12">Dez</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtDataIni" runat="server" CssClass="calendario" Width="80px" data-ToolTip="default" ToolTip="Período inicial da apuração." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtDataFin" runat="server" CssClass="calendario" Width="80px" data-ToolTip="default" ToolTip="Período final da apuração." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Parâmetros:
                </div>
                <div class="coltxt right">
                    <asp:Button ID="Button1" OnClick="Button1_Click" runat="server" UseSubmitBehavior="False"
                        CssClass="botao" Text="Processar" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CheckBox1" runat="server" Enabled="False" Text="431 - Arquivo Mestre de Mercadorias e Servicos" data-ToolTip="default" ToolTip="Selecionar os parâmetros desejados e processar." />
                </div>
            </div>
            <div class="row ident">
                <div class="coltxt  right">
                    <asp:Button ID="Button2" CssClass="botao" OnClick="Button2_Click2" runat="server" UseSubmitBehavior="False"
                        Text="Processar" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CheckBox2" runat="server" Enabled="False" Text="432 - Arquivo de Itens de Mercadorias e Servicos" data-ToolTip="default" ToolTip="Selecionar os parâmetros desejados e processar." />
                </div>
            </div>
            <div class="row ident">
                <div class="coltxt  right">
                    <asp:Button ID="Button3" CssClass="botao" OnClick="Button3_Click1" runat="server" UseSubmitBehavior="False"
                        Text="Processar" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CheckBox3" runat="server" Text="433 - Arquivo Mestre de Mercadorias e Servicos (Entradas) - Emitidas por Terceiros"
                        Enabled="False" data-ToolTip="default" ToolTip="Selecionar os parâmetros desejados e processar." />
                </div>
            </div>
            <div class="row ident">
                <div class="coltxt  right">
                    <asp:Button ID="Button4" CssClass="botao" OnClick="Button4_Click1" runat="server" UseSubmitBehavior="False"
                        Text="Processar" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CheckBox4" runat="server" Text="434 - Arquivo de Itens de Mercadorias e Servicos(Entradas) - Emitidas por Terceiros"
                        Enabled="False" data-ToolTip="default" ToolTip="Selecionar os parâmetros desejados e processar." />
                </div>
            </div>
            <div class="row ident">
                <div class="coltxt  right">
                    <asp:Button ID="Button5" CssClass="botao" OnClick="Button5_Click1" runat="server" Enabled="False"
                        UseSubmitBehavior="False" Text="Processar" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CheckBox5" runat="server" Text="461 - Arquivo de Insumos Relacionados"
                        Enabled="False" data-ToolTip="default" ToolTip="Selecionar os parâmetros desejados e processar." />
                </div>
            </div>
            <div class="row ident">
                <div class="coltxt  right">
                    <asp:Button ID="Button6" CssClass="botao" OnClick="Button6_Click1" runat="server" UseSubmitBehavior="False"
                        Text="Processar" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CheckBox6" Text="491 - Arquivo de Cadastro de Pessoas Juridicas e Fisicas"
                        runat="server" Enabled="False" data-ToolTip="default" ToolTip="Selecionar os parâmetros desejados e processar." />
                </div>
            </div>
            <div class="row ident">
                <div class="coltxt  right">
                    <asp:Button ID="Button7" CssClass="botao" OnClick="Button7_Click1" runat="server" UseSubmitBehavior="False"
                        Text="Processar" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CheckBox7" runat="server" Text="494 - Tabela de Natureza da Operacao"
                        Enabled="False" data-ToolTip="default" ToolTip="Selecionar os parâmetros desejados e processar." />
                </div>
            </div>
            <div class="row ident">
                <div class="coltxt  right">
                    <asp:Button ID="Button8" CssClass="botao" OnClick="Button8_Click1" runat="server" UseSubmitBehavior="False"
                        Text="Processar" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CheckBox8" Text="495 - Tabela de Mercadorias e servicos" runat="server"
                        Enabled="False" data-ToolTip="default" ToolTip="Selecionar os parâmetros desejados e processar." />
                </div>
            </div>
            <div class="row ident">
                <div class="coltxt  right">
                    <asp:Button ID="Button9" CssClass="botao" OnClick="Button9_Click1" runat="server" UseSubmitBehavior="False"
                        Text="Processar" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CheckBox9" Text="4101- Arquivo Complementar de Registro de Saída de Merc/Serviços, Emitidas pela PJ"
                        runat="server" Enabled="False" data-ToolTip="default" ToolTip="Selecionar os parâmetros desejados e processar." />
                </div>
            </div>
            <div class="row ident">
                <div class="coltxt  right">
                    <asp:Button ID="Button10" CssClass="botao" OnClick="Button10_Click1" runat="server"
                        UseSubmitBehavior="False" Text="Processar" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CheckBox10" Text="4104- Arquivo Complementar de Registro de Entrada de Merc/Serviços, Emitidas pela PJ"
                        runat="server" Enabled="False" data-ToolTip="default" ToolTip="Selecionar os parâmetros desejados e processar." />
                </div>
            </div>
            <div class="row ident">
                <div class="coltxt  right">
                    <asp:Button ID="Button11" CssClass="botao" OnClick="Button11_Click1" runat="server"
                        UseSubmitBehavior="False" Text="Processar" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="CheckBox11" Text="4105 - Arquivo Complementar de Registro de Entrada de Merc/Serviços (Terceiros)"
                        runat="server" Enabled="False" data-ToolTip="default" ToolTip="Selecionar os parâmetros desejados e processar." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
