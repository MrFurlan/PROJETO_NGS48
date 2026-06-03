<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ImportaSpedContabil.aspx.vb" Inherits="NGS.Web.UI.ImportaSpedContabil" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server" />
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngImportaSpedContabil" runat="server" AsyncPostBackTimeout="900" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlImportaSpedContabil" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Importa Dados para o Sped
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="591px" />
                </div>
                <div class="coltxt" style="margin-left: 124px;">
                    <asp:Button ID="cmdAjuda" runat="server" CssClass="botao" Text="Ajuda" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Plano de Contas:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPlanoDeContas" runat="server" Width="582px" data-ToolTip="default"
                        ToolTip="Selecionar arquivo de importação do plano de contas." />
                </div>
                <div class="coltxt">
                    <input style="width: 117px" id="file1" data-tooltip="default" title="..." type="file" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdPlanoDeContas" OnClick="cmdPlanoDeContas_Click" runat="server"
                        CssClass="botao" Text="Importar" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Centros de Custos:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCentroDeCusto" runat="server" Width="582px" data-ToolTip="default"
                        ToolTip="Selecionar arquivo de importação dos centros de custo." />
                </div>
                <div class="coltxt">
                    <input style="width: 117px" id="file2" type="file" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdCustos" OnClick="cmdCustos_Click" runat="server" CssClass="botao"
                        Text="Importar" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Clientes:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtClientes" runat="server" Width="582px" data-ToolTip="default"
                        ToolTip="Selecionar arquivo de importação dos clientes." />
                </div>
                <div class="coltxt">
                    <input style="width: 117px" id="file4" type="file" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdClientes" OnClick="cmdClientes_Click" runat="server" CssClass="botao"
                        Text="Importar" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Razão:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtRazao" runat="server" Width="582px" data-ToolTip="default" ToolTip="Selecionar arquivo de importação do razão." />
                </div>
                <div class="coltxt">
                    <input style="width: 117px" id="file3" type="file" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdRazao" OnClick="cmdRazao_Click" runat="server" CssClass="botao"
                        Text="Importar" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
