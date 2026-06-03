<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ComprasVendasAFixar.aspx.vb" Inherits="NGS.Web.UI.ComprasVendasAFixar" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server" />
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="ScripscmngComprasVendasAFixar" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlComprasVendasAFixar" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Compras e Vendas à Fixar
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server" class="iconRelatorio">
                            <asp:LinkButton ID="lnkRelatorio" runat="server" Text="Relatório" />
                        </li>
                        <li runat="server" class="iconLimpar">
                            <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server" class="iconAjuda">
                            <asp:LinkButton ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="600px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
            </div>
            <div class="row">
                <div class="collbl">
                    Até a Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="TxtData" runat="server" CssClass="calendario" data-ToolTip="default"
                        ToolTip="Selecionar a data final desejada." />
                </div>
                <div class="collbl">
                    Selecione:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RdCompras" runat="server" Checked="True" Text="Compras" GroupName="Afixar"
                        data-ToolTip="default" ToolTip="Selecionar se é comrpa, venda ou zerados." />
                    &nbsp;
                    <asp:RadioButton ID="RdVendas" runat="server" Text="Vendas" GroupName="Afixar" data-ToolTip="default"
                        ToolTip="Selecionar se é comrpa, venda ou zerados." />
                    &nbsp;
                    <asp:CheckBox ID="ChkZerados" runat="server" Text="Zerados" data-ToolTip="default"
                        ToolTip="Selecionar se é comrpa, venda ou zerados." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
