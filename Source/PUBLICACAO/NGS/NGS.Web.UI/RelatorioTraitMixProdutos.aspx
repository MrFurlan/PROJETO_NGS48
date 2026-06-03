<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioTraitMixProdutos.aspx.vb" Inherits="NGS.Web.UI.RelatorioTraitMixProdutos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadTraitMix() {
            $("#MainContent_lstEncargos").multiselect({
                header: "Escolha apenas 5 encargos!",
                selectedList: 5,
                click: function (e) {
                    if ($(this).multiselect("widget").find("input:checked").length > 5) {
                        return false;
                    }
                }
            }).multiselectfilter();

            $("#MainContent_lstClasseOp").multiselect().multiselectfilter();

        }

        $(document).ready(function () {
            pageLoadTraitMix();
            var prmRelatoriotraitmix = Sys.WebForms.PageRequestManager.getInstance();
            prmRelatoriotraitmix.add_endRequest(pageLoadTraitMix);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngEstados" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlTrait" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Relatório Trait Mix de Produtos
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
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." runat="server" Text='Empresa:' />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" AutoPostBack="True" Width="595px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Marca do Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMarcaProd" runat="server" Width="595px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="595px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe Operacao:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstClasseOp" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="370px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkPerido" runat="server" AutoPostBack="True" Text="Usar Período" />
                </div>
                <asp:Panel ID="pnlPeriodo" runat="server" Visible="False">
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataDe" runat="server" CssClass="calendario" data-ToolTip="default" ToolTip="Marcar para inserir o período." />
                        <asp:TextBox ID="txtDataAte" runat="server" CssClass="calendario" data-ToolTip="default" ToolTip="Marcar para inserir o período." />
                    </div>
                </asp:Panel>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbValor" AutoPostBack="True" Checked="true" GroupName="gpQtdeValor"
                        Text="Valor" runat="server" data-ToolTip="default" ToolTip="Marcar para filtrar por valor ou quantidade." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbQtde" runat="server" AutoPostBack="True" GroupName="gpQtdeValor"
                        Text="Quantidade" data-ToolTip="default" ToolTip="Marcar para filtrar por valor ou quantidade." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rbNota" Checked="true" GroupName="gpNotaEmb" Visible="false"
                        Text="Nota" runat="server" data-ToolTip="default" ToolTip="Marcar para filtrar por valor ou quantidade." />
                    <asp:RadioButton ID="rbEmbalagem" GroupName="gpNotaEmb" Visible="false" Text="Embalagem"
                        runat="server" data-ToolTip="default" ToolTip="Marcar para filtrar por valor ou quantidade." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
