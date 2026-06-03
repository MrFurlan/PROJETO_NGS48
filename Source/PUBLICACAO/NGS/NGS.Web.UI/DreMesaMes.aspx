<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="DreMesaMes.aspx.vb" Inherits="NGS.Web.UI.DreMesaMes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadDreMesaMes() {
            $("#MainContent_lstEmpresa").multiselect();
        }

        $(document).ready(function () {
            pageLoadDreMesaMes();
            var prmDreMesaMes = Sys.WebForms.PageRequestManager.getInstance();
            prmDreMesaMes.add_endRequest(pageLoadDreMesaMes);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrm" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlDreMesaMes" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Demonstrativo da contas de resultado mês a mês
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkProcessar" Text="Processar" runat="server" />
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
                    <asp:CheckBox ID="chkConsolidado" runat="server" />Empresa:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstEmpresa" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="604px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Exercício:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlExercicio" runat="server" Style="width: 240px;" />
                </div>
                <div class="collbl">
                    Até o Mês:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMes" runat="server" Style="width: 235px;" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkPorProduto" runat="server" Text="Por Produto"
                        data-ToolTip="default" ToolTip="Filtrar com o tipo de relatóio desejado." />
                    <asp:CheckBox ID="chkInversao" runat="server" Text="Inversão de sinal"
                        data-ToolTip="default" ToolTip="Filtrar com o tipo de relatóio desejado." />
                    <asp:CheckBox ID="chkCentroDeCusto" runat="server" Text="Centro de Custo"
                        data-ToolTip="default" ToolTip="Filtrar com o tipo de relatóio desejado." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
