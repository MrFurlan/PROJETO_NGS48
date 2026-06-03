<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="BalanceteAuxiliarVBA.aspx.vb" Inherits="NGS.Web.UI.BalanceteAuxiliarVBA" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngBalancetes" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlBalancetes" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Balancetes
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
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
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" ToolTip="Unidade de negócio empresarial." runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged"
                        Width="672px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." runat="server" Text="Empresa:" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="672px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Periodo:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" Width="88px"
                        data-ToolTip="default" ToolTip="Data inicial do relatório." />
                </div>
                <div class="collbl">
                    à
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="88px"
                        data-ToolTip="default" ToolTip="Data final do relatório." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdAuxiliar" runat="server" Text="Auxiliar" Checked="true" GroupName="filtro"
                        data-ToolTip="default" ToolTip="Filtrar com o tipo de relatóio desejado." />
                    <asp:RadioButton ID="rdGerencial" runat="server" Text="Gerencial" GroupName="filtro"
                        data-ToolTip="default" ToolTip="Filtrar com o tipo de relatóio desejado." />
                    <asp:RadioButton ID="rdAdiantamentos" runat="server" Text="Adiantamentos" GroupName="filtro"
                        data-ToolTip="default" ToolTip="Filtrar com o tipo de relatóio desejado." />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="margin-left: 129px;">
                    <asp:CheckBox ID="chkCompensacao" runat="server" Checked="true" Text="Isolar Grupos de Compensação"
                        data-ToolTip="default" ToolTip="" />
                    <asp:CheckBox ID="chkClientes" runat="server" Checked="true" Text="Desdobrar Grupos de Clientes"
                        data-ToolTip="default" ToolTip="" />
                    <asp:CheckBox ID="chkTitulos" runat="server" Text="Executar Comparativos de Títulos"
                        data-ToolTip="default" ToolTip="" />
                    <asp:CheckBox ID="chkPlanilha" runat="server" Text="Geral Planilha Analítica"
                        data-ToolTip="default" ToolTip="" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
