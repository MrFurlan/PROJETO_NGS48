<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioMovimentacoesFinanceiras.aspx.vb" Inherits="NGS.Web.UI.RelatorioMovimentacoesFinanceiras" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="smcgRelatorioMovimentacoesFinanceiras" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <asp:HiddenField ID="HID" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioMovimentacoesFinanceiras" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Relatório Movimentações Financeiras
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
                    <asp:DropDownList ID="DdlUnidadeConsultaTitulos" TabIndex="1" runat="server" Width="594px"
                        AutoPostBack="True" OnSelectedIndexChanged="DdlUnidadeConsultaTitulos_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresaConsultaTitulos" TabIndex="2" runat="server" Width="594px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNomeCliente" runat="server" Width="553px" Enabled="false" />
                    <asp:HiddenField ID="CodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Carteira:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlCarteiras" TabIndex="4" runat="server" Width="594px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" CssClass="calendario" TabIndex="5"
                        runat="server" Width="100px" data-ToolTip="default" ToolTip="Data inicial da pesquisa." />
                </div>
                <div class="collbl">
                    Período Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" CssClass="calendario" TabIndex="6"
                        runat="server" Width="100px" CausesValidation="True" data-ToolTip="default" ToolTip="Data final da pesquisa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadAtivos" TabIndex="7" runat="server" Text="Ativos   " GroupName="Selecao"
                        data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:RadioButton ID="RadBaixados" TabIndex="8" runat="server" Text="Baixados" GroupName="Selecao"
                        data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:RadioButton ID="RadTodos" TabIndex="9" runat="server" Text="Todos" GroupName="Selecao"
                        data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                    <asp:RadioButton ID="RadDebitosECreditos" runat="server" GroupName="Selecao" TabIndex="10"
                        Text="Débitos e Créditos" data-ToolTip="default" ToolTip="Selecionar o tipo de relatório desejado a partir das opções." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
