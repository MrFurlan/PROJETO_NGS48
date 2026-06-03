<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeFretesExcel.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeFretesExcel" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrm" runat="server" AsyncPostBackTimeout="10000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelacaoDeNotasPisCofins" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Relação de Fretes
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
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCliente" runat="server" Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupo" runat="server" Width="634px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProduto" runat="server" Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdoTodosFretes" runat="server" Text="Todos os Fretes" GroupName="Filtros"
                        Checked="true" />
                    <asp:RadioButton ID="rdoFretesSemNota" runat="server" Text="Fretes sem Nota" GroupName="Filtros" />
                    <asp:RadioButton ID="rdoNotasSemFrete" runat="server" Text="Notas Sem Frete" GroupName="Filtros" />
                    <asp:RadioButton ID="rdoResumoCustos" runat="server" Text="Resumo dos Custos" GroupName="Filtros" />
                    <asp:CheckBox ID="chkDataNota" runat="server" Text="Data da Nota" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" Width="96px" CssClass="calendario" />
                    &nbsp;à&nbsp;
                    <asp:TextBox ID="txtDataFinal" runat="server" Width="96px" CssClass="calendario" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
