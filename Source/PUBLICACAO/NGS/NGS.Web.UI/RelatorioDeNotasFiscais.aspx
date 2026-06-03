<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioDeNotasFiscais.aspx.vb" Inherits="NGS.Web.UI.RelatorioDeNotasFiscais" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioDeNotasFiscais" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioDeNotasFiscais" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Relatório De Notas Fiscais
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
                            <asp:LinkButton ID="lnkAjudar" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" ToolTip="Unidade de negócio empresarial." runat="server" Width="470px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="470px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlDeposito" runat="server" Width="470px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Dep. Destino:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlDepositoDestino" runat="server" Width="470px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCliente" runat="server" Width="470px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupo" runat="server" Width="470px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProduto" runat="server" Width="470px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlOperacao" runat="server" Width="470px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    SubOperação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSubOperacao" runat="server" Width="470px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" Width="80px" runat="server" />&nbsp;&nbsp;à&nbsp;
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" Width="80px" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Filtros:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkNotasAgrupadas" runat="server" Text="Notas Agrupadas" />
                    <asp:CheckBox ID="chkIncluiNotasProdutosDiversos" runat="server" Text="Incluir Notas De Produtos Diversos" />
                    <asp:CheckBox ID="chkObservacoesFiscais" runat="server" Text="Incluir Observações Fiscais" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="margin-left: 123px;">
                    <asp:RadioButton ID="rdoTodasNotas" Checked="true" runat="server" GroupName="Notas"
                        Text="Todas as Notas" />
                    <asp:RadioButton ID="rdoNotasSemRomaneio" runat="server" GroupName="Notas" Text="Notas Sem Romaneios" />
                    <asp:RadioButton ID="rdoLaudoSemNota" runat="server" GroupName="Notas" Text="Laudo Sem Nota" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
