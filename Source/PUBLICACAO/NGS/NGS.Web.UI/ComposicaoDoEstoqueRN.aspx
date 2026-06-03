<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ComposicaoDoEstoqueRN.aspx.vb" Inherits="NGS.Web.UI.ComposicaoDoEstoqueRN" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngComposicaoDoEstoque" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlComposicaoDoEstoque" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Composição Diária do Estoque Romaneios/Nota Fiscal
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" runat="server" Text="Relatório"
                                 />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar"  />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda"  />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" ToolTip="Unidade de negócio empresarial." runat="server" Width="600px" AutoPostBack="True"
                        OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Depósito:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlDeposito" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
            </div>
            <div class="row">
                <div class="collbl">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdAtivo" runat="server" Text="Apenas Produtos Ativos" Checked="True" GroupName="Situacao" data-ToolTip="default" ToolTip="Listar apenas os Produtos com Situação Normal." />
                    <asp:RadioButton ID="rdInativo" runat="server" Text="Apenas Produto(s) Inativo(s)/Bloqueado(s)" GroupName="Situacao" data-ToolTip="default" ToolTip="Listar apenas os Produtos com a Situação diferente da Normal." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Estoque:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdFisico" runat="server" Text="Fisico" GroupName="Estoque" Checked="True" />
                    <asp:RadioButton ID="rdFiscal" runat="server" Text="Fiscal" GroupName="Estoque" />
                    <asp:CheckBox ID="chkContraPartida" runat="server" Text="Gerar Contra Partida dos Depositos de Terceiros" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
