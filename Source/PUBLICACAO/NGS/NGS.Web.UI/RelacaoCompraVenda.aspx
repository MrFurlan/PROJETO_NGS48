<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelacaoCompraVenda.aspx.vb" Inherits="NGS.Web.UI.RelacaoCompraVenda" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmgRelacaoCompraVenda" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelacaoCompraVenda" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Relação De Compra e Venda
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li class="iconPdf">
                                    <asp:LinkButton ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li class="iconExcel">
                                    <asp:LinkButton ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
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
                    <asp:DropDownList ID="ddlUnidadeDeNegocio" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" data-ToolTip="default" ToolTip="Consolidar o cnpj da empresa." runat="server" Text="Empresa:" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="600px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkMovimento" runat="server" Text="Movimento:" AutoPostBack="True" />
                </div>
                <div runat="server">
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataDe" runat="server" Width="110px" CssClass="calendario" data-ToolTip="default" ToolTip="Data da movimentação." />
                    </div>
                    <div class="collbl">
                        À:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataAte" runat="server" Width="110px" CssClass="calendario" data-ToolTip="default" ToolTip="Data da movimentação." />
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedidos com Preço:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkPrecoFixo" runat="server" Checked="True" Text="Fixo" data-ToolTip="default" ToolTip="Selecionar se o preço é fixo ou a fixar." />
                    <asp:CheckBox ID="chkPrecoAFixar" runat="server" Text="A Fixar" data-ToolTip="default" ToolTip="Selecionar se o preço é fixo ou a fixar." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo Lançamento:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkNormal" runat="server" Text="Normal" Checked="True" data-ToolTip="default" ToolTip="Marcar se o pedido foi normal, complemento, estorno ou fixação." />
                    <asp:CheckBox ID="chkComplemento" runat="server" Text="Complemento" Checked="True" data-ToolTip="default" ToolTip="Marcar se o pedido foi normal, complemento, estorno ou fixação." />
                    <asp:CheckBox ID="chkEstorno" runat="server" Text="Estorno" Checked="True" data-ToolTip="default" ToolTip="Marcar se o pedido foi normal, complemento, estorno ou fixação." />
                    <asp:CheckBox ID="chkFixacao" runat="server" Text="Fixação" Checked="True" data-ToolTip="default" ToolTip="Marcar se o pedido foi normal, complemento, estorno ou fixação." />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 100%;">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
