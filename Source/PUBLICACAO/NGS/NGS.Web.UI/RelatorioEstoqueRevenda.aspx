<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioEstoqueRevenda.aspx.vb" Inherits="NGS.Web.UI.RelatorioEstoqueRevenda" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadRelatorioEstoqueRevenda() {
            $("#<%=lstEmpresa.ClientID%>").multiselect({
                header: "Escolha a(s) empresa(s)!"
            }).multiselectfilter();
            //debugger;
            if ($("#<%=chkMapa.ClientID%>").attr('checked') == "checked") {
                $("#pnlColunas").show("drop");
            }
            else {
                $("#pnlColunas").hide("drop");
            }

            $("#<%=chkMapa.ClientID%>").click(function () {
                var checked = $(this).attr('checked') == "checked";
                if (checked) {
                    $("#pnlColunas").show("drop");
                }
                else {
                    $("#pnlColunas").hide("drop");
                }
            });
        }

        $(document).ready(function () {
            pageLoadRelatorioEstoqueRevenda();
            var prmRelatorioEstoqueRevenda = Sys.WebForms.PageRequestManager.getInstance();
            prmRelatorioEstoqueRevenda.add_endRequest(pageLoadRelatorioEstoqueRevenda);
        });
    </script>
    <style type="text/css">
        .collbl {
            width: 130px;
        }

        .painelleft {
            margin-right: 37px;
        }
    </style>

</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioEstoqueRevenda" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioEstoqueRevenda" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Relatório de Estoque/Preços
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
                    Tipo Relatório:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RdEstoqueComPreco" runat="server" GroupName="Tipo" Text="Estoque/Preço"
                        AutoPostBack="True" data-ToolTip="default" ToolTip="Selecionar de acordo com o relatório desejado." />
                    <asp:RadioButton ID="rbEstoque" runat="server" Checked="True" GroupName="Tipo" Text="Estoques"
                        AutoPostBack="True" data-ToolTip="default" ToolTip="Selecionar de acordo com o relatório desejado." />
                    <asp:RadioButton ID="rbPreco" runat="server" GroupName="Tipo" Text="Avaliação do Estoque"
                        AutoPostBack="True" data-ToolTip="default" ToolTip="Selecionar de acordo com o relatório desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstEmpresa" runat="server" CssClass="multiselect" SelectionMode="Multiple" Width="648px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período de:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" CssClass="calendario" data-ToolTip="default" ToolTip="Data Inicial a final da consulta." />
                    &nbsp;&nbsp;à&nbsp;
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" data-ToolTip="default" ToolTip="Data Inicial a final da consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Seleções:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkLoteSemente" runat="server" Text="Agrupar Lotes" Checked="True" data-ToolTip="default" ToolTip="Selecionar de acordo com as informações desejadas." />
                    <asp:CheckBox ID="chkEmbalagens" runat="server" Text="Agrupar Embalagens" Checked="True" data-ToolTip="default" ToolTip="Selecionar de acordo com as informações desejadas." />
                    <asp:CheckBox ID="chkSaldoQtde" runat="server" Text="Com Saldo de Quantidade" data-ToolTip="default" ToolTip="Selecionar de acordo com as informações desejadas." />
                    <asp:CheckBox ID="chkVendaCompraSaldo" runat="server" Text="Com Pedidos de Compra/Venda em Aberto" data-ToolTip="default" ToolTip="Selecionar de acordo com as informações desejadas." />
                    <asp:CheckBox ID="ckPorSafra" runat="server" Text="Lista Por Safra" AutoPostBack="True" data-ToolTip="default" ToolTip="Selecionar de acordo com as informações desejadas." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkMapa" runat="server" Text="Mapa Comparativo" />
                </div>
                <div id="pnlColunas" class="coltxt" style="display: none;">
                    <div class="painelleft">
                        <asp:CheckBox ID="chkSaldoAnterior" runat="server" Text="Saldo Anterior" data-ToolTip="default" ToolTip="" />
                        <br />
                        <asp:CheckBox ID="ChkConsignado" runat="server" Text="Em Nosso Poder Consignado" data-ToolTip="default" ToolTip="" />
                        <br />
                        <asp:CheckBox ID="ChkSaldoProprio" runat="server" Text="Saldo Proprio" data-ToolTip="default" ToolTip="" />
                        <br />
                        <asp:CheckBox ID="ChkCompra" runat="server" Text="Pedido De Compra Não Recebido" data-ToolTip="default" ToolTip="" />
                    </div>
                    <div class="painelleft">
                        <asp:CheckBox ID="ChkEntradas" runat="server" Text="Entradas No Periodo" data-ToolTip="default" ToolTip="" />
                        <br />
                        <asp:CheckBox ID="ChkAmostraGratis" runat="server" Text="Amostra Gratis" data-ToolTip="default" ToolTip="" />
                        <br />
                        <asp:CheckBox ID="ChkSaldoFora" runat="server" Text="Produto Nosso em Poder de Terceiros" data-ToolTip="default" ToolTip="" />
                        <br />
                        <asp:CheckBox ID="ChkVenda" runat="server" Text="Pedido de Venda Não Entregue" data-ToolTip="default" ToolTip="" />
                    </div>
                    <div class="painelleft">
                        <asp:CheckBox ID="ChkSaidas" runat="server" Text="Saidas No Periodo" data-ToolTip="default" ToolTip="" />
                        <br />
                        <asp:CheckBox ID="ChkTerceiros" runat="server" Text="Produto de Terceiros em Nosso Poder" data-ToolTip="default" ToolTip="" />
                        <br />
                        <asp:CheckBox ID="ChkFatNaoEntregue" runat="server" Text="Faturado Não Entregue" data-ToolTip="default" ToolTip="" />
                        <br />
                        <asp:CheckBox ID="ChkPrevisao" runat="server" Text="Previsão Estoque Final" data-ToolTip="default" ToolTip="" />
                    </div>
                    <div>
                    </div>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra Final:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="649px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Marca:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMarca" runat="server" AutoPostBack="True" Width="650px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
