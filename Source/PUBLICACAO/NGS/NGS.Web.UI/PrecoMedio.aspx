<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PrecoMedio.aspx.vb" Inherits="NGS.Web.UI.PrecoMedio" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 140px;
        }
    </style>

    <script type="text/javascript">
        function pageLoadPrecoMedio() {
            inicializar();
            $("#<%=chkTroca.ClientID%>").click(function () {
                inicializar();
            });


            function inicializar() {
                var checked = $("#<%=chkTroca.ClientID%>").attr('checked') == "checked";
                if (checked) {
                    $("#<%=ChkTrocaFrete.ClientID%>").parent().show();
                }
                else {
                    $("#<%=ChkTrocaFrete.ClientID%>").attr('checked', checked);
                    $("#<%=ChkTrocaFrete.ClientID%>").parent().hide();
                }
            };
        }

        $(document).ready(function () {
            pageLoadPrecoMedio();
            var prmPrecoMedio = Sys.WebForms.PageRequestManager.getInstance();
            prmPrecoMedio.add_endRequest(pageLoadPrecoMedio);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngESNotas" runat="server" AsyncPostBackTimeout="5000" />
    <asp:HiddenField ID="HID" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPosicaoDePedidos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Preço Medio
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
                    <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeNegocio_SelectedIndexChanged"
                        Width="618px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." Text="Consolidar Empresa:" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="618px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkConsolidarCliente" ToolTip="Consolidar o cpf/cnpj do cliente." runat="server" Text="Consolidar Cliente:" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" TabIndex="3" runat="server" Width="585px" Font-Names="monospace"
                        Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaClientes" TabIndex="4" OnClick="cmdConsultaClientes_Click"
                        CssClass="btn" runat="server" Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Unificar as informações por cliente." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbSafra" runat="server" Width="615px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo de Pedido:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ChkNormal" runat="server" Text="Normal" Checked="true" data-ToolTip="default" ToolTip="Selecionar uma das opções (troca, compra/venda antecipada, recompra) de acordo com a particularidade do pedido." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkAntecipada" runat="server" Text="Compra/Venda Antecipada" Checked="false" data-ToolTip="default" ToolTip="Selecionar uma das opções (troca, compra/venda antecipada, recompra) de acordo com a particularidade do pedido." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkRecompra" runat="server" Text="Recompra" Checked="false" data-ToolTip="default" ToolTip="Selecionar uma das opções (troca, compra/venda antecipada, recompra) de acordo com a particularidade do pedido." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkTroca" runat="server" Text="Troca" Checked="false" data-ToolTip="default" ToolTip="Selecionar uma das opções (troca, compra/venda antecipada, recompra) de acordo com a particularidade do pedido." />
                </div>
                <div class="coltxt" runat="server">
                    <asp:CheckBox ID="ChkTrocaFrete" runat="server" Text="Consolidar Frete CIF/FOB" Visible="true" data-ToolTip="default" ToolTip="Selecionar uma das opções (troca, compra/venda antecipada, recompra) de acordo com a particularidade do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Posição no dia:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" runat="server" CssClass="calendario" Width="96px" data-ToolTip="default" ToolTip="Selecionar a data desejada para consulta." />
                </div>
                <div class="collbl" style="width: 113px;">
                    Moeda:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMoeda" runat="server" Width="125px" />
                </div>
                <div class="collbl" style="width: 113px;">
                    Tipo Relatorio:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdAnalitico" runat="server" GroupName="GR" Text="Analitico" data-ToolTip="default" ToolTip="Selecionar se o relatório será analítico ou sintético." />
                    <asp:RadioButton ID="rdSintetico" runat="server" Checked="True" GroupName="GR" Text="Sintetico" data-ToolTip="default" ToolTip="Selecionar se o relatório será analítico ou sintético." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">Frete:</div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlFrete" runat="server" Width="130px">
                        <asp:ListItem>Todos</asp:ListItem>
                        <asp:ListItem Value="NEN">Nenhum</asp:ListItem>
                        <asp:ListItem>CIF</asp:ListItem>
                        <asp:ListItem>FOB</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 100%;">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
