<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PosicaoDePedidosConsolidada.aspx.vb" Inherits="NGS.Web.UI.PosicaoDePedidosConsolidada" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 150px;
        }
    </style>
    <script type="text/javascript">
        function pageLoadPosicaoDePedidosConsolidada() {
            $("#MainContent_lstEmpresa1").multiselect({
                header: "Escolha as empresas!"
            }).multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadPosicaoDePedidosConsolidada();
            var prmPosicaoDePedidosConsolidada = Sys.WebForms.PageRequestManager.getInstance();
            prmPosicaoDePedidosConsolidada.add_endRequest(pageLoadPosicaoDePedidosConsolidada);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngESNotas" runat="server" AsyncPostBackTimeout="10000" />
    <asp:HiddenField ID="HID" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPosicaoDePedidos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Posição Consolidada de Pedidos Por Cliente
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
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstEmpresa1" runat="server" CssClass="multiselect" SelectionMode="Multiple"
                        Width="615px" data-ToolTip="default" ToolTip="Empresa Para negociação/consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkClienteConsolidado" runat="server" ToolTip="Consolidar todos os endereços do Cliente." data-ToolTip="default" />Cliente:                   
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" runat="server" Enabled="false" Width="575px" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnCliente" runat="server" Text=">" CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="615px"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Limite do Financeiro:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" runat="server" Width="100px" class="calendario" data-ToolTip="default" ToolTip="Informar a data limite do financeiro que deseja consultar." />
                </div>
                <div class="collbl">
                    Pedidos:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkPedidosComSaldo" runat="server" Checked="True" GroupName="TipoRelatorio"
                        Text="Com Saldo" data-ToolTip="default" ToolTip="Status dos pedidos para listagem." />
                </div>
                <div class="collbl">
                    Tipo de Frete:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlFretes" runat="server">
                        <asp:ListItem Value="0">TODOS</asp:ListItem>
                        <asp:ListItem Value="1">CIF</asp:ListItem>
                        <asp:ListItem Value="2">FOB</asp:ListItem>
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
