<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PosicaoFinanceiraDeVendas.aspx.vb" Inherits="NGS.Web.UI.PosicaoFinanceiraDeVendas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 150px;
        }
    </style>
    <script type="text/javascript">
        function pageLoadPosicaoFinanceiraDeVendas() {
            $("#MainContent_lstEmpresa1").multiselect({
                header: "Escolha as empresas!"
            }).multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadPosicaoFinanceiraDeVendas();
            var prmPosicaoFinanceiraDeVendas = Sys.WebForms.PageRequestManager.getInstance();
            prmPosicaoFinanceiraDeVendas.add_endRequest(pageLoadPosicaoFinanceiraDeVendas());
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPosicaoFinanceira" runat="server" AsyncPostBackTimeout="10000" />
    <asp:HiddenField ID="HID" runat="server" />
    <orea:ajaxupdating id="ajaxUpdating" runat="server" text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPosicaoFinanceira" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Posição Financeira de Vendas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkExcel" Text="Relatório" runat="server" />
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
                    <asp:CheckBox ID="chkConsolidarEmpresa" ToolTip="Consolidar o cnpj da empresa." runat="server"
                        Text="Empresa:" data-ToolTip="default" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="600px" />
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
                    Data Limite do Financeiro:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" runat="server" Width="100px" class="calendario" data-ToolTip="default" ToolTip="Informar a data limite do financeiro que deseja consultar." />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="chkPedidosComSaldo" runat="server" Checked="True" GroupName="TipoRelatorio"
                        Text="Pedidos Com Saldo" data-ToolTip="default" ToolTip="Status dos pedidos para listagem." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:consultaclientes id="ucConsultaClientes" runat="server" />
</asp:Content>
