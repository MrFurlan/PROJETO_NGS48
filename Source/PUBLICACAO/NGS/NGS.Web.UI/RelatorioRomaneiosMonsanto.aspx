<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RelatorioRomaneiosMonsanto.aspx.vb" Inherits="NGS.Web.UI.RelatorioRomaneiosMonsanto" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .wit {
            width: 160px;
        }
    </style>

    <script type="text/javascript">
        function pageLoadRelatorioRomaneiosMonsanto() {
            $("#MainContent_lstClasseOp").multiselect().multiselectfilter();
        }

        $(document).ready(function () {
            pageLoadRelatorioRomaneiosMonsanto();
            var prmRelatorioRomaneiosMonsanto = Sys.WebForms.PageRequestManager.getInstance();
            prmRelatorioRomaneiosMonsanto.add_endRequest(pageLoadRelatorioRomaneiosMonsanto);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRelatorioRomaneiosMonsanto" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRelatorioRomaneiosMonsanto" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório de Romaneios Monsanto
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
                    <asp:DropDownList ID="cmbUnidadeNegocio" runat="server" Width="624px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbUnidadeNegocio_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbEmpresa" runat="server" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkClientes" runat="server" Text="Cliente:" />
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtClientes" runat="server" Width="585px" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdConsultaCliente" OnClick="cmdConsultaCliente_Click" CssClass="btn"
                        runat="server" Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Pedido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPedido" runat="server" Enabled="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdBuscaPedido" runat="server" CssClass="btn" CausesValidation="False"
                        OnClick="cmdBuscaPedido_Click" Text=">" UseSubmitBehavior="False" data-ToolTip="default"
                        ToolTip="Informar o número do pedido." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbOperacao" runat="server" Width="304px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbOperacao_SelectedIndexChanged" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbSubOperacao" runat="server" Width="315px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlGrupoDeProdutos" runat="server" Width="624px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlGrupoDeProdutos_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbProdutos" runat="server" Width="624px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Inicial:
                </div>
                <div class="coltxt wit">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="70px"
                        data-ToolTip="default" ToolTip="Informar o período inicial de consulta." />
                </div>
                <div class="collbl">
                    Data Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="70px"
                        data-ToolTip="default" ToolTip="Informar o período final de consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Entrada/Saída:
                </div>
                <div class="coltxt wit">
                    <asp:RadioButton ID="optEntrada" runat="server" Text="Entrada" GroupName="EntSai"
                        Checked="True" data-ToolTip="default" ToolTip="Informar se é entrada ou saída." />
                    <asp:RadioButton ID="optSaida" runat="server" Text="Saida" GroupName="EntSai" data-ToolTip="default"
                        ToolTip="Informar se é entrada ou saída." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
</asp:Content>
