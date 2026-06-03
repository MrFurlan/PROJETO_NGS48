<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="RelatorioFiscalPorOperacao.aspx.vb" Inherits="NGS.Web.UI.RelatorioFiscalPorOperacao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        var i = 1;
        function pageLoadRFPO() {
            $('#' + "<%= txtMes.ClientID%>").onblur(function (e) {
                if ($(this).val() != '' && $(this).val() > 12) {
                    msgbox(i, "ATENÇÃO!", "Info");
                    i++;
                    $(this).val('');
                }
            });
        }

        $(document).ready(function () {
            pageLoadRFPO();
            var prmRFPO = Sys.WebForms.PageRequestManager.getInstance();
            prmRFPO.add_endRequest(pageLoadRFPO);
        });
    </script>    
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngPlanoDeContas" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPlanoDeContas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Relatório Fiscal Por Operação
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server"><a class="iconRelatorio rel">Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton class="iconPdf" ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel" />
                                </li>
                            </ul>
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="618px" AutoPostBack="true" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="618px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCliente" Enabled="false" runat="server" Width="580px" />
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnBuscaCliente" runat="server" Text=">" CssClass="btn" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Classe:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlClasse" runat="server" Width="618px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProdutoVenda" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Safra:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSafra" runat="server" Width="318px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Mês:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtMes" runat="server" Width="40px" CssClass="txtNumerico2" />
                </div>

                <div class="collbl">
                    Ano:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAno" runat="server" Width="40px" CssClass="txtNumerico4" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
