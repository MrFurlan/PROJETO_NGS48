<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucBaixaLoteFinanceiro.ascx.vb" Inherits="NGS.Web.UI.ucBaixaLoteFinanceiro" %>

<script type="text/javascript">
    function pageLoadBaixaLoteFinanceiro() {
    }
    $(document).ready(function () {
        pageLoadBaixaLoteFinanceiro();
    });

    var prmBaixaLoteFinanceiro = Sys.WebForms.PageRequestManager.getInstance();
    prmBaixaLoteFinanceiro.add_endRequest(pageLoadBaixaLoteFinanceiro);
</script>

<div id="divBaixaLoteFinanceiro" class="uc" title="Baixa Financeiro em Lote" style="display: none;">
    <asp:UpdatePanel ID="updpnlBaixaLoteFinanceiro" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton ID="lnkBaixaLoteFinanceiro" CssClass="iconConsultar" Text="Gravar"
                                runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton ID="lnkFecharBaixaLoteFinanceiro" CssClass="iconSair" runat="server" Text="Fechar" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:Label ID="lblEmpresaLote" runat="server" Text="Empresa Pagadora:" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresaPagadoraLote" runat="server" Width="550px" OnSelectedIndexChanged="DdlEmpresaPagadoraLote_SelectedIndexChanged" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:Label ID="lblTipoPgtoRecLote" runat="server" Text="Tipo Pgto:" />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlTiposDePagamentosLote" runat="server" Width="550px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Banco:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlBancoPagadorLote" runat="server" Width="550px" OnSelectedIndexChanged="DdlBancoPagadorLote_SelectedIndexChanged" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlContaPagadoraLote" runat="server" Width="550px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Carteira do Título:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCarteiraDoTituloLote" runat="server" Width="550px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data Da Baixa:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtMovimentoLote" runat="server" Width="75px" Font-Size="9pt" Font-Bold="False"
                        ClientIDMode="Static" data-ToolTip="default" ToolTip="Data do movimento para contabilização de todos os títulos." CssClass="calendario" />
                </div>
                <div class="collbl">
                    Vencimento Atual:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtProrrogacaoLote" runat="server" Width="75px" Font-Size="9pt" Font-Bold="False"
                        ClientIDMode="Static" data-ToolTip="default" ToolTip="Data do vencimento para todos os títulos." CssClass="calendario" />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
