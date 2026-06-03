<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucProdutoCupomFiscal.ascx.vb"
    Inherits="NGS.Web.UI.ucProdutoCupomFiscal" %>
<script type="text/javascript">
    function pageLoadProdutoCupomFiscal() {
        $("#MainContent_ucProdutoCupomFiscal_txtUnitario", "#divProdutoCupomFiscal").change(function () {
            calcValor();
        });

        $("#MainContent_ucProdutoCupomFiscal_txtQuantidade", "#divProdutoCupomFiscal").change(function () {
            calcValor();
        });
    }

    function calcValor() {
        var unitario = $("#MainContent_ucProdutoCupomFiscal_txtUnitario", "#divProdutoCupomFiscal").val();
        var quantidade = $("#MainContent_ucProdutoCupomFiscal_txtQuantidade", "#divProdutoCupomFiscal").val();
        var parameters = "{ unitario: '" + unitario + "', quantidade: '" + quantidade + "' }";

        $.ajax({
            type: "POST",
            async: true,
            url: rootPath + "/WebMethods.asmx/calcValor",
            data: parameters,
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (msg) {
                var valor = eval('(' + msg.d + ')');
                $("#MainContent_ucProdutoCupomFiscal_txtValorTotal", "#divProdutoCupomFiscal").val(valor);
            },
            error: function (jqXHR, textStatus, errorThrown) {
                alert("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')");
                //msgbox("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')", "POSSÍVEL ERRO OCORRIDO!", "Erro");
            }
        });
    }

    $(document).ready(function () {
        pageLoadProdutoCupomFiscal();
    });

    var prmProdutoCupomFiscal = Sys.WebForms.PageRequestManager.getInstance();
    prmProdutoCupomFiscal.add_endRequest(pageLoadProdutoCupomFiscal);
</script>
<div id="divProdutoCupomFiscal" class="uc" title="Informe o produto" style="display: none;">
    <asp:UpdatePanel ID="updpnlProdutoCupomFiscal" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="PosicaoItem" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkGravar" Text="Adicionar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconFechar" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbGrupoProduto" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbGrupoProduto_SelectedIndexChanged"
                        TabIndex="13" Width="620px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbProdutos" runat="server" Width="590px" OnSelectedIndexChanged="cmbProdutos_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnConsultaProduto" CssClass="btn" runat="server" UseSubmitBehavior="False"
                        Text=" > " data-ToolTip="default" ToolTip="Nome do produto." />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Quantidade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtQuantidade" runat="server" CssClass="txtDecimal4" TabIndex="14"
                        Width="150px" data-ToolTip="default" ToolTip="Quantidade comprada/vendida." />
                </div>
                <div class="collbluc">
                    Unitário:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtUnitario" runat="server" CssClass="txtDecimal10" TabIndex="15"
                        Width="100px" data-ToolTip="default" ToolTip="valor unitário comprado/vendido." />
                </div>
                <div class="collbluc">
                    Total:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtImpostosOld" runat="server" Value="0" />
                    <asp:HiddenField ID="txtValorTotalOld" runat="server" Value="0" />
                    <asp:TextBox ID="txtValorTotal" runat="server" BackColor="#FFFFC0" CssClass="txtDecimal"
                        Font-Bold="True" ForeColor="Red" ReadOnly="True" TabIndex="16" Width="100px" data-ToolTip="default" ToolTip="Saldo total da compra/venda." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
