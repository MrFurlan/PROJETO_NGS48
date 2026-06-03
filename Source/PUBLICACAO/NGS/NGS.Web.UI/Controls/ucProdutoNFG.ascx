<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucProdutoNFG.ascx.vb"
    Inherits="NGS.Web.UI.ucProdutoNFG" %>
<script type="text/javascript">
    function pageLoadProdutoNFG() {
        $("#MainContent_ucProdutoNFG_txtUnitario", "#divProdutoNFG").change(function () {
            calcValor();
        });

        $("#MainContent_ucProdutoNFG_txtQuantidade", "#divProdutoNFG").change(function () {
            calcValor();
        });
    }

    function calcValor() {
        var unitario = $("#MainContent_ucProdutoNFG_txtUnitario", "#divProdutoNFG").val();
        var quantidade = $("#MainContent_ucProdutoNFG_txtQuantidade", "#divProdutoNFG").val();
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
                $("#MainContent_ucProdutoNFG_txtValorTotal", "#divProdutoNFG").val(valor);
            },
            error: function (jqXHR, textStatus, errorThrown) {
                alert("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')");
                //msgbox("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')", "POSSÍVEL ERRO OCORRIDO!", "Erro");
            }
        });
    }

    $(document).ready(function () {
        pageLoadProdutoNFG();
    });

    var prmProdutoNFG = Sys.WebForms.PageRequestManager.getInstance();
    prmProdutoNFG.add_endRequest(pageLoadProdutoNFG);
</script>
<div id="divProdutoNFG" class="uc" title="Informe o produto" style="display: none;">
    <asp:UpdatePanel ID="updpnlProdutoNFG" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hidReabroveitarDados" runat="server" />
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
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtOperacao" runat="server" Font-Names="monospace" Enabled="false" Width="589px"
                        data-ToolTip="default" ToolTip="Referente a operação de criação da nota." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imbOperacao" runat="server" Style="cursor: pointer; border: 0 none;"
                        data-ToolTip="default" ToolTip="Consulta de Operações" ImageAlign="AbsMiddle"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imbOperacao_Click" />
                </div>
            </div>
            <div class="row" id="divProdutoCusto" runat="server" visible="false">
                <div class="collbluc">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbGrupoProdutoCusto" runat="server" AutoPostBack="True" OnSelectedIndexChanged="cmbGrupoProdutoCusto_SelectedIndexChanged" Width="620px" />
                </div>
                <div class="collbluc">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbProdutosCusto" runat="server" Width="590px" OnSelectedIndexChanged="cmbProdutosCusto_SelectedIndexChanged" AutoPostBack="True" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnConsultaProdutoCusto" CssClass="btn" runat="server" UseSubmitBehavior="False"
                        Text=" > " data-ToolTip="default" ToolTip="Nome do produto." />
                </div>
            </div>
            <div id="divCentroDeCusto" class="row" runat="server">
                <div class="collbluc">
                    Centro de Custo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbCentroCusto" runat="server" Font-Names="monospace" Width="620px" />
                </div>
            </div>
            <div class="row" runat="server">
                <div class="collbl">
                    Retenção:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkRetencao" runat="server" Text="haverá retenção dos seguintes encargos -> " AutoPostBack="True" data-ToolTip="default" ToolTip="" />
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
