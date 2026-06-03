<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucComissoesXBaixas.ascx.vb"
    Inherits="NGS.Web.UI.ucComissoesXBaixas" %>
<script type="text/javascript">
    function pageLoadComissoesXBaixas() {
        $("#btnConfirmar", "#divComissoesXBaixas").button();
        $("#btnFechar", "#divComissoesXBaixas").button();

        $("input[type='text'].txtDecimal", "#divComissoesXBaixas").change(function () {
            if ($.trim($(this).val()) != "" && $.trim($(this).val()) != undefined) {
                calcTotal($.trim($(this).val()));
            }
        });
    }

    function calcTotal(valor) {
        var valores = "";
        $("input[type='text'].txtDecimal", "#divComissoesXBaixas").each(function () {
            if ($.trim($(this).val()) != "" && $.trim($(this).val()) != undefined) {
                valores += $(this).val();
                if (valores != "") {
                    valores += ";";
                }
            }
        });

        $.ajax({
            type: "POST",
            async: true,
            url: rootPath + "/WebMethods.asmx/calcValores",
            data: "{ valores: '" + valores + "' }",
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (msg) {
                var resultado = eval('(' + msg.d + ')');
                if (resultado != "" && resultado != undefined) {
                    $("#MainContent_ucComissoesXBaixas_txtVlrRateado", "#divComissoesXBaixas").val("");
                    $("#MainContent_ucComissoesXBaixas_txtVlrRateado", "#divComissoesXBaixas").val(resultado);
                    $.ajax({
                        type: "POST",
                        async: true,
                        url: rootPath + "/WebMethods.asmx/calcDiff",
                        data: "{ total: '" + $("#MainContent_ucComissoesXBaixas_txtVlrTotal", "#divComissoesXBaixas").val() + "', valor: '" + $("#MainContent_ucComissoesXBaixas_txtVlrRateado", "#divComissoesXBaixas").val() + "' }",
                        contentType: "application/json; charset=utf-8",
                        dataType: "json",
                        success: function (msg) {
                            var diff = eval('(' + msg.d + ')');
                            $("#MainContent_ucComissoesXBaixas_txtDiff", "#divComissoesXBaixas").val("");
                            $("#MainContent_ucComissoesXBaixas_txtDiff", "#divComissoesXBaixas").val(diff);
                        },
                        error: function (jqXHR, textStatus, errorThrown) {
                            alert("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')");
                            //msgbox("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')", "POSSÍVEL ERRO OCORRIDO!", "Erro");
                        }
                    });
                }
            },
            error: function (jqXHR, textStatus, errorThrown) {
                alert("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')");
               // msgbox("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')", "POSSÍVEL ERRO OCORRIDO!", "Erro");
            }
        });
    }

    $(document).ready(function () {
        pageLoadComissoesXBaixas();
    });

    var prmComissoesXBaixas = Sys.WebForms.PageRequestManager.getInstance();
    prmComissoesXBaixas.add_endRequest(pageLoadComissoesXBaixas);
</script>
<div id="divComissoesXBaixas" class="uc" title="Comissões x Baixas" style="display: none;">
    <asp:UpdatePanel ID="updpnlComissoesXBaixas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovoAut" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimparAut" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconSair" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grd" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" ShowHeaderWhenEmpty="True"
                    EmptyDataText="NENHUM REGISTRO ENCONTRADO!" OnRowDataBound="grd_RowDataBound">
                    <EditRowStyle BackColor="#999999" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EmptyDataRowStyle VerticalAlign="Middle" HorizontalAlign="Center" Font-Bold="True" />
                    <Columns>
                        <asp:BoundField DataField="Pedido" HeaderText="Pedido">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Empresa" HeaderText="Empresa">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Operacao" HeaderText="Operação">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Produto" HeaderText="Produto">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Comissao" DataFormatString="{0:N2}" HeaderText="Comissão">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Baixado" DataFormatString="{0:N2}" HeaderText="Baixado">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Saldo" DataFormatString="{0:N2}" HeaderText="Saldo">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Vlr Rateado">
                            <ItemTemplate>
                                <asp:TextBox ID="txtValor" runat="server" CssClass="txtDecimal" Text='<%# String.Format("{0:N2}", Decimal.Zero)%>'
                                    BorderColor="White" Width="100px" Style="text-align: right;" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="painelright">
                <div class="row">
                    <div class="collbluc">
                        VALOR TOTAL:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtVlrTotal" runat="server" Enabled="false" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        VALOR RATEADO:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtVlrRateado" runat="server" Enabled="false" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        VALOR DIFERENÇA:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDiff" runat="server" Enabled="false" />
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
