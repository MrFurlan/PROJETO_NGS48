<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucNotaFiscalReferencial.ascx.vb"
    Inherits="NGS.Web.UI.ucNotaFiscalReferencial" %>
<script type="text/javascript">
    function pageLoadNotaFiscalReferencial() {
        $("#btnFechar", "#divNotaFiscalReferencial").button();
    }

    var prmNotaFiscalReferencial = Sys.WebForms.PageRequestManager.getInstance();
    prmNotaFiscalReferencial.add_endRequest(pageLoadNotaFiscalReferencial);

    function float2Moeda(num, comDecimais) {
        x = 0;
        var resto = num.toString().split(".");
        if (num < 0) {
            num = Math.abs(num);
            x = 1;
        }
        if (isNaN(num)) num = "0";
        num = Math.floor((num * 100 + 0.5) / 100).toString();
        for (var i = 0; i < Math.floor((num.length - (1 + i)) / 3); i++)
            num = num.substring(0, num.length - (4 * i + 3)) + '.' + num.substring(num.length - (4 * i + 3));
        ret = num;
        if (x == 1) ret = ' - ' + ret;
        var retorno;
        if (comDecimais) {
            if (isNaN(resto[1])) resto[1] = "0";
            retorno = ret + "," + resto[1];
        } else {
            retorno = ret;
        }
        return retorno;
    }

    function limpaCampos() {
        $("#MainContent_ucNotaFiscalReferencial_txtPesoTotalSelecionado").val(0);
        $("#MainContent_ucNotaFiscalReferencial_txtValorTotalSelecionado").val(0.00);
        $("#MainContent_ucNotaFiscalReferencial_txtFreteValorTotal").val(0.00);

        $("input[type='checkbox']", "#MainContent_ucNotaFiscalReferencial_grdNFeSaldo").not("#chkAlls").each(function () {
            $(this).attr("checked", false);
        });
    };

</script>
<div id="divNotaFiscalReferencial" class="uc" title="Nota Fiscal Referencial" style="display: none;">
    <asp:UpdatePanel ID="updpnlNotaFiscalReferencial" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <script type="text/javascript">

                function selectAllCheckboxes(chkAll) {
                    var chk = $('#' + chkAll.id);
                    var checked = chk.attr('checked') == "checked";
                    $("#MainContent_ucNotaFiscalReferencial_txtPesoTotalSelecionado").val(0);
                    $("#MainContent_ucNotaFiscalReferencial_txtValorTotalSelecionado").val(0.00);
                    $("#MainContent_ucNotaFiscalReferencial_txtFreteValorTotal").val(0.00);

                    $("input[type='checkbox']", "#MainContent_ucNotaFiscalReferencial_grdNFeSaldo").not("#chkAlls").each(function () {
                        $(this).attr("checked", checked);
                        var txtFreteValorTotal = $(this).parent().parent().find(".txtFreteValorTotal")

                        if ($(this).attr("checked") == "checked") {
                            txtFreteValorTotal.removeClass("aspNetDisabled");
                            txtFreteValorTotal.attr("disabled", false);
                            somaSelecionados(this);
                        }
                        else {
                            txtFreteValorTotal.attr("disabled", true);
                            txtFreteValorTotal.val(0.00);
                            txtFreteValorTotal.focus();
                        }
                    });
                }

                function somaSelecionados(ischk) {
                    var checked = $('#' + ischk.id).attr('checked') == "checked";
                    var txtPeso = $(ischk).parent().parent().find(".txtVlrPeso");
                    var txtValor = $(ischk).parent().parent().find(".txtValorAtualizado");
                    var lblPesoFiscal = $(ischk).parent().parent().find(".lblPesoFiscal");
                    var txtPesoTotal = $("#MainContent_ucNotaFiscalReferencial_txtPesoTotalSelecionado");
                    var txtValorTotal = $("#MainContent_ucNotaFiscalReferencial_txtValorTotalSelecionado");
                    var txtFreteValorTotal = $(ischk).parent().parent().find(".txtFreteValorTotal");

                    //Exportação
                    if ((txtPeso.val() != undefined) && (txtValor.val() != undefined)) {
                        if ((txtPesoTotal.val() != undefined) && (txtValorTotal.val() != undefined)) {

                            txtPeso.attr('disabled', !checked);
                            var pesoTotal = parseInt(txtPesoTotal.val().split(".").join(""));
                            var valorTotal = parseFloat(txtValorTotal.val().split(".").join("").split(",").join("."));

                            if (checked) {
                                pesoTotal += parseInt(txtPeso.val().split(".").join(""));
                                valorTotal += parseFloat(txtValor.val().split(".").join("").split(",").join("."));
                            } else {
                                pesoTotal -= parseInt(txtPeso.val().split(".").join(""));
                                valorTotal -= parseFloat(txtValor.val().split(".").join("").split(",").join("."));
                            }
                            txtPesoTotal.val(float2Moeda(pesoTotal, false));
                            txtValorTotal.val(float2Moeda(valorTotal.toFixed(2), true));
                        }
                    }

                    //RPA
                    if ((lblPesoFiscal.text() != undefined) && (txtFreteValorTotal.val() != undefined)) {
                        if ((txtPesoTotal.val() != undefined) && (txtValorTotal.val() != undefined)) {
                            var pesoTotal = parseInt(txtPesoTotal.val().split(".").join(""));
                            var valorTotal = parseFloat(txtValorTotal.val().split(".").join("").split(",").join("."));

                            if (checked) {
                                pesoTotal += parseInt(lblPesoFiscal.text().split(".").join(""));
                                valorTotal += parseFloat(txtFreteValorTotal.val().split(".").join(""));
                            } else {
                                pesoTotal -= parseInt(lblPesoFiscal.text().split(".").join(""));
                                valorTotal -= parseFloat(txtFreteValorTotal.val().split(".").join(""));
                            }
                            txtPesoTotal.val(float2Moeda(pesoTotal, false));
                            txtValorTotal.val(float2Moeda(valorTotal.toFixed(2), true));
                        }
                    }

                    if (checked) {
                        txtFreteValorTotal.removeClass("aspNetDisabled");
                        txtFreteValorTotal.attr("disabled", false);
                        somaSelecionados(this);
                    }
                    else {
                        txtFreteValorTotal.attr("disabled", true);
                        txtFreteValorTotal.val(0.00);
                        txtFreteValorTotal.focus();
                    }
                }
            </script>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkConfirmar" Text="Confirmar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" OnClick="lnkLimpar_click" Text="Limpar" OnClientClick="limpaCampos();"
                                runat="server" />
                        </li>
                        <li class="iconSair" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 200px">
                    Tipo Referencial:
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblTipoReferencial" runat="server" />
                </div>
            </div>
            <div class="row" id="divValorNFOrigem" runat="server">
                <div class="collbl" style="width: 150px">
                    Valor do RPA:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValorRPA" runat="server" CssClass="txtDecimal" Enabled="false"
                        Font-Bold="true" />
                </div>
            </div>
            <div class="row">
                <div class="collbl" style="width: 150px">
                    Peso Total Selecionado:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPesoTotalSelecionado" runat="server" CssClass="txtInteiro" Text="0"
                        Enabled="false" Font-Bold="true" />
                </div>
                <div class="collbl" style="width: 200px">
                    Valor Total Selecionado:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtValorTotalSelecionado" runat="server" CssClass="txtDecimal txtVSel"
                        Text="0,00" Enabled="false" Font-Bold="true" />
                </div>
            </div>
            <div class="bordagrid" style="height: 365px;">
                <asp:GridView ID="grdNFeSaldo" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" ShowHeaderWhenEmpty="True"
                    EmptyDataText="NENHUM REGISTRO ENCONTRADO!" Visible="False" DataKeyNames="Nota_id"
                    PageSize="5">
                    <EditRowStyle BackColor="#999999" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EmptyDataRowStyle VerticalAlign="Middle" HorizontalAlign="Center" Font-Bold="True" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <Columns>
                        <asp:TemplateField>
                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            <HeaderTemplate>
                                <input id="chkAlls" onclick="selectAllCheckboxes(this);" type="checkbox" />
                            </HeaderTemplate>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkSelecionado" onclick="somaSelecionados(this);" runat="server" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="EntradaSaida_Id" HeaderText="E/S">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Nota_Id" HeaderText="Nota">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Serie_Id" HeaderText="Série">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Cliente_Id" HeaderText="Cliente">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="EndCliente_Id" HeaderText="End">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Nome" HeaderText="Nome">
                            <HeaderStyle HorizontalAlign="Left" Width="275px" Wrap="true"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="275px" Wrap="true"></ItemStyle>
                        </asp:BoundField>
                         <asp:BoundField DataField="Produto" HeaderText="Produto">
                            <HeaderStyle HorizontalAlign="Left" Width="275px" Wrap="true"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left" Width="275px" Wrap="true"></ItemStyle>
                        </asp:BoundField>
                         <asp:BoundField DataField="sequencia_id" HeaderText="Seq">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                          <asp:BoundField DataField="cfop_id" HeaderText="cfop">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Operacao" HeaderText="OP">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="SubOperacao" HeaderText="SO">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Data" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Peso">
                            <ItemTemplate>
                                <asp:Label ID="lblPesoFiscal" CssClass="lblPesoFiscal" runat="server" Text='<%# String.Format("{0:N0}", Eval("PesoFiscal"))%>' />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="ConsumidoPeso" DataFormatString="{0:N0}" HeaderText="Consumido">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="PesoSaldo" DataFormatString="{0:N0}" HeaderText="Saldo">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Saldo Kg">
                            <ItemTemplate>
                                <asp:TextBox ID="PesoSaldo" Style="text-align: right;" CssClass="txtInteiro" Text='<%# String.Format("{0:N0}", Eval("PesoSaldo"))%>'
                                    runat="server" BorderColor="White" Width="100px" Enabled="false" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Peso Kg">
                            <ItemTemplate>
                                <asp:TextBox ID="txtVlrPeso" Style="text-align: right;" CssClass="txtInteiro txtVlrPeso"
                                    Enabled="true" Text='<%# String.Format("{0:N0}", Eval("PesoSaldo"))%>' runat="server"
                                    BorderColor="White" Width="100px" AutoPostBack="true" OnTextChanged="txtVlrPeso_TextChanged" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Valor Atualizado">
                            <ItemTemplate>
                                <asp:TextBox ID="txtValorAtualizado" Style="text-align: right;" CssClass="txtDecimal txtValorAtualizado"
                                    Text='<%# Bind("ValorOficial", "{0:N2}")%>' runat="server" BorderColor="White"
                                    Width="100px" Enabled="false" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Valor Do Frete">
                            <ItemTemplate>
                                <asp:TextBox ID="txtFreteValorTotal" Style="text-align: right;" onclick="somaSelecionados(chkSelecionado);"
                                    CssClass="txtDecimal txtFreteValorTotal" Text='<%# Bind("FreteValorTotal", "{0:N2}")%>'
                                    runat="server" BorderColor="White" Width="100px" AutoPostBack="true"  OnTextChanged="txtFreteValorTotal_TextChanged" Enabled="False" />
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
            <div class="row">
                <div class="painelright" id="divPaineltotais" runat="server">
                    <div class="collbl">
                        Totais:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtSaldoTotal" runat="server" CssClass="txtInteiro" Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPesoTotal" runat="server" CssClass="txtInteiro" Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtTotalValorAtualizado" runat="server" CssClass="txtDecimal" Enabled="false" />
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
