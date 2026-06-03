<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucPedidoLancamentoItem.ascx.vb"
    Inherits="NGS.Web.UI.ucPedidoLancamentoItem" %>
<script type="text/javascript">
    function pageLoadPedidoLancamentoItem() {
        $("#MainContent_ucPedidoXLancamento_txtDataEntregaItem", "#divPedidoLancamentoItem").datepicker({
            dateFormat: 'dd/mm/yy',
            dayNames: ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'],
            dayNamesMin: ['D', 'S', 'T', 'Q', 'Q', 'S', 'S', 'D'],
            dayNamesShort: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'],
            monthNames: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
            monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
            nextText: 'Próximo',
            prevText: 'Anterior',
            showOn: "button",
            buttonImage: "Images/calendar.png",
            buttonImageOnly: true
        }).setMask("date");

        $("#MainContent_ucPedidoXLancamento_txtQuantidadeFat", "#divPedidoLancamentoItem").keypress(function (e) {
            var tecla = (e.keyCode ? e.keyCode : e.which);
            if (tecla == 13) {
                var txtunitfat = $("#MainContent_ucPedidoXLancamento_txtUnitarioFat", "#divPedidoLancamentoItem");
                var txtqtde = $("#MainContent_ucPedidoXLancamento_txtQuantidade", "#divPedidoLancamentoItem");
                if (txtunitfat != null) {
                    if (parseFloat($(this).val()) != 0) {
                        txtqtde.attr("disabled", true);
                        txtunitfat.focus();
                        txtunitfat.val(txtunitfat.val() + ' ');
                    }
                    else {
                        txtqtde.attr("disabled", false);
                        txtqtde.focus();
                        txtqtde.val(txtunitfat.val() + ' ');
                    }
                }
            }
        });

        $("#MainContent_ucPedidoXLancamento_txtQuantidade", "#divPedidoLancamentoItem").keypress(function (e) {
            var tecla = (e.keyCode ? e.keyCode : e.which);
            if (tecla == 13) {
                var txtunitario = $("#MainContent_ucPedidoXLancamento_txtUnitario", "#divPedidoLancamentoItem");
                var txtqtdefat = $("#MainContent_ucPedidoXLancamento_txtQuantidadeFat", "#divPedidoLancamentoItem");
                if (txtunitario != null) {
                    if (parseFloat($(this).val()) != 0) {
                        txtqtdefat.attr("disabled", true);
                        txtunitario.focus();
                        txtunitario.val(txtunitario.val() + ' ');
                    }
                    else {
                        txtqtdefat.attr("disabled", false);
                        txtqtdefat.focus();
                        txtqtdefat.val(txtunitfat.val() + ' ');
                    }
                }
            }
        });

        $("#MainContent_ucPedidoXLancamento_txtUnitario", "#divPedidoLancamentoItem").keypress(function (e) {
            var tecla = (e.keyCode ? e.keyCode : e.which);
            if (tecla == 13) {
                var lnknovo = $("#MainContent_ucPedidoXLancamento_txtUnitarioFat", "#divPedidoLancamentoItem");
                if (lnknovo != null) {
                    lnknovo.focus();
                }
            }
        });

        $("#MainContent_ucPedidoXLancamento_txtUnitarioFat", "#divPedidoLancamentoItem").keypress(function (e) {
            var tecla = (e.keyCode ? e.keyCode : e.which);
            if (tecla == 13) {
                var lnknovo = $("#MainContent_ucPedidoXLancamento_txtUnitario", "#divPedidoLancamentoItem");
                if (lnknovo != null) {
                    lnknovo.focus();
                }
            }
        });

        $("#MainContent_ucPedidoXLancamento_txtUnitario", "#divPedidoLancamentoItem").change(function () {
            calcValorDaqui(1);
        });

        $("#MainContent_ucPedidoXLancamento_txtQuantidade", "#divPedidoLancamentoItem").change(function () {
            calcValorDaqui(2);
        });

        $("#MainContent_ucPedidoXLancamento_txtUnitarioFat", "#divPedidoLancamentoItem").change(function () {
            calcValorDaqui(3);
        });

        $("#MainContent_ucPedidoXLancamento_txtQuantidadeFat", "#divPedidoLancamentoItem").change(function () {
            calcValorDaqui(4);
        });
    }

    function calcValorDaqui(Quem) {
        var unitario = $("#MainContent_ucPedidoXLancamento_txtUnitario").val();
        var quantidade = $("#MainContent_ucPedidoXLancamento_txtQuantidade").val();
        var unitarioFat = $("#MainContent_ucPedidoXLancamento_txtUnitarioFat").val();
        var quantidadeFat = $("#MainContent_ucPedidoXLancamento_txtQuantidadeFat").val();
        var fatorConversao = $("#MainContent_ucPedidoXLancamento_HFC").val();

        var parameters = "{ unitario: '" + unitario + "', quantidade: '" + quantidade + "', unitarioFat: '" + unitarioFat + "', quantidadeFat: '" + quantidadeFat + "', fatorConversao: '" + fatorConversao + "', Quem:" + Quem + "}";

        $.ajax({
            type: "POST",
            async: true,
            url: rootPath + "/WebMethods.asmx/calcValorEdson",
            data: parameters,
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (msg) {
                var valor = eval('(' + msg.d + ')');
                var vlr = valor.split(";");
                $("#MainContent_ucPedidoXLancamento_txtUnitario").val(vlr[0]);
                $("#MainContent_ucPedidoXLancamento_txtQuantidade").val(vlr[1]);
                $("#MainContent_ucPedidoXLancamento_txtUnitarioFat").val(vlr[2]);
                $("#MainContent_ucPedidoXLancamento_txtQuantidadeFat").val(vlr[3]);
                $("#MainContent_ucPedidoXLancamento_txtValorTotal").val(vlr[4]);
                    
            },
            error: function (jqXHR, textStatus, errorThrown) {
                alert("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')");
                //msgbox("'(textStatus: '" + textStatus + "', errorThrown: '" + errorThrown + "')", "ATENÇÃO!", "Info");

            }
        });
    }

    $(document).ready(function () {
        pageLoadPedidoLancamentoItem();
    });

    var prmPedidoLancamentoItem = Sys.WebForms.PageRequestManager.getInstance();
    prmPedidoLancamentoItem.add_endRequest(pageLoadPedidoLancamentoItem);
</script>
<div id="divPedidoLancamentoItem" class="uc" title="Lançamentos do Pedido" style="margin-right: 10px; clear: both; display: none;">
    <asp:UpdatePanel ID="updpnlPedidoLancamentoItem" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HFMediaMoedaCompra" runat="server" />
            <asp:HiddenField ID="HFMediaOficialCompra" runat="server" />
            <asp:HiddenField ID="HFC" runat="server" />
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="HIDLinhaProduto" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkProdutoNovo" Text="Novo Produto" runat="server" />
                        </li>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" TabIndex="3" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconSair" runat="server">
                            <asp:LinkButton ID="LnkSair" Text="Sair" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="painelleft">
                    <div class="row">
                        <div class="collbluc">
                            Grupo:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlGrupoProduto" runat="server" AutoPostBack="True" Width="265px" />
                        </div>
                        <div class="coltxt">
                            <asp:LinkButton ID="lnkBuscaProduto" runat="server" Height="20px" Width="20px">
                                <asp:Image ID="imgSearch" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                                    ToolTip="Consulta Produto" />
                            </asp:LinkButton>
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Produto:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlProdutos" runat="server" AutoPostBack="True" Width="285px" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Classificação:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlClassificacao" runat="server" Width="285px" />
                        </div>
                    </div>
                </div>
                <div class="painelleft">
                    <div class="row">
                        <div class="collbluc">
                            Movimento:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtDataMovimentoItem" runat="server" CssClass="calendario" Width="80px"
                                data-ToolTip="default" ToolTip="Data da criação do pedido." />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Entrega:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtDataEntregaItem" runat="server" AutoPostBack="True" Width="80px"
                                data-ToolTip="default" ToolTip="Refere-se ao prazo de final da safra." />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Condição:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlTipoLancamento" runat="server" AutoPostBack="True" Width="135px">
                                <asp:ListItem Value="N">Normal</asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                </div>
                <div class="painelleft">
                    <div class="subtitulodiv">
                        Unidade
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Faturamento:
                        </div>
                        <div class="coltxt">
                            <asp:Literal ID="lblUnidFat" runat="server" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Comercialização:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlUnidadeComercializacao" runat="server" Width="150px" AutoPostBack="True" />
                        </div>
                    </div>
                </div>
                <div class="painelleft" runat="server">
                    <div class="subtitulodiv">
                        Quantidade
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:TextBox ID="txtQuantidadeFat" runat="server" AutoPostBack="false" CssClass="txtDecimal4"
                                TabIndex="1" Style="color: Blue; text-align: right" data-ToolTip="default" ToolTip="">0,0000</asp:TextBox>
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:TextBox ID="txtQuantidade" TabIndex="4" runat="server" AutoPostBack="false"
                                CssClass="txtDecimal4" Style="color: Blue; text-align: right" data-ToolTip="default" ToolTip="">0,0000</asp:TextBox>
                        </div>
                    </div>
                </div>
                <div class="painelleft" runat="server">
                    <div class="subtitulodiv">
                        Unitário
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:TextBox ID="txtUnitarioFat" runat="server" AutoPostBack="false" CssClass="txtDecimal10"
                                TabIndex="2" Style="color: Blue; text-align: right" data-ToolTip="default" ToolTip="">0,0000000000</asp:TextBox>
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:TextBox ID="txtUnitario" runat="server" AutoPostBack="false" CssClass="txtDecimal10"
                                Style="color: Blue; text-align: right" TabIndex="5" data-ToolTip="default" ToolTip="">0,0000000000</asp:TextBox>
                        </div>
                    </div>
                </div>
                <div class="painelleft">
                    <div class="subtitulodiv">
                        Valor
                    </div>
                    <div class="row">
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:TextBox ID="txtValorTotal" runat="server" CssClass="txtDecimal" Style="color: Blue; text-align: right"
                                Width="120px" />
                        </div>
                    </div>
                </div>

                <div class="painelleft" runat="server">
                    <div class="subtitulodiv">
                        Vlr. em Moeda Estrangeira
                    </div>
                    <div class="row">
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:TextBox ID="txtValorTotalMoeda" runat="server" CssClass="txtDecimal" Style="color: Blue; text-align: right"
                                Width="120px" AutoPostBack="true" />
                        </div>
                    </div>
                </div>


            </div>
            <div class="row" runat="server">
                <div class="collbl">
                    Retenção:
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkRetencao" runat="server" Text="haverá retenção dos seguintes encargos -> "
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="painelleft" style="width: 70%; height: 320px;">
                <div class="bordagrid" style="height: 320px;">
                    <asp:GridView ID="gridProdutos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%">
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:TemplateField HeaderText="Produto">
                                <ItemTemplate>
                                    <asp:Label ID="Label1" runat="server" Text='<%# Bind("ItemPedido.CodigoProduto") %>' />
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Nome">
                                <ItemTemplate>
                                    <asp:Label ID="Label2" runat="server" Text='<%# Bind("ItemPedido.Descricao") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Umidade">
                                <ItemTemplate>
                                    <asp:Label ID="Label3" runat="server" Text='<%# Bind("ItemPedido.Classificacao.Descricao") %>' />
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="TipoLancamento" HeaderText="Condição">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento"
                                HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="DataEntrega" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Entrega"
                                HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="Quantidade">
                                <ItemTemplate>
                                    <div runat="server" style="float: left; margin-right: 3px;">
                                        <asp:Label ID="lblQtdeFat" runat="server" Text='<%# Bind("QuantidadeFaturamento", "{0:N4}") %>' />
                                        <br />
                                        <asp:Label ID="lblQtdeCom" runat="server" Text='<%# Bind("QuantidadeComercializacao", "{0:N4}") %>' />
                                    </div>
                                    <div style="float: right;">
                                        <asp:Label ID="lblUnidFat" runat="server" Text='<%# Bind("ItemPedido.Produto.Unidade") %>' />
                                        <br />
                                        <asp:Label ID="lblUnidCom" runat="server" Text='<%# Bind("ItemPedido.CodigoUnidadeComercializacao")%>' />
                                    </div>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Unitário Compra">
                                <ItemTemplate>
                                    <asp:Label ID="lblUnitCompFat" runat="server" Text='<%# Bind("UnitarioCompra", "{0:N10}") %>' />
                                    <br />
                                    <asp:Label ID="lblUnitCompCom" runat="server" Text='<%# Bind("UnitarioCompraComercializacao", "{0:N10}") %>' />
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Unitário">
                                <ItemTemplate>
                                    <div runat="server">
                                        <asp:Label ID="lblUnitFat" runat="server" Text='<%# Bind("UnitarioFaturamento", "{0:N10}") %>' />
                                        <br />
                                        <asp:Label ID="lblUnitCom" runat="server" Text='<%# Bind("UnitarioComercializacao", "{0:N10}") %>' />
                                    </div>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="Total" DataFormatString="{0:N2}" HeaderText="Total " HtmlEncode="False">
                                <HeaderStyle HorizontalAlign="Right" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="TotalOficial" Visible="false">
                                <ItemTemplate>
                                    <asp:Label ID="lblTotalOficial" runat="server" Text='<%# Bind("TotalOficial", "{0:N2}") %>' />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                            </asp:TemplateField>

                            <asp:TemplateField HeaderText="Excluir">
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkExcluirLancamento" runat="server" OnClick="lnkExcluirLancamento_Click">
                                        <asp:Image ID="imgExcluirLancamento" ImageUrl="~/Images/deletar.gif" runat="server"
                                            data-ToolTip="default" ToolTip="Excluir" />
                                    </asp:LinkButton>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            </asp:TemplateField>
                        </Columns>
                        <EditRowStyle BackColor="#999999" />
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    </asp:GridView>
                </div>
            </div>
            <div class="painelright" style="width: 29%;">
                <div class="painelleft">
                    <div class="subtitulodiv">
                        Resumo Do Item
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Quantidade:
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Unitário Compra:
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Unitário Medio:
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Valor:
                        </div>
                    </div>
                </div>
                <div class="painelright">
                    <div class="subtitulodiv">
                        Unid.Faturamento
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:Label ID="txtResumoQtde" runat="server" Style="text-align: right" data-ToolTip="default" ToolTip="Quantidade total vendida/comprada." />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:Label ID="txtUnitarioCompra" runat="server" Style="text-align: right" data-ToolTip="default" ToolTip="Valor unitário da Compra." />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:Label ID="txtResumoUntMedio" runat="server" Style="text-align: right" data-ToolTip="default" ToolTip="Valor médio das unidades vendidas/compradas." />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:Label ID="txtResumoValor" runat="server" Style="text-align: right" data-ToolTip="default" ToolTip="Total dos produtos comprados/vendidos." />
                        </div>
                    </div>
                </div>
                <div class="painelright">
                    <div class="subtitulodiv">
                        Unid.Comercialização
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:Label ID="txtResumoQtdeCom" runat="server" Style="text-align: right" data-ToolTip="default" ToolTip="" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:Label ID="txtUnitarioCompraCom" runat="server" Style="text-align: right" data-ToolTip="default" ToolTip="" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:Label ID="txtResumoUntMedioCom" runat="server" Style="text-align: right" data-ToolTip="default" ToolTip="" />
                        </div>
                    </div>
                    <div class="row">
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
