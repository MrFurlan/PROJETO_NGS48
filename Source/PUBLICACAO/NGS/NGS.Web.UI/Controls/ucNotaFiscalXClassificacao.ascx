<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucNotaFiscalXClassificacao.ascx.vb"
    Inherits="NGS.Web.UI.ucNotaFiscalXClassificacao" %>
<script type="text/javascript">
    function pageLoadNotaFiscalXClassificacao() {
        $("#btnCalcular", "#divConsultaNotaFiscalXClassificacao").button();
        $("#btnOK", "#divConsultaNotaFiscalXClassificacao").button();
        $("#btnLimpar", "#divConsultaNotaFiscalXClassificacao").button();
        $("#btnFechar", "#divConsultaNotaFiscalXClassificacao").button();
        $(".txtInteiro").setMask('integer');
        $(".txtDecimal").setMask('decimal');


        $('.txt', '#MainContent_ucNotaFiscalXClassificacao_GridDescontos').keydown(function (e) {
            //get the next index of text input element
            var next_idx = $('.txt', '#MainContent_ucNotaFiscalXClassificacao_GridDescontos').index(this) + 1;
            //get number of text input element in a html document
            var tot_idx = $('#MainContent_ucNotaFiscalXClassificacao_GridDescontos').find('.txt').length;
            //enter button in ASCII code
            if (e.keyCode == 13) {
                if (next_idx == ($("#MainContent_ucNotaFiscalXClassificacao_GridDescontos > tbody > tr").size() - 1)) {
                    //go to the next button input element
                    $("#btnCalcular", "#divConsultaNotaFiscalXClassificacao").focus();
                }
                else if (tot_idx == next_idx) {
                    //go to the first text element if focused in the last text input element
                    $('.txt:eq(0)', '#MainContent_ucNotaFiscalXClassificacao_GridDescontos').focus();
                }
                else {
                    //go to the next text input element
                    $('.txt:eq(' + next_idx + ')', '#MainContent_ucNotaFiscalXClassificacao_GridDescontos').focus();
                }
                return false;
            }
        });


        $("#txtPrimeiraPesagem", "#divConsultaNotaFiscalXClassificacao").keydown(function (e) {
            if (e.keyCode == 13) {
                if ($("#txtSegundaPesagem", "#divConsultaNotaFiscalXClassificacao") != undefined) {
                    $('.txt:eq(0)', '#MainContent_ucNotaFiscalXClassificacao_GridDescontos').focus();
                }
                else {
                    $("#txtSegundaPesagem", "#divConsultaNotaFiscalXClassificacao").val("").focus();
                }


                $('select', '#MainContent_ucNotaFiscalXClassificacao_GridDescontos').each(function () {
                    $(this).removeClass("aspNetDisabled");
                    $(this).removeAttr("disabled");
                });
            }
        });

        $("#txtSegundaPesagem", "#divConsultaNotaFiscalXClassificacao").keydown(function (e) {
            if (e.keyCode == 13) {
                //setarPesoBruto($(this).val());
                $('.txt:eq(0)', '#MainContent_ucNotaFiscalXClassificacao_GridDescontos').focus();
            }
        });
    }

    $(document).ready(function () {
        pageLoadNotaFiscalXClassificacao();
    });

    var prmNotaFiscalXClassificacao = Sys.WebForms.PageRequestManager.getInstance();
    prmNotaFiscalXClassificacao.add_endRequest(pageLoadNotaFiscalXClassificacao);
</script>
<style type="text/css">
    .collbluc {
        width: 100px;
    }
</style>

<div id="divConsultaNotaFiscalXClassificacao" class="uc" title="Classificação Nota Fiscal" style="display: none;">
    <asp:UpdatePanel ID="updpnlNotaFiscalXClassificacao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="txtEntradaSaida" runat="server" />
            <asp:HiddenField ID="txtProduto" runat="server" />
            <div class="subtitulodiv" style="text-align: center !important;">
                <asp:Label ID="lblEntSai" runat="server" />
            </div>
            <div class="painelleft" style="width: 39%;">
                <div class="row">
                    <div class="collbluc">
                        Classificação:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="DdlClassificacao" runat="server" Width="150px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        <asp:Label ID="lblPesagem" runat="server" Text="Pesagem"></asp:Label>
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPrimeiraPesagem" runat="server" ClientIDMode="Static" CssClass="txtInteiro"
                            TabIndex="160" Width="80px" />
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkPesagens" runat="server" Text="Habil. 2º Pesagen" AutoPostBack="true" Checked="false" />
                    </div>
                </div>
                <div class="row" id="row2Pesagem" runat="server">
                    <div class="collbluc">
                        2º Pesagem:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtSegundaPesagem" runat="server" ClientIDMode="Static" CssClass="txtInteiro"
                            TabIndex="160" Width="80px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Peso Bruto:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPesoBruto" runat="server" Enabled="False" Width="80px" CssClass="txtInteiro" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Desconto:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDesconto" runat="server" Enabled="False" Width="80px" CssClass="txtInteiro" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Liquido:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtLiquido" runat="server" Enabled="False" Width="80px" CssClass="txtInteiro" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc">
                        Observação:
                    </div>
                    <div class="coltxt" style="width: 235px;">
                        <asp:TextBox ID="txtObservacao" runat="server" Width="100%" TextMode="MultiLine" />
                    </div>
                </div>
            </div>
            <div class="painelleft" style="width: 60%;">
                <div class="bordagrid" style="height: 207px;">
                    <asp:GridView ID="GridDescontos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:BoundField DataField="CodigoAnalise" HeaderText="Código" ItemStyle-VerticalAlign="Middle" ItemStyle-HorizontalAlign="Center">
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Descricao" HeaderText="Descrição" />
                            <asp:TemplateField HeaderText="Percentual">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtPercentual" runat="server" CssClass="txtDecimal txt" Width="80px" Text='<%# IIf(Convert.ToString(Eval("Percentual")) = "0", "",  Eval("Percentual", "{0:N2}"))%>' />
                                    <asp:DropDownList ID="ddlOpcao" runat="server" Visible="False" Width="98%" CssClass="txt"></asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Indice">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtIndice" runat="server" Enabled="False" Width="56px" Text='<%# Eval("Indice")%>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Desconto">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtDesconto" runat="server" Enabled="False" Width="80px" Text='<%# Eval("Desconto")%>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </div>
                <div class="row" style="text-align: right;">
                    <div class="coltxt">
                        <asp:Button ID="btnOK" runat="server" Visible="false" ClientIDMode="Static" CssClass="botao" Text="OK"
                            UseSubmitBehavior="False" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnCalcular" runat="server" ClientIDMode="Static" CssClass="botao"
                            Text="Calcular" UseSubmitBehavior="False" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnAjustar" runat="server" Visible="false" ClientIDMode="Static" CssClass="botao" Text="Ajustar"
                            UseSubmitBehavior="False" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnCriarRomaneio" runat="server" Visible="false" ClientIDMode="Static" CssClass="botao" Text="Cria Romaneio"
                            UseSubmitBehavior="False" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnLimpar" runat="server" ClientIDMode="Static" CssClass="botao"
                            Text="Limpar" UseSubmitBehavior="False" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnFechar" runat="server" ClientIDMode="Static" CssClass="botao"
                            Text="Fechar" UseSubmitBehavior="False" />
                    </div>
                </div>
            </div>
        </ContentTemplate>
        <Triggers>
            <asp:AsyncPostBackTrigger ControlID="btnOK" EventName="Click" />
            <asp:AsyncPostBackTrigger ControlID="btnCalcular" EventName="Click" />
            <asp:AsyncPostBackTrigger ControlID="btnAjustar" EventName="Click" />
            <asp:AsyncPostBackTrigger ControlID="btnCriarRomaneio" EventName="Click" />
            <asp:AsyncPostBackTrigger ControlID="btnLimpar" EventName="Click" />
            <asp:AsyncPostBackTrigger ControlID="btnFechar" EventName="Click" />
        </Triggers>
    </asp:UpdatePanel>
</div>
