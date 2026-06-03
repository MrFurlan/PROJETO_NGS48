<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucPesoDeChegada.ascx.vb"
    Inherits="NGS.Web.UI.ucPesoDeChegada" %>
<script type="text/javascript">
    //function pageLoadPesoDeChegada() {

    //}

    function SumDescontos() {
        var txtDesconto = $("#<%=txtDesconto.ClientID%>", "#divPesoDeChegada");
        txtDesconto.val(0);
        var total = 0.00;
        $(".desc", "#MainContent_ucPesoDeChegada_GridDescontos").each(function () {
            if ($(this).val() != '' && $(this).val() != undefined) {
                var desc = parseFloat($(this).val().replace('.', '').replace(',', '.'));
                total = total + desc;
            }
        });
        txtDesconto.val(total.toFixed(2).replace(".", ","));
    }

    //$(document).ready(function () {
    //    pageLoadPesoDeChegada();
    //});

    //var prmPesoDeChegada = Sys.WebForms.PageRequestManager.getInstance();
    //prmPesoDeChegada.add_endRequest(pageLoadPesoDeChegada);

</script>
<div id="divPesoDeChegada" class="uc" title="Consulta de Dados da Chegada" style="display: none;">
    <asp:UpdatePanel ID="updpnlPesoDeChegada" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <script type="text/javascript">
                function TotalizarFrete() {
                    var txtTarifaFrete = $("#<%=txtTarifaFrete.ClientID%>", "#divPesoDeChegada");
                    var txtBrutoDeChegada = $("#<%=txtBrutoDeChegada.ClientID%>", "#divPesoDeChegada");
                    var txtValorDoFrete = $("#<%=txtValorDoFrete.ClientID%>", "#divPesoDeChegada");

                    if (txtTarifaFrete.val() != undefined && txtTarifaFrete.val() != '' && txtBrutoDeChegada.val() != undefined && txtBrutoDeChegada.val() != '') {
                        var tarifaFrete = parseFloat(txtTarifaFrete.val().replace('.', '').replace(',', '.'));
                        var brutoDeChegada = parseFloat(txtBrutoDeChegada.val().replace('.', '').replace(',', '.'));
                        var valorDoFrete = brutoDeChegada * tarifaFrete / 1000;
                    }
                    txtValorDoFrete.val(valorDoFrete.toFixed(2).replace(".", ","));
                }
            </script>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton ID="lnkConfirmar" Class="iconNovo" Text="Confirmar" runat="server"
                                OnClientClick="if(!confirm('Confirma inclusão do peso de chegada?')) return false;" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton ID="lnkFechar" Class="iconSair" runat="server" Text="Fechar" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Bruto de Saída:
                </div>
                <div class="coltxt" style="width: 156px;">
                    <asp:TextBox ID="txtBrutoDeSaida" runat="server" CssClass="txtInteiro" Width="125px"
                        Enabled="false" />
                </div>
                <div class="collbluc">
                    Bruto de Chegada:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtBrutoDeChegada" runat="server" CssClass="txtInteiro" AutoPostBack="True"
                        MaxLength="6" OnTextChanged="txtBrutoDeChegada_TextChanged" Style="color: red; text-align: right"
                        Width="125px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Data de Chegada:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataDeChegada" runat="server" CssClass="calendario" Width="125px" />
                </div>
                <div class="collbluc">
                    Quebra / Sobra:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSaldo" runat="server" CssClass="txtInteiro" Width="125px" Enabled="false" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Descontos:
                </div>
                <div class="coltxt" style="width: 156px;">
                    <asp:TextBox ID="txtDesconto" runat="server" CssClass="txtInteiro" AutoPostBack="True"
                        Width="125px" MaxLength="9" OnTextChanged="txtDesconto_TextChanged" />
                </div>
                <div class="collbluc">
                    Líquido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtLiquido" runat="server" CssClass="txtInteiro" Width="125px" MaxLength="9"
                        Enabled="false" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Tarifa do Frete:
                </div>
                <div class="coltxt" style="width: 156px;">
                    <asp:TextBox ID="txtTarifaFrete" runat="server" ToolTip="Tarifa do Frete por Tonelada."
                        onblur="TotalizarFrete();" CssClass="txtDecimal" Width="125px" />
                </div>
                <div class="collbluc">
                    Valor Total do Frete:
                </div>
                <div class="coltxt" style="width: 156px;">
                    <asp:TextBox ID="txtValorDoFrete" runat="server" ToolTip="Valor Total do Frete."
                        Enabled="false" CssClass="txtDecimal" Width="125px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    <asp:CheckBox ID="chkSinistro" runat="server" Text="Possui sinistro?" />
                </div>
            </div>
            <div class="row">
                <div class="bordagrid" style="height: auto;">
                    <asp:GridView ID="GridDescontos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%">
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <EditRowStyle BackColor="#999999" />
                        <Columns>
                            <asp:BoundField DataField="Codigo" HeaderText="Código">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            </asp:BoundField>
                            <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="Percentual">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                <ItemTemplate>
                                    <asp:TextBox ID="txtPercentual" runat="server" Width="80px" CssClass="txtDecimal" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Índice">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                <ItemTemplate>
                                    <asp:TextBox ID="txtIndice" runat="server" Enabled="false" Text="0,00" CssClass="txtDecimal"
                                        Width="80px" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Desconto">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                <ItemTemplate>
                                    <asp:TextBox ID="txtDesconto" runat="server" Width="80px" CssClass="txtDecimal desc"
                                        onchange="SumDescontos();" />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <%--<div class="row">
                <div class="painelright">
                    <asp:Button ID="btnConfirmar" runat="server" Text="Confirmar" Width="90px" OnClientClick="if(!confirm('Confirma inclusão do peso de chegada?')) return false;" />
                    <asp:Button ID="btnFechar" runat="server" Text="Fechar" Width="90px" />
                </div>
            </div>--%>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
