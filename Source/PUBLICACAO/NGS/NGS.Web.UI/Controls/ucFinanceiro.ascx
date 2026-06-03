<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucFinanceiro.ascx.vb"
    Inherits="NGS.Web.UI.ucFinanceiro" %>
<script type="text/javascript" language="javascript">

    function EscondePedidoTroca(esconde) {
        if (esconde) {
            $('.w3').css("display", "none");
        }
        else {
            $('.w3').css("display", "block");
        }
    }
</script>
<style type="text/css">
    .st {
        width: 377px;
    }

    .w1 {
        width: 370px;
    }

    .w2, .w3 {
        width: 100px;
        text-align: right;
        padding-right: 10px;
    }
</style>
<div>
    <asp:HiddenField ID="HID" runat="server" />
    <asp:HiddenField ID="HOrigem" runat="server" />
    <asp:HiddenField ID="IndiceGrid" runat="server" Value="-1" />
    <asp:Panel ID="pnlFormaPagamento" runat="server" Width="100%">
        <div class="menu_acoes" id="divBarraBotoes" runat="server">
            <div class="acoes">
                <ul>
                    <li id="Li1" class="iconMais" runat="server">
                        <asp:LinkButton ID="LnkParcelamento" runat="server" Text="Parcelar" />
                    </li>
                    <li id="Li2" class="iconLimpar" runat="server">
                        <asp:LinkButton ID="lnkLimpar" runat="server" Text="Limpar" />
                    </li>
                    <li id="Li3" class="iconMais" runat="server">
                        <asp:LinkButton ID="LnkReajuste" runat="server" Text="Reajuste" />
                    </li>
                </ul>
            </div>
        </div>
        <div id="divEntrega" runat="server">
            <div id="divLinhaFaturamento" runat="server" visible="false" class="row">
                <div class="collbl" style="width: 125px;">
                    Faturamento:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlTipoFaturamento" runat="server" AutoPostBack="True" Width="340px"
                        OnSelectedIndexChanged="ddlTipoFaturamento_SelectedIndexChanged" />
                </div>
            </div>
            <div id="divLinhaPeridiocidade" runat="server" visible="false" class="row">
                <div class="collbl" style="width: 125px;">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlPeridiocidade" runat="server" AutoPostBack="True" Width="340px">
                        <asp:ListItem Value="7" Text="Semanal" />
                        <asp:ListItem Value="15" Text="Quinzenal" />
                        <asp:ListItem Value="30" Text="Mensal" />
                    </asp:DropDownList>
                </div>
            </div>
            <div id="divLinhaQdteLote" runat="server" visible="false" class="row">
                <div class="collbl" style="width: 125px;">
                    Quantidade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtQtdeLote" runat="server" CssClass="txtDecimal4" />
                </div>
            </div>
            <div id="divLinhaCondicoesPgtoEntrega" runat="server" visible="false" class="row">
                <div class="collbl" style="width: 125px;">
                    Condições:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="lstCondicoesPgtoEntrega" runat="server" AutoPostBack="False"
                        Width="340px" OnSelectedIndexChanged="lstCondicoesPgtoEntrega_SelectedIndexChanged" />
                </div>
            </div>
            <div id="divLinhaQuotaDeEntrega" runat="server" visible="false" class="row">
                <div class="collbl" style="width: 125px;">
                    Quota da Entrega:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtQuotaDeEntrega" runat="server" AutoPostBack="True" BackColor="#FFFFC0"
                        CssClass="txtDecimal4" Font-Bold="True" ForeColor="Red" Text="0" OnTextChanged="txtQuotaDeEntrega_TextChanged" />
                </div>
            </div>
        </div>
        <div id="divLinhaCondicaoPagamento" runat="server" visible="false" class="row">
            <div class="collbl" style="width: 125px;">
                Condições:
            </div>
            <div class="coltxt">
                <asp:DropDownList ID="ddlCondicaoPagamento" runat="server" AutoPostBack="False" OnSelectedIndexChanged="ddlCondicaoPagamento_SelectedIndexChanged"
                    Width="340px" />
            </div>
        </div>
        <div id="divFormaDePagamento" runat="server" visible="false" class="row">
            <div class="collbl" style="width: 125px;">
                Forma de Pagamento:
            </div>
            <div class="coltxt">
                <asp:DropDownList ID="ddlFormaPagamentoGeral" runat="server" AutoPostBack="False"
                    Width="340px" />
            </div>
        </div>
    </asp:Panel>
    <div class="subtitulodiv">
        Vencimentos
    </div>
    <div class="painelleft" style="width: 51%;">
        <div id="divLinhaAdiantamento" runat="server">
            <div class="subtitulodiv">
                Adiantamentos
            </div>
            <div class="bordagrid" style="height: 215px;">
                <asp:GridView ID="gridAdiantamentosDisponiveis" runat="server" AutoGenerateColumns="False"
                    Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <Columns>
                        <asp:BoundField DataField="CodigoTitulo" HeaderText="Título" />
                        <asp:BoundField DataField="DescMoeda" HeaderText="Moeda" />
                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento" />
                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento" />
                        <asp:BoundField DataField="Valor" DataFormatString="{0:N2}" HeaderText="Valor" />
                        <asp:BoundField DataField="TotalBaixado" DataFormatString="{0:N2}" HeaderText="Valor Baixado" />
                        <asp:BoundField DataField="SaldoValor" DataFormatString="{0:N2}" HeaderText="Saldo" />
                        <asp:BoundField DataField="Taxa" DataFormatString="{0:N2}" HeaderText="Taxa" />
                        <asp:TemplateField HeaderText="Valor Baixa">
                            <ItemTemplate>
                                <asp:TextBox ID="txtVlrBaixa" runat="server" CssClass="txtDecimal" DataFormatString="{0:N2}"
                                    Text='<%# Eval("VlrBaixa", "{0:N2}") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                    <EditRowStyle BackColor="#999999" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                </asp:GridView>
                <asp:GridView ID="gridBaixasAdiantamentos" runat="server" AutoGenerateColumns="False"
                    Width="100%" CellPadding="4" ForeColor="#333333" GridLines="None">
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <Columns>
                        <asp:BoundField DataField="CodigoTituloAdiantamento" DataFormatString="{0:N0}" HeaderText="Adiant." />
                        <asp:BoundField DataField="DataAdiantamento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Dt Adiant." />
                        <asp:BoundField DataField="ValorAdiantamentoOficial" DataFormatString="{0:N2}" HeaderText="Vlr Ad. Oficial" />
                        <asp:BoundField DataField="IndiceAdiantamento" DataFormatString="{0:N4}" HeaderText="Indice Ad." />
                        <asp:BoundField DataField="ValorAdiantamentoMoeda" DataFormatString="{0:N2}" HeaderText="Vlr Ad. Moeda" />
                        <asp:BoundField DataField="CodigoTituloBaixa" HeaderText="Titulo BX" />
                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Dt Baixa" />
                        <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Vlr BX Oficial" />
                        <asp:BoundField DataField="IndiceBaixa" DataFormatString="{0:N4}" HeaderText="Indice Bx" />
                        <asp:BoundField DataField="ValorMoeda" DataFormatString="{0:N2}" HeaderText="Vlr BX Moeda" />
                        <asp:BoundField DataField="SaldoOficialAdiantamento" DataFormatString="{0:N2}" HeaderText="Saldo Ad Oficial" />
                        <asp:BoundField DataField="SaldoMoedaAdiantamento" DataFormatString="{0:N2}" HeaderText="Saldo Ad Moeda" />
                    </Columns>
                    <EditRowStyle BackColor="#999999" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                </asp:GridView>
            </div>
        </div>
        <div class="row" id="divLinhaTipoDoTitulo" runat="server">
            <div class="collbluc">
                Situação:
            </div>
            <div class="coltxt">
                <asp:DropDownList ID="ddlSituacao" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlSituacao_SelectedIndexChanged">
                    <asp:ListItem Value="0">Todos</asp:ListItem>
                    <asp:ListItem Value="2">Previsão</asp:ListItem>
                    <asp:ListItem Value="3">Provisão</asp:ListItem>
                    <asp:ListItem Value="1">Baixa</asp:ListItem>
                    <asp:ListItem Value="4">Compensação</asp:ListItem>
                    <asp:ListItem Value="5">Adiantamentos</asp:ListItem>
                </asp:DropDownList>
            </div>
        </div>
        <div class="bordagrid" style="height: 215px;">
            <asp:GridView ID="gridVencimentosFinanceiroNovo" runat="server" AutoGenerateColumns="False" CellPadding="4"
                Width="100%" GridLines="None" OnSelectedIndexChanged="gridVencimentos_SelectedIndexChanged">
                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <Columns>
                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                    <asp:BoundField DataField="Codigo" HeaderText="Titulo" />
                    <asp:BoundField DataField="ReceberPagar" HeaderText="Tipo" />
                    <asp:BoundField DataField="DescricaoProvisao" HeaderText="Situação" />
                    <asp:BoundField DataField="Reprogramacao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:BoundField>
                    <asp:BoundField DataField="DescricaoMoeda" HeaderText="Moeda" />
                    <asp:TemplateField HeaderText="Vlr Doc. Oficial" Visible="false">
                        <ItemTemplate>
                            <asp:Label ID="Label1" runat="server" Text='<%# Bind("Valores.EncargoValorDocumento.ValorOficial", "{0:N2}") %>' />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Vlr Liq. Oficial">
                        <ItemTemplate>
                            <asp:Label ID="Label2" runat="server" Text='<%# Bind("Valores.EncargoValorLiquido.ValorOficial", "{0:N2}") %>' />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Vlr Doc. Moeda" Visible="false">
                        <ItemTemplate>
                            <asp:Label ID="Label3" runat="server" Text='<%# Bind("Valores.EncargoValorDocumento.ValorMoeda", "{0:N2}") %>' />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Vlr Liq. Moeda">
                        <ItemTemplate>
                            <asp:Label ID="Label4" runat="server" Text='<%# Bind("Valores.EncargoValorLiquido.ValorMoeda", "{0:N2}") %>' />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:BoundField DataField="Quantidade" HeaderText="Lote" ItemStyle-HorizontalAlign="Right" />
                    <asp:BoundField DataField="Quantidade" HeaderText="Saldo" ItemStyle-HorizontalAlign="Right" />
                </Columns>
                <EditRowStyle BackColor="#999999" />
                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            </asp:GridView>
            <asp:GridView ID="gridVencimentos" runat="server" AutoGenerateColumns="false" CellPadding="4"
                Width="100%" GridLines="None" OnSelectedIndexChanged="gridVencimentos_SelectedIndexChanged">
                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <Columns>
                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                    <asp:BoundField DataField="Codigo" HeaderText="Titulo" />
                    <asp:BoundField DataField="ReceberPagar" HeaderText="Tipo" />
                    <asp:BoundField DataField="DescricaoProvisao" HeaderText="Situação" />
                    <asp:BoundField DataField="Prorrogacao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:BoundField>
                    <asp:BoundField DataField="DescricaoMoeda" HeaderText="Moeda" />
                    <asp:TemplateField HeaderText="Vlr Doc. Oficial" Visible="false">
                        <ItemTemplate>
                            <asp:Label ID="Label1" runat="server" Text='<%# Bind("ValorDoDocumento", "{0:N2}") %>' />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Vlr Liq. Oficial">
                        <ItemTemplate>
                            <asp:Label ID="Label2" runat="server" Text='<%# Bind("ValorLiquido", "{0:N2}") %>' />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Vlr Doc. Moeda" Visible="false">
                        <ItemTemplate>
                            <asp:Label ID="Label3" runat="server" Text='<%# Bind("MoedaValorDoDocumento", "{0:N2}") %>' />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Vlr Liq. Moeda">
                        <ItemTemplate>
                            <asp:Label ID="Label4" runat="server" Text='<%# Bind("MoedaValorLiquido", "{0:N2}") %>' />
                        </ItemTemplate>
                        <ItemStyle HorizontalAlign="Right" />
                    </asp:TemplateField>
                    <asp:BoundField DataField="Quantidade" HeaderText="Lote" ItemStyle-HorizontalAlign="Right" />
                    <asp:BoundField DataField="Quantidade" HeaderText="Saldo" ItemStyle-HorizontalAlign="Right" />
                </Columns>
                <EditRowStyle BackColor="#999999" />
                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            </asp:GridView>
        </div>
    </div>
    <div class="painelleft" style="width: 48%; margin-left: 8px;">
        <asp:Panel ID="pnlTitulos" runat="server" Width="100%" Visible="false">
            <div class="row">
                <div class="collbluc">
                    Titulo:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtCodigoTitulo" runat="server" Font-Bold="True" />&nbsp;<asp:Label
                        ID="txtMoeda" runat="server" Font-Bold="True" />
                </div>
                <div class="collbluc">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataVencimento" Width="80px" runat="server" CssClass="calendario" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Tipo Pagto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlTipoDePagto" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlTipoDePagto_SelectedIndexChanged"
                        Width="370px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Cod.Barra:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigoDeBarra" runat="server" AutoPostBack="True" Width="361px" />&nbsp;
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkDigitado" runat="server" AutoPostBack="True" OnCheckedChanged="chkDigitado_CheckedChanged"
                        Text="Digitado" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkPreImpresso" runat="server" AutoPostBack="True" Text="Pré-Impresso" />
                </div>
            </div>
            <div class="bordagrid" style="height: 257px;">
                <asp:GridView ID="gridValores" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    Width="100%" ForeColor="#333333" GridLines="None">
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <Columns>
                        <asp:BoundField DataField="CodigoContaEncargo" HeaderText="Contas " />
                        <asp:BoundField DataField="Descricao" HeaderText="Titulo" />
                        <asp:TemplateField HeaderText="Valor Oficial">
                            <ItemTemplate>
                                <asp:TextBox ID="txtValorOficial" runat="server" CssClass="txtDecimal" AutoPostBack="true"
                                    OnTextChanged="txtValorOficial_TextChanged" Text='<%# Bind("ValorOficial", "{0:N2}") %>'
                                    Width="100px" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Valor Moeda">
                            <ItemTemplate>
                                <asp:TextBox ID="txtValorMoeda" runat="server" CssClass="txtDecimal" Text='<%# Bind("ValorMoeda", "{0:N2}") %>'
                                    Width="100px" OnTextChanged="txtValorMoeda_TextChanged" AutoPostBack="True" />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Deb./Cred.">
                            <ItemTemplate>
                                <asp:DropDownList ID="ddlDebitoCredito" runat="server" Enabled="False" SelectedValue='<%# Bind("DC") %>'
                                    Width="126px">
                                    <asp:ListItem Selected="True" Value="D">Debito</asp:ListItem>
                                    <asp:ListItem Value="C">Credito</asp:ListItem>
                                    <asp:ListItem Value="I"> INFORME</asp:ListItem>
                                </asp:DropDownList>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Centro De Custo" Visible="False">
                            <ItemTemplate>
                                <asp:DropDownList ID="ddlCentroDeCusto" runat="server" AutoPostBack="True" Width="369px">
                                </asp:DropDownList>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                    <EditRowStyle BackColor="#999999" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                </asp:GridView>
            </div>
            <div class="row">
                <asp:Button ID="BtnAdicionarConta" runat="server" Text=" + " UseSubmitBehavior="False"
                    CssClass="btn" />
                <asp:HiddenField ID="ValorDocumento" runat="server" />
            </div>
            <div id="divLinhaBancariaTitulo" runat="server" visible="false" class="subtitulodiv">
                Dados Bancários
            </div>
            <div id="divLinhaBancariaDados" runat="server" class="row">
                <div class="collbluc">
                    Banco:
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnConta" runat="server" OnClick="BtnConta_Click" Text="&gt;" UseSubmitBehavior="False"
                        CssClass="btn" />&#160;&nbsp;
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblBanco" runat="server" Font-Bold="True" Text="Banco" />&#160;&nbsp;
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblAgencia" runat="server" Font-Bold="True" Text="Agência" />&#160;&nbsp;
                </div>
                <div class="coltxt">
                    <asp:Label ID="lblContaCorrente" runat="server" Font-Bold="True" Text="Conta" />
                </div>
            </div>
            <div class="row">
                <asp:Button ID="btnAtualizarTitulo" CssClass="botao" runat="server" Text="Atualizar" />
            </div>
        </asp:Panel>
    </div>
    <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%"
        Visible="false">
        <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
            <HeaderTemplate>
                Programação Financeira do Documento
            </HeaderTemplate>
            <ContentTemplate>
                <!-- Retirado -->
            </ContentTemplate>
        </ajaxToolkit:TabPanel>
        <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
            <HeaderTemplate>
                Resumo Financeiro
            </HeaderTemplate>
            <ContentTemplate>
                <div class="row">
                    <div class="coltxt">
                        <div class="subtitulodiv st">
                            Pedido x Financeiro
                        </div>
                    </div>
                    <div class="coltxt">
                        <div class="subtitulodiv w2">
                            Este Pedido
                        </div>
                    </div>
                    <div class="coltxt">
                        <div class="subtitulodiv w3">
                            Pedido De Troca
                        </div>
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Pedido
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtPedido01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtPedido02" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Valor Total Pedido
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtValorPedido01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtValorPedido02" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="coltxt">
                        <div class="subtitulodiv st">
                            Financeiro Sintetico
                        </div>
                    </div>
                    <div class="coltxt">
                        <div class="subtitulodiv w2">
                        </div>
                    </div>
                    <div class="coltxt">
                        <div class="subtitulodiv w3">
                        </div>
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Valor Titulos em Previsão (Gerado por Pedido / Fixação / Devoluções)
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtValorTitulosPrevisao01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtValorTitulosPrevisao02" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Valor Titulos em Provisão (Titulos Com Nota)
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtValorTitulosProvisao01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtValorTitulosProvisao02" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Valor Titulos Baixado
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtValorTitulosBaixados01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtValorTitulosBaixados02" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Valor Titulos Compensados
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtValorTitulosCompensados01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtValorTitulosCompensados02" runat="server" />
                    </div>
                </div>
                <br />
                <div class="row">
                    <div class="collbluc w1">
                        Valor Adiantamento Original
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtValorAdiantamentoOriginal01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtValorAdiantamentoOriginal02" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Valor Adiantamento
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtValorAdiantamento01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtValorAdiantamento02" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Valor Adiantamento Compensado
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtAdiantamentoCompensado01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtAdiantamentoCompensado02" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Valor Adiantado Pago (Em Dinheiro)
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtAdiantamentoPago01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtAdiantamentoPago02" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Valor Amortizado pelo Adiantamento
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtAdiantamentoAmortizado01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtAdiantamentoAmortizado02" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Saldo Restante p/ Baixa Adiantamento
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtSaldoAdiantamento01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtSaldoAdiantamento02" runat="server" />
                    </div>
                </div>
                <br />
                <div class="row">
                    <div class="collbluc w1">
                        Valor Pago:
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtValorPago01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtValorPago02" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbluc w1">
                        Saldo Restante p/ Baixa Pedido
                    </div>
                    <div class="coltxt w2">
                        <asp:Label ID="txtSaldoBaixa01" runat="server" />
                    </div>
                    <div class="coltxt w3">
                        <asp:Label ID="txtSaldoBaixa02" runat="server" />
                    </div>
                </div>
            </ContentTemplate>
        </ajaxToolkit:TabPanel>
    </ajaxToolkit:TabContainer>
</div>
