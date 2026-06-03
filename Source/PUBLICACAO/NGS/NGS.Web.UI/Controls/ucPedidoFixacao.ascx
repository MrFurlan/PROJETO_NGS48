<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucPedidoFixacao.ascx.vb"
    Inherits="NGS.Web.UI.ucPedidoFixacao" %>
<%@ Register TagPrefix="ngs" TagName="Financeiro" Src="~/Controls/ucFinanceiro.ascx" %>

<div id="divPedidoFixacao" class="uc" title="Fixações do Item do Pedido" style="display: none;">
    <asp:UpdatePanel ID="updpnlPedidoFixacao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="HIDPedido" runat="server" />
            <asp:HiddenField ID="HIDLinhaProduto" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" Text="Gravar" ID="lnkNovo" runat="server" OnClientClick=" if(!confirm('Deseja realmente Incluir a Fixação?')) return false;" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" Text="Valores" ID="lnkCalcularValores" runat="server"
                                OnClientClick=" if(!confirm('Deseja Calcular a Fixação?')) return false;" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" Text="Limpar" ID="lnkLimpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" Text="Sair" ID="LnkSair" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Produto:
                </div>
                <div class="coltxt" style="width: 420px;">
                    <asp:Label ID="lblProdutoAfixar" runat="server" Font-Bold="True" />
                </div>
                <div class="collbluc">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataAfixar" runat="server" Enabled="False" Width="100px" AutoPostBack="True" CssClass="calendario" />
                </div>
                <div class="collbluc">
                    Quantidade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtQuantidadeAFixar" runat="server" CssClass="txtDecimal4" Enabled="False" />
                </div>
                <div class="collbluc">
                    Saldo:
                </div>
                <div class="coltxt">
                    <asp:Label ID="txtSaldoAfixar" runat="server" Font-Bold="True" ForeColor="Red" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Fixação:
                </div>
                <div class="coltxt" style="width: 420px;">
                    <asp:Label ID="lblFixacao" runat="server" Font-Bold="True" ForeColor="Red" />
                </div>
                <div class="collbluc">
                    Fixar Dólar:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtIndiceFixadoItem" runat="server" CssClass="txtDecimal8" data-ToolTip="default"
                        ToolTip="Só fixe o dólar com autorização do Comercial. Caso o dólar seja fixado, será desconsiderado o ptax do dia" />
                    <img src="../Images/question.jpg" alt="" width="19" height="19" id="imgAjudaFixarDolarItem"
                        onclick="helpFixarDolar();" style="cursor: pointer" />
                </div>
                <div class="collbluc">
                    Unitario:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtUnitarioAfixar" runat="server" CssClass="txtDecimal10" Enabled="False" />
                </div>
                <div class="collbluc">
                    Total Bruto:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTotalAfixar" runat="server" ReadOnly="True" BackColor="#FFFFCC" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSubAFixar" runat="server" Enabled="False" Width="420px" />
                </div>
                <div class="collbluc">
                    Procurações:
                </div>
                <div class="coltxt" style="width: 132px;">
                    <asp:TextBox ID="txtCessaoDeCreditoAfixar" runat="server" Enabled="False" />
                    <asp:ImageButton ID="imgCessaoDeCredito" runat="server" ImageUrl="~/Images/search.png"
                        Style="border: 0;" data-ToolTip="default" ToolTip="Selecionar Cessão de Crédito" Visible="false" />
                </div>
                <div class="collbluc">
                    Saldo Ces. Cred.:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtSaldoProcuracao" runat="server" CssClass="txtDecimal4" Enabled="False" />
                </div>
                <div class="collbluc">
                    Total Liquido:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtTotalAfixarLiquido" runat="server" BackColor="#FFFFCC" ReadOnly="True" />
                </div>
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="gridFixacao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarFixacao" runat="server" Height="20px" ImageUrl="~/Images/search.png"
                                                data-ToolTip="default" ToolTip="Consultar Fixação" Width="20px" Style="cursor: pointer"
                                                OnClick="imgConsultarFixacao_Click" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField HeaderText="Fixa&#231;&#227;o" DataField="Codigo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Cessão Crédito" DataField="Procuracao">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Movimento" DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Quantidade" DataField="Quantidade">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Unit. R$" DataField="UnitarioOficial">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Unit. U$" DataField="UnitarioMoeda">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Valor R$" DataField="TotalOficial">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Valor U$" DataField="TotalMoeda">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Cond.Pgto." DataField="CondicaoPagamento">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirFixacao" runat="server" Height="18px" ImageUrl="~/Images/erro.jpg"
                                                data-ToolTip="default" ToolTip="Excluir Fixação" Width="18px" Style="cursor: pointer"
                                                OnClientClick="if(!confirm('Deseja realmente excluir a fixação?')) return false;"
                                                OnClick="imgExcluirFixacao_Click" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
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
                    </ContentTemplate>
                    <HeaderTemplate>
                        Fixações
                    </HeaderTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="gridFixacaoXNotaFiscal" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField HeaderText="E/S">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <%# Eval("NotaFiscalXItem.NotaFiscal.EntradaSaida")%>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Serie">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                        <ItemTemplate>
                                            <%#Eval("NotaFiscalXItem.NotaFiscal.Serie")%>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Nota">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                        <ItemTemplate>
                                            <%#Eval("NotaFiscalXItem.NotaFiscal.Codigo")%>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="QtdeFixacao" DataFormatString="{0:N6}" HeaderText="Qtde Fixado">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorFixacao" DataFormatString="{0:N2}" HeaderText="Valor Fixado">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <%--<asp:BoundField DataField="ValorFixacaoReal" DataFormatString="{0:N2}" HeaderText="Vlr. Fixado Real">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>--%>
                                    <asp:BoundField DataField="QtdeAfixarLiberado" DataFormatString="{0:N2}" HeaderText="Qtde Liberado">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorAFixarLiberado" DataFormatString="{0:N2}" HeaderText="Valor Liberado">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                    <HeaderTemplate>
                        Notas Fiscais
                    </HeaderTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel3" runat="server" HeaderText="TabPanel3">
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="grdEncargosFixacao" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField HeaderText="Fixação">
                                        <ItemTemplate>
                                            <asp:Label ID="Label1" runat="server" Text='<%# Bind("FixacaoPedido.Codigo") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" Width="120px" />
                                        <ItemStyle HorizontalAlign="Left" Width="120px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoEncargo" FooterText="Encargo" HeaderText="Encargo"
                                        ReadOnly="True">
                                        <FooterStyle HorizontalAlign="Left" />
                                        <HeaderStyle HorizontalAlign="Left" Width="120px" />
                                        <ItemStyle HorizontalAlign="Left" Width="120px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Percentual" DataFormatString="{0:N9}" FooterText="Aliquota"
                                        HeaderText="Aliquota" ReadOnly="True">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="BaseOficial" DataFormatString="{0:N2}" HeaderText="Base R$"
                                        ReadOnly="True">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" FooterText="Valor"
                                        HeaderText="Valor R$" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="BaseMoeda" HeaderText="Base U$" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorMoeda" HeaderText="Valor U$" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                    <HeaderTemplate>
                        Encargos
                    </HeaderTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabVencimentosOld" runat="server" HeaderText="VencimentoOld">
                    <ContentTemplate>
                        <div id="divVencimento" runat="server">
                            <div class="menu_acoes">
                                <div class="acoes">
                                    <ul>
                                        <li runat="server">
                                            <asp:LinkButton class="iconNovo" ID="lnkParcelar" Text="Atualizar Parcelas" runat="server" />
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbluc">
                                    Condições:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlCondicoesFixacao" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlCondicoesFixacao_SelectedIndexChanged"
                                        Width="395px" />
                                </div>
                                <div class="collbluc">
                                    Data:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataVencimentoFixacao" runat="server" CssClass="calendario" Enabled="False"
                                        Width="70px" />
                                </div>
                                <div class="collbluc">
                                    Valor:
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="ValorVencimentoFixacao" runat="server" />
                                    <asp:TextBox ID="txtValorVencimentoFixacao" runat="server" CssClass="txtDecimal"
                                        Enabled="False" />
                                </div>
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridCondicoesFixacao" runat="server" AutoGenerateColumns="False"
                                OnSelectedIndexChanged="gridCondicoesFixacao_SelectedIndexChanged" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField SelectText="&gt;&gt;" ShowSelectButton="True" ButtonType="Button">
                                        <HeaderStyle HorizontalAlign="Center" Width="30px"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center" Width="30px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Codigo">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" DataFormatString="{0:n2}" HeaderText="Valor R$"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorMoeda" DataFormatString="{0:n2}" HeaderText="Valor U$"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataPagamento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Pagto." HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorPagamento" DataFormatString="{0:n2}" HeaderText="Pagto. R$"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorPagamentoMoeda" DataFormatString="{0:n2}" HeaderText="Pagto. U$"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Saldo" DataFormatString="{0:n2}" HeaderText="Saldo R$"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="SaldoMoeda" DataFormatString="{0:n2}" HeaderText="Saldo U$"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Atraso" DataFormatString="{0:n2}" HeaderText="Atraso"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
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
                            <div class="collbl">
                                Total Parcelado:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTotalPaceladoFixacao" runat="server" ForeColor="Red" Font-Bold="True"
                                    BackColor="#FFFFC0" ReadOnly="True" />
                            </div>
                            <div class="collbl">
                                Total Pago:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTotalPagoFixacao" runat="server" ForeColor="Red" Font-Bold="True"
                                    BackColor="#FFFFC0" ReadOnly="True" />
                            </div>
                            <div class="collbl">
                                Saldo:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSaldoVencimentosFixacao" runat="server" ForeColor="Red" Font-Bold="True"
                                    BackColor="#FFFFC0" ReadOnly="True" />
                            </div>
                        </div>
                    </ContentTemplate>
                    <HeaderTemplate>
                        Vencimentos
                    </HeaderTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabVencimentos" runat="server" HeaderText="Vencimentos">
                    <HeaderTemplate>
                        Vencimentos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <ngs:Financeiro ID="ngsFinanceiro" runat="server"></ngs:Financeiro>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel4" runat="server" HeaderText="TabPanel4">
                    <HeaderTemplate>
                        Observações
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton CssClass="iconMais" ID="lnkAdicionarObservacaoFixacao" runat="server"
                                            Text="Adicionar Observação" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Texto:
                            </div>
                            <div class="coltxt" style="width: 89%;">
                                <asp:TextBox ID="txtAddObservacaoFixacao" runat="server" TextMode="MultiLine" Width="100%" />
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            Observação
                        </div>
                        <div class="row">
                            <div class="coltxt" style="width: 99%;">
                                <asp:TextBox ID="txtObservacaoFixacao" runat="server" Width="100%" TextMode="MultiLine"
                                    Rows="1" Height="210px" ReadOnly="True" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
