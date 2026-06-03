<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaPedidoDeTroca.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaPedidoDeTroca" %>
<div id="divConsultaPedidoDeTroca" class="uc" title="Consulta Pedido de Troca" style="display: none;">
    <asp:UpdatePanel ID="updConsultaPedidoDeTroca" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConfirmar" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkConfirmar" runat="server">
                                <span>Confirmar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconSair" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkFechar" runat="server">
                                <span>Fechar</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="gridview" style="border-color: blue;">
                <div class="subtitulodiv">
                    Pedidos
                </div>
                <asp:GridView ID="gridPedido" runat="server" CellPadding="0" ForeColor="#333333"
                    GridLines="None" Width="100%" PageSize="5" AutoGenerateColumns="False">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Empresa_Id" HeaderText="Empresa" />
                        <asp:BoundField DataField="EndEmpresa_Id" HeaderText=" " />
                        <asp:BoundField DataField="Nome" HeaderText=" " />
                        <asp:BoundField DataField="Safra" HeaderText="Safra" />
                        <asp:BoundField DataField="DataPedido" HeaderText="DataPedido" />
                        <asp:BoundField DataField="Pedido_Id" HeaderText="Pedido" />
                        <asp:BoundField DataField="PedidoEfetivo" HeaderText="Pedido(CN)" />
                        <asp:BoundField DataField="Cliente" HeaderText="Cliente" />
                        <asp:BoundField DataField="EndCliente" HeaderText=" " />
                        <asp:BoundField DataField="ClienteNome" HeaderText=" " />
                        <asp:BoundField DataField="Complemento" HeaderText="Comp." />
                        <asp:BoundField DataField="DescMoeda" HeaderText="Moeda" />
                        <asp:BoundField DataField="Origem" HeaderText="Vlr Venda" DataFormatString="{0:N2}" />
                        <asp:BoundField DataField="Destino" HeaderText="Vlr Compra" DataFormatString="{0:N2}" />
                        <asp:BoundField DataField="Saldo" HeaderText="Saldo" DataFormatString="{0:N2}" />
                    </Columns>
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                </asp:GridView>
            </div>
            <div id="divItens" class="gridview" runat="server" style="display: none;">
                <div class="subtitulodiv">
                    Itens do Pedido
                </div>
                <asp:GridView ID="GridItens" runat="server" CellPadding="0" ForeColor="#333333"
                    GridLines="None" Width="100%" PageSize="5" AutoGenerateColumns="False">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <EditRowStyle BackColor="#999999" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <Columns>
                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" />
                        <asp:BoundField DataField="Descricao" HeaderText="Nome" />
                        <asp:TemplateField HeaderText="Qtde Pedido">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                            <ItemTemplate>
                                <div style="float: Left;">
                                    <asp:Label ID="lblQtdeFat" runat="server" Text='<%# String.Format("{0:N4}",eval("QuantidadePedidoFaturamento")) %>' />
                                    <br />
                                    <asp:Label ID="lblQtdeCom" runat="server" Text='<%# String.Format("{0:N4}",eval("QuantidadePedidoComercializacao")) %>' />
                                </div>
                                <div style="float: Right;">
                                    <asp:Label ID="lblUnFat" runat="server" Text='<%# eval("UnidadeFaturamento") %>' />
                                    <br />
                                    <asp:Label ID="lblUnCom" runat="server" Text='<%# Eval("CodigoUnidadeComercializacao")%>' />
                                </div>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Unitario Medio">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                            <ItemTemplate>
                                <asp:Label ID="lblUnMdFat" runat="server" Text='<%# String.Format("{0:N10}",eval("UnitarioMedioFaturamento")) %>' />
                                <br />
                                <asp:Label ID="lblUnMdCom" runat="server" Text='<%# String.Format("{0:N10}",eval("UnitarioMedioComercializacao")) %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="PedidoValor" DataFormatString="{0:N2}" HeaderText="Valor" />
                        <asp:TemplateField HeaderText="Qtde Entregue">
                            <ItemTemplate>
                                <div style="float: Left; padding-right: 6px;">
                                    <asp:Literal ID="Literal1" runat="server" Text="Unid/Moeda"></asp:Literal>
                                    <br />
                                    <asp:Literal ID="LD" runat="server" Text="Direta"></asp:Literal>
                                    <br />
                                    <asp:Literal ID="LG" runat="server" Text="Futura"></asp:Literal>
                                    <br />
                                    <asp:Literal ID="LR" runat="server" Text="Remessa"></asp:Literal>
                                </div>
                                <div style="float: Left;">
                                    <asp:Literal ID="QEF" runat="server" Text='<%# eval("UnidadeFaturamento") %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L17F" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.QtdeEntregueFiscalDireta")) %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L18F" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.QtdeEntregueFiscalGlobal"))%>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L19F" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.QtdeEntregueFiscalRemessa")) %>'></asp:Literal>
                                </div>
                                <div style="float: right;">
                                    <asp:Literal ID="QEC" runat="server" Text='<%# Eval("CodigoUnidadeComercializacao")%>'></asp:Literal>
                                    <br />
                                    <asp:Label ID="L17C" runat="server" Text='<%# String.Format("{0:N2}", Eval("SaldoItem.QtdeEntregueFiscalDireta") / Eval("UnidadeComercializacaoFatorDeConversao"))%>' />
                                    <br />
                                    <asp:Label ID="L18C" runat="server" Text='<%# String.Format("{0:N2}", Eval("SaldoItem.QtdeEntregueFiscalGlobal") / Eval("UnidadeComercializacaoFatorDeConversao"))%>' />
                                    <br />
                                    <asp:Label ID="L19C" runat="server" Text='<%# String.Format("{0:N2}", Eval("SaldoItem.QtdeEntregueFiscalRemessa") / Eval("UnidadeComercializacaoFatorDeConversao"))%>' />
                                </div>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Vlr Entregue">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                            <ItemTemplate>
                                <div id="Div3" runat="server" style="float: right;">
                                    <asp:Literal ID="L08C" runat="server" Text='<%# eval("SaldoItem.CifraoPedido") %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L08" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaMoedaDiretaBruto")) %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L09" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaMoedaGlobalBruto")) %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L10" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaMoedaRemessaBruto")) %>'></asp:Literal>
                                </div>
                                <div style="float: right; padding-right: 4px;">
                                    <asp:Literal ID="L05C" runat="server" Text='<%# eval("SaldoItem.CifraoOficial") %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L05" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaOficialDiretaBruto")) %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L06" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaOficialGlobalBruto")) %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L07" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.VlrNotaOficialRemessaBruto")) %>'></asp:Literal>
                                </div>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Qtde Saldo">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                            <ItemTemplate>
                                <div style="float: right;">
                                    <asp:Literal ID="QSC" runat="server" Text='<%# Eval("CodigoUnidadeComercializacao")%>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L20C" runat="server" Text='<%# String.Format("{0:N2}", Eval("SaldoItem.SaldoQtdeDiretoFiscal") / Eval("UnidadeComercializacaoFatorDeConversao"))%>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L21C" runat="server" Text='<%# String.Format("{0:N2}", Eval("SaldoItem.SaldoQtdeGlobalFiscal") / Eval("UnidadeComercializacaoFatorDeConversao"))%>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L22C" runat="server" Text='<%#String.Format("{0:N2}", eval("SaldoItem.SaldoQtdeRemessaFiscal") / eval("UnidadeComercializacaoFatorDeConversao")) %>'></asp:Literal>
                                </div>
                                <div style="float: left;">
                                    <asp:Literal ID="QSF" runat="server" Text='<%# eval("UnidadeFaturamento") %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L20F" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoQtdeDiretoFiscal")) %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L21F" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoQtdeGlobalFiscal"))%>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L22F" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoQtdeRemessaFiscal")) %>'></asp:Literal>
                                </div>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Vlr Saldo">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                            <ItemTemplate>
                                <div id="Div4" runat="server" style="float: right;">
                                    <asp:Literal ID="L14C" runat="server" Text='<%# eval("SaldoItem.CifraoPedido") %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L14" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorMoedaGlobalDireto")) %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L15" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorMoedaGlobalDireto")) %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L16" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorMoedaRemessa")) %>'></asp:Literal>
                                </div>
                                <div style="float: right; padding-right: 4px;">
                                    <asp:Literal ID="L11C" runat="server" Text='<%# eval("SaldoItem.CifraoOficial") %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L11" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorOficialGlobalDireto")) %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L12" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorOficialGlobalDireto")) %>'></asp:Literal>
                                    <br />
                                    <asp:Literal ID="L13" runat="server" Text='<%# String.Format("{0:N2}",eval("SaldoItem.SaldoValorOficialRemessa")) %>'></asp:Literal>
                                </div>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                    <EditRowStyle BackColor="#999999" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                    <SortedAscendingCellStyle BackColor="#E9E7E2" />
                    <SortedAscendingHeaderStyle BackColor="#506C8C" />
                    <SortedDescendingCellStyle BackColor="#FFFDF8" />
                    <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
