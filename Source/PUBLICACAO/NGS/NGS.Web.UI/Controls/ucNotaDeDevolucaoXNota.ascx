<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucNotaDeDevolucaoXNota.ascx.vb"
    Inherits="NGS.Web.UI.ucNotaDeDevolucaoXNota" %>


<div id="divNotaDeDevolucaoXNota" class="uc" title="Consulta de Nota Devolução x Nota" style="display: none;">
    <asp:UpdatePanel ID="updpnlNotaDeDevolucaoXNota" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdfIndice" runat="server" />
            <asp:HiddenField ID="hdfControle" runat="server" />

            <div class="menu_acoes">
                <div class="acoes">
                    <ul>

                        <li class="iconNovo" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkConfirmar" runat="server">
                                <span>Confirmar</span>
                            </asp:LinkButton>
                        </li>

                        <li class="iconAtualizar" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkRecarregar" runat="server">
                                <span>Recarregar</span>
                            </asp:LinkButton>
                        </li>

                        <li class="iconLimpar" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server">
                                <span>Limpar</span>
                            </asp:LinkButton>
                        </li>

                        <li class="iconExcluir" style="width: 15%;" runat="server">
                            <asp:LinkButton ID="lnkCancelar" runat="server">
                                <span>Cancelar e Sair</span>
                            </asp:LinkButton>
                        </li>

                        <li class="iconLimpar" style="width: 10%; display: none;" runat="server">
                            <asp:LinkButton ID="lnkRegravar" runat="server">
                                <span>Regravar</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="painelleft" style="width: 35%; margin-right: 15px;">
                    <asp:Panel ID="PnlResultDev" runat="server" CssClass="bordagrid" Height="75px" ScrollBars="Auto" Width="100%">
                        <%--<asp:GridView ID="grdResultDev" runat="server"></asp:GridView>--%>
                        <asp:GridView ID="grdResultDev" runat="server" CellPadding="4" ForeColor="#333333"
                            GridLines="None" Width="100%" AutoGenerateColumns="False">
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#999999" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <Columns>
                                <asp:BoundField DataField="Quantidade" DataFormatString="{0:N2}" HeaderText="Qtd.">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Unitario" HeaderText="Unitario" DataFormatString="{0:N10}">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Valor" DataFormatString="{0:N2}" HeaderText="Valor">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                            </Columns>
                        </asp:GridView>
                    </asp:Panel>

                    <div class="row" id="divReajusteUnitario" runat="server">
                        <div class="coltxt">
                            <asp:CheckBox ID="chkReajusteUnitario" runat="server" Text="Reajuste de Unitário." Checked="true" AutoPostBack="true" ToolTip="Quando marcado listará apenas notas com quantidade fiscal e cujo unitário seja maior do que o informado no campo Unitário Médio. Caso não marcado serão listadas apenas notas de quantidade zero já que será uma devolução de valor." />
                        </div>
                        <div class="coltxt" id="divUnitarioMedio" runat="server">
                            <asp:TextBox ID="txtUnitarioMedio" runat="server" CssClass="txtDecimal10" AutoPostBack="true" />
                        </div>
                    </div>
                </div>
                <div class="painelleft" style="width: 20%;">
                    <div class="row">
                        <div class="collbl" style="margin-bottom: 5px;">
                            Origem
                        </div>
                        <div class="collbl" style="margin-bottom: 5px;">
                            Qtd. 
                        </div>
                        <div class="collbl">
                            Valor
                        </div>
                    </div>
                </div>
                <div class="painelleft" style="width: 20%;">
                    <div class="row">
                        <div class="collbl" style="margin-bottom: 5px;">
                            Nota Devolucao.
                        </div>

                        <div class="coltxt" style="width: 113px; margin-bottom: 5px;">
                            <asp:Label ID="lblQtdeNota" runat="server" Width="100px" />
                        </div>
                        <div class="coltxt" style="width: 113px;">
                            <asp:Label ID="lblVlrNota" runat="server" Width="100px" />
                        </div>

                    </div>
                </div>
                <div class="painelleft" style="width: 20%;">
                    <div class="row">
                        <div class="row">
                            <div class="collbl" style="width: 113px; margin-bottom: 5px;">
                                Falta Referenciar.
                            </div>
                            <div class="coltxt" style="width: 113px; margin-bottom: 5px;">
                                <asp:Label ID="lblQuantidade" runat="server" Width="100px" />
                            </div>
                            <div class="coltxt" style="width: 113px;">
                                <asp:Label ID="lblValor" runat="server" Width="100px" />
                            </div>
                        </div>
                    </div>
                </div>
            </div>



            <div class="row">
                <div class="painelleft" style="width: 100%;">
                    <asp:Panel ID="pnlNotaDeDevolucaoXNota" runat="server" CssClass="bordagrid" Height="350px" ScrollBars="Auto" Width="100%">
                        <asp:GridView ID="gridNotaDeDevolucaoXNota" runat="server" CellPadding="4" ForeColor="#333333"
                            GridLines="None" Width="100%" AutoGenerateColumns="False">
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <EditRowStyle BackColor="#999999" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <Columns>
                                <asp:BoundField DataField="NumeroNota" HeaderText="Nota">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Serie" HeaderText="Serie">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="DataDaNota" HeaderText="Data" DataFormatString="{0:d}">
                                    <HeaderStyle HorizontalAlign="Left" />
                                    <ItemStyle HorizontalAlign="Left" />
                                </asp:BoundField>
                                <asp:BoundField DataField="QuantidadeSaldo" HeaderText="Qtde Saldo" DataFormatString="{0:N4}">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="ValorSaldo" HeaderText="Vlr Saldo" DataFormatString="{0:N2}">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="UnitarioNota" HeaderText="Unitario" DataFormatString="{0:N9}">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="Qtde Dev.">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtQuantidade" runat="server" CssClass="txtDecimal4" Text='<%# String.Format("{0:N4}", Eval("QuantidadeDevolucao"))%>'
                                            AutoPostBack="True" OnTextChanged="txtQuantidade_TextChanged" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Vlr Dev.">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                    <ItemTemplate>
                                        <asp:TextBox ID="txtValor" runat="server" CssClass="txtDecimal2" Text='<%# Eval("ValorDevolucao", "{0:N2}") %>'
                                            AutoPostBack="True" Enabled="False" OnTextChanged="txtValor_TextChanged" />
                                        &nbsp;
                                        <asp:ImageButton ID="imgMenos" runat="server" Visible="false" Width="12px" Height="12px" ImageUrl="~/Images/ico-menos.gif" OnClick="imgMenos_Click" data-ToolTip="default" ToolTip="Diminuir 50(cinquenta) centavos." />
                                        &nbsp;
                                        <asp:ImageButton ID="imgMais" runat="server" Visible="false" Width="12px" Height="12px" ImageUrl="~/Images/ico-mais.gif" OnClick="imgMais_Click" data-ToolTip="default" ToolTip="Aumentar 50(cinquenta) centavos." />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="QuantidadeFixado" DataFormatString="{0:N2}" HeaderText="Qtde Fix.">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:BoundField DataField="ValorFixado" DataFormatString="{0:N2}" HeaderText="Vlr Fix.">
                                    <HeaderStyle HorizontalAlign="Right" />
                                    <ItemStyle HorizontalAlign="Right" />
                                </asp:BoundField>
                                <asp:TemplateField HeaderText="">
                                    <ItemTemplate>
                                        <asp:ImageButton ID="imgInfNota" runat="server"
                                            Width="16px" Height="16px" ImageUrl="~/images/important.png" Visible="false" data-ToolTip="default"
                                            ToolTip="Nota Fiscal com Título baixado não pode ser utilizada." />
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                    </asp:Panel>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
