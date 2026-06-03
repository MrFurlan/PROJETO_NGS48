<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucControleNumeroDeLote.ascx.vb" Inherits="NGS.Web.UI.ucControleNumeroDeLote" %>


<div id="divControleNumeroDeLote" class="uc" title="Controle do Número de Lote" style="display: none; text-align: center;">
    <asp:UpdatePanel ID="updpnlControleNumeroDeLote" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton ID="lnkConfirmar" class="iconConfirmar" runat="server" text="Confirmar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton ID="lnkAdicionar" class="iconNovo" runat="server" text="Adicionar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton ID="lnkLimpar" class="iconLimpar" runat="server" text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton ID="lnkFechar" class="iconSair" runat="server" text="Fechar" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="painelleft" style="width: 100%; margin-left: 4px; margin-right: 4px; align-items:center;">
                <div class="row">
                    <asp:Label id="seq" runat="server" style="width: 55px; padding-left: 10px; text-align: right;" />
                    <div class="collbl" style="width: 70px; text-align: left;">
                        Número:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtNumeroDoLote" runat="server" Style="color: Blue; Width:80px;" data-ToolTip="default" ToolTip="Número do Lote no máximo 20 caractéres." />
                    </div>
                    <div class="collbl" style="width: 70px; text-align: left;">
                        Fabricado:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataFabricadoLote" runat="server" CssClass="calendario" ClientIDMode="Static" Style="color: Blue; text-align: left; Width:80px;" data-ToolTip="default" ToolTip="Fabricação do Lote." />
                    </div>
                    <div class="collbl" style="width: 70px; text-align: left;">
                        Validade:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataValidadeLote" runat="server" CssClass="calendario" ClientIDMode="Static" Style="color: Blue; text-align: left; Width:80px;" data-ToolTip="default" ToolTip="Validade do Lote." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl" style="width: 85px; text-align: left;">
                        Quantidade:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtQuantidadeLote" runat="server" CssClass="txtDecimal4" Width="135px" Style="color: Blue; text-align: left" data-ToolTip="default" ToolTip="">0,0000</asp:TextBox>
                        &nbsp;         
                    </div>
                    <div class="collbl" style="width: 150px; text-align: left;">
                        Quantidade de Consumo:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtQuantidadeConsumo" runat="server" CssClass="txtDecimal4" Width="135px" Style="color: Blue; text-align: left" data-ToolTip="default" ToolTip="">0,0000</asp:TextBox>
                        &nbsp;         
                    </div>
                </div>
                <div class="row">
                    <div class="bordagrid" style="height: 195px;">
                        <asp:GridView ID="gridInfLote" runat="server" AutoGenerateColumns="False"
                             ForeColor="#333333" GridLines="None" Width="100%">
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <Columns>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                            <asp:LinkButton ID="lnkSelecionarLote" CssClass="lnk"
                                            data-tooltip="default" ToolTip="Selecionar Lote." runat="server" Text=" &gt; "
                                            OnClick="lnkSelecionarLote_Click">
                                            <i class="fa fa-arrow-right seta"></i>
                                        </asp:LinkButton>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Lote" HeaderText="Lote">
                                    <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    <ItemStyle HorizontalAlign="Left" Width="100px" VerticalAlign="Middle" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Fabricado" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Fabricado">
                                    <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    <ItemStyle HorizontalAlign="Left" Width="100px" VerticalAlign="Middle" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Validade" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Validade">
                                    <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    <ItemStyle HorizontalAlign="Left" Width="100px" VerticalAlign="Middle" />
                                </asp:BoundField>
                                <asp:BoundField DataField="Quantidade" DataFormatString="{0:N4}" HeaderText="Quantidade">
                                    <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                </asp:BoundField>
                                <asp:BoundField DataField="QuantidadeDeConsumo" DataFormatString="{0:N4}" HeaderText="Qtd.Consumo">
                                    <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                </asp:BoundField>
                                <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:ImageButton ID="imgRemoverLote" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                            onClick="imgRemoverLote_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente remover o Lote?');" />
                                    </ItemTemplate>
                                    <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
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
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
