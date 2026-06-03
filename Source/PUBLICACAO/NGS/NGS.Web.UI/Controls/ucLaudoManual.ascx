<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucLaudoManual.ascx.vb" Inherits="NGS.Web.UI.ucLaudoManual" %>
<script type="text/javascript">
    $("#txtDataDeEmissao").setMask("date");
    $("#txtDataDeEntrada").setMask("date");
</script>

<div id="divLaudoManual" class="uc" title="Laudo Manual" style="display: none;">
    <asp:UpdatePanel ID="updpnlLaudoManual" runat="server">
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
                        <li class="iconLimpar" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server">
                                <span>Limpar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconSair" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkFechar" runat="server">
                                <span>Fechar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconAjuda" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server">
                                <span>Ajuda</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupoProdutoProducao" runat="server" AutoPostBack="True" Width="610px" />
                </div>
                <div class="coltxt">
                    <asp:LinkButton ID="lnkBuscaProduto" runat="server" Height="20px" Width="20px"><asp:Image ID="imgSearch" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                            ToolTip="Consulta Produto" /></asp:LinkButton>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProdutos" runat="server" AutoPostBack="True" Width="630px" />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    N° do lote:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNLote" CssClass="txtNumerico" runat="server" data-tooltip="default" ToolTip="Número de lote." />
                </div>
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtData" CssClass="calendario" runat="server"
                        Width="100px" data-ToolTip="default" ToolTip="Informar a data." />
                </div>
                <div class="collbl">
                    Validade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtVencimento" CssClass="calendario" runat="server"
                        Width="100px" data-ToolTip="default" ToolTip="Informar o período de validade." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Especificação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEspecificacao" runat="server" Width="630px" data-ToolTip="default" ToolTip="Código da Especificação" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Faixa inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFaixaInicial" runat="server" CssClass="txtDecimal2" Width="225px" data-ToolTip="default" ToolTip="Faixa Inicial." />
                </div>
                <div class="collbl">
                    Faixa final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFaixaFinal" runat="server" CssClass="txtDecimal2" Width="225px" data-ToolTip="default" ToolTip="Faixa Final." />
                </div>
                <div class="coltxt">
                    <asp:Button ID="lnkNovaEspecificacao" runat="server" width="25px" Height="25px" Text="+ " />
                </div>
            </div>
            <div class="subtitulodiv">
                Especificações do Produto
            </div>
            <div class="bordagrid" style="height: 200px;">
                <asp:GridView ID="gridEspecificacaoDoProduto" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="Codigo" HeaderText="Código">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Especificação">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Referencia" Visible="true">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                            <ItemTemplate>
                                <asp:TextBox ID="txtReferencia" runat="server" Text='<%# Eval("Referencia")%>'/>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Resultado" Visible="true">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                            <ItemTemplate>
                                <asp:TextBox ID="txtResultado" runat="server" Text='<%# Eval("ResultadoTxt")%>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:ImageButton ID="imgRemoverEspecificacao" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                    onClick="imgRemoverEspecificacao_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente remover a Especificação?');" />
                            </ItemTemplate>
                            <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                            <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>