<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucNFObsProduto.ascx.vb" Inherits="NGS.Web.UI.ucNFObsProduto" %>

<script type="text/javascript">
    function pageLoadNFObsProduto() {
        $('#MainContent_ucNFObsProduto_txtObservacoesDoProduto').attr('maxLength', 500);

        $("#txtDataFabricadoLote:enabled").datepicker({
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
            buttonImageOnly: true,
            showButtonPanel: true
        });

        $("#txtDataValidadeLote:enabled").datepicker({
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
            buttonImageOnly: true,
            showButtonPanel: true
        });

        $("#txtDataFabricadoLote").setMask("date");
        $("#txtDataValidadeLote").setMask("date");
    }

    var prmNFObsProduto = Sys.WebForms.PageRequestManager.getInstance();
    prmNFObsProduto.add_endRequest(pageLoadNFObsProduto);
</script>

<div id="divNFObsProduto" class="uc" title="Observação do Produto" style="display: none;">
    <asp:UpdatePanel ID="updpnlNFObsProduto" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="txtIndiceDoProduto" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li id="liNovo" class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconFechar" runat="server">
                            <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="bordagrid" style="height: auto;">
                <div class="row">
                    <div class="painelleft" style="width: 49%; margin-left: 4px; margin-right: 4px;">
                        <div class="subtitulodiv">
                            <label>
                                Observação do Produto
                            </label>
                        </div>
                        <div class="row">
                            <div class="coltxt lg">
                                <asp:TextBox ID="txtObservacoesDoProduto" runat="server" Style="height: 138px; width: 97%;"
                                    data-ToolTip="default" ToolTip="Digite as observações do produto em seguida clique em Gravar para atualizar (Máximo de 500 Caracteres.)"
                                    TextMode="MultiLine" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:Label ID="lblImpostos" runat="server" />
                            </div>
                        </div>
                    </div>
                    <div class="painelleft" style="width: 49%;">
                        <asp:Panel ID="pnlDevolucao" runat="server">
                            <div class="subtitulodiv">
                                <label>
                                    Mensagem de Devolução
                                </label>
                            </div>
                            <div class="row">
                                <div class="coltxt lg">
                                    <asp:TextBox ID="txtMsgDevolucaoProduto" runat="server" Style="height: 138px; width: 98%;"
                                        TextMode="MultiLine" ReadOnly="true" />
                                </div>
                            </div>
                        </asp:Panel>
                    </div>
                </div>
                <div class="row">
                    <div id="divInfLote" runat="server">
                        <div class="painelleft" style="width: 80%; margin-left: 4px; margin-right: 4px;">
                            <div class="subtitulodiv">
                                <label>
                                    Informações do Lote
                                </label>
                            </div>
                            <div class="row">
                                <div class="collbl" style="width: 55px;">
                                    Número:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNumeroDoLote" runat="server" Width="100px" Enabled="false" MaxLength="20" data-ToolTip="default" ToolTip="Número do Lote no máximo 20 caractéres." />
                                    <asp:ImageButton ID="imgSelecionaLote" runat="server" Visible="false" ImageUrl="~/images/search.png" ToolTip="Buscar Número(s) do(s) Lote(s)." OnClick="imgSelecionaLote_Click" />
                                </div>
                                <div class="collbl" style="width: 55px;">
                                    Fabricado:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataFabricadoLote" runat="server" ClientIDMode="Static" Width="60px" Enabled="false" data-ToolTip="default" ToolTip="Fabricação do Lote." />
                                </div>
                                <div class="collbl" style="width: 55px;">
                                    Validade:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataValidadeLote" runat="server" ClientIDMode="Static" Width="60px" Enabled="false" data-ToolTip="default" ToolTip="Validade do Lote." />
                                </div>
                                <div class="collbl" style="width: 70px;">
                                    Quantidade:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtQuantidadeLote" runat="server" CssClass="txtDecimal4" Width="100px" Enabled="false" Style="color: Blue; text-align: right" data-ToolTip="default" ToolTip="">0,0000</asp:TextBox>
                                    &nbsp;
                                    <asp:ImageButton ID="btnAdicionarInfLote" runat="server" CssClass="btn" Visible="false" OnClick="btnAdicionarInfLote_Click" ImageUrl="~/images/ico-mais.gif" Width="17px" data-ToolTip="default" ToolTip="Adiciona informações do Lote" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="bordagrid" style="height: 195px;">
                                    <asp:GridView ID="gridInfLote" runat="server" AutoGenerateColumns="False"
                                        CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:CommandField InsertText="" SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                                <ItemStyle Width="30px" />
                                            </asp:CommandField>
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
                                            <asp:TemplateField>
                                                <ItemTemplate>
                                                    <asp:ImageButton ID="imgExcluirLote" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                        OnClick="imgExcluirLote_Click" data-ToolTip="default" Visible="false" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente excluir o Lote?');" />
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
                    </div>
                    <asp:Panel ID="pnlPecas" runat="server">
                        <div class="painelleft" style="width: 38%;">
                            <div class="collbl">
                                Peças/Meios:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPecas" runat="server" Style="text-align: right" Width="65px" />
                            </div>
                        </div>
                    </asp:Panel>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
