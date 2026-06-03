<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Pagamentos.aspx.vb" Inherits="NGS.Web.UI.Pagamentos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function partyping() {
            $("#MainContent_cmdParcelas").click()
        };
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngPagamentos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPagamentos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Condições de Pagamentos
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" runat="server">
                                <span>Gravar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkAtualizar" runat="server">
                                <span>Atualizar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;">
                                <span>Excluir</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server">
                                <span>Limpar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" runat="server">
                                <span>Relatório</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" runat="server">
                                <span>Ajuda</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Código:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" TabIndex="1" runat="server" Width="66px" data-ToolTip="default"
                        ToolTip="Código de cadastro das condições de pagamento." Enabled="false" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Descrição:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" TabIndex="2" MaxLength="100" runat="server" Width="600px" data-ToolTip="default"
                        ToolTip="Descrição da condição de pagamento." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Qtd. Parcelas:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtParcelas" class="txtNumerico" onkeyup="setTimeout(partyping, 500);" TabIndex="3" runat="server" Width="56px" data-ToolTip="default"
                        ToolTip="Quantidade de pagamentos." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkAVista" runat="server" Text="Considerado À Vista" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkAntecipado" runat="server" Text="Pagamento Antecipado" />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkVencimentoPedido" runat="server" Text="Vencimento do Pedido" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt">
                    <asp:Button ID="cmdParcelas" TabIndex="4" OnClick="cmdParcelas_Click" runat="server"
                        Text="Prazos >>" CssClass="btn" />
                </div>
            </div>
            <div class="painelleft" style="width: 775px;">
                <div class="bordagrid">
                    <asp:GridView ID="GridPagamentos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridOperacoes_SelectedIndexChanged">
                        <EditRowStyle BackColor="#999999" />
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"
                                ItemStyle-HorizontalAlign="Center"></asp:CommandField>
                            <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo" ItemStyle-HorizontalAlign="Center">
                            </asp:BoundField>
                            <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o" ItemStyle-HorizontalAlign="Center">
                            </asp:BoundField>
                            <asp:BoundField DataField="Parcelas" HeaderText="Parcelas" ItemStyle-HorizontalAlign="Center">
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="À Vista">
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                <EditItemTemplate>
                                    <asp:CheckBox ID="CheckBox1" runat="server" Checked='<%# Bind("AVista") %>' />
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="CheckBox1" runat="server" Checked='<%# Bind("AVista") %>' Enabled="false" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Pag. Ant.">
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                <EditItemTemplate>
                                    <asp:CheckBox ID="CheckBox2" runat="server" Checked='<%# Bind("Antecipado") %>' />
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="CheckBox2" runat="server" Checked='<%# Bind("Antecipado") %>' Enabled="false" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Venc.Pedido">
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                <EditItemTemplate>
                                    <asp:CheckBox ID="CheckBox3" runat="server" Checked='<%# Bind("VencimentoPedido") %>' />
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="CheckBox3" runat="server" Checked='<%# Bind("VencimentoPedido") %>' Enabled="false" />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
            <div class="painelright" style="width: 165px;">
                <div class="bordagrid">
                    <asp:GridView ID="GridParcelas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        ForeColor="#333333" GridLines="None" Width="100%">
                        <EditRowStyle BackColor="#999999" />
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                        <Columns>
                            <asp:BoundField DataField="Sequencia" HeaderText="Ordem">
                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="Dias">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtDias" runat="server" CssClass="txtNumerico" Width="90%" />
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
