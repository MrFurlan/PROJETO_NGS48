<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="BancosXContas.aspx.vb" Inherits="NGS.Web.UI.BancosXContas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function preencheChequeAtual(txt) {
            var thisctrl = $('#' + txt.id);
            var destinoctrl = $('#MainContent_txtNumChequeAtual');
            if ($.trim(thisctrl.val()) != "" && $.trim(thisctrl.val()) != "0,00" && thisctrl.val() != undefined) {
                destinoctrl.val('0');
            }
        }
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngBancosXContas" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True">
    </asp:ScriptManager>
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlBancosXContas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Bancos X Contas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" runat="server" Text="Gravar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server" Text="Atualizar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" runat="server" Text="Excluir"
                                OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" runat="server" Text="Relatório" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlClienteDadosBancarios" runat="server" Width="627px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Banco
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlBancos" runat="server" Width="627px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cheque Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumChequeInicial" CssClass="txtNumerico" runat="server" Style="width: 113px;"
                        onchange="preencheChequeAtual(this);" />
                </div>
                <div class="collbl">
                    Cheque Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumChequeFinal" CssClass="txtNumerico" runat="server" Style="width: 113px;" />
                </div>
                <div class="collbl">
                    Cheque Atual:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNumChequeAtual" CssClass="txtNumerico" runat="server" Style="width: 113px;
                        margin-right: 3px;" Enabled="false" />
                    <asp:Button ID="btnInutilizarCheques" runat="server" Text="Inutilizar" CssClass="botao"
                        Width="89px" Visible="false" data-ToolTip="default" ToolTip="Inutiliza cheques a serem emitidos" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Agência:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtAgencia" runat="server" Width="80px" MaxLength="4" data-ToolTip="default"
                        ToolTip="Informar o número da agência bancária." />
                    <asp:TextBox ID="txtDigitoAgencia" runat="server" Width="20px" data-ToolTip="default"
                        ToolTip="Informar o número da agência bancária." />
                </div>
                <div class="collbl">
                    Conta:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtContaCorrente" runat="server" Width="80px" data-ToolTip="default"
                        ToolTip="Informar a conta bancária." />
                    <asp:TextBox ID="txtDigitoDaConta" runat="server" Width="20px" data-ToolTip="default"
                        ToolTip="Informar a conta bancária." />
                </div>
                <div class="collbl">
                    <asp:CheckBox ID="chkFluxoDeCaixa" runat="server" Text="Fluxo de Caixa:" data-ToolTip="default"
                        ToolTip="Marcar quando hover fluxo de caixa." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Tipo Conta:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlTipoConta" runat="server" Width="123px">
                        <asp:ListItem Value="C">C. Corrente</asp:ListItem>
                        <asp:ListItem Value="P">C. Poupança</asp:ListItem>
                        <asp:ListItem Value="A">C. Aplicação</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="collbl">
                    Ativo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlAtivo" runat="server" Width="123px">
                        <asp:ListItem Value="1">Sim</asp:ListItem>
                        <asp:ListItem Value="0">Não</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="collbl">
                    Limite Bancario:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtLimiteBanc" runat="server" CssClass="txtDecimal" Width="113px"
                        data-ToolTip="default" ToolTip="Informar o limite bancário da conta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    C.Contábil:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlGrupoDeContas" runat="server" Width="627px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Observações:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtObservacoesDaConta" runat="server" Width="618px" TextMode="MultiLine"
                        MaxLength="200" data-ToolTip="default" ToolTip="Preencher com observações quando necessário." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridContasCorrentes" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridContasCorrentes_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                        </asp:CommandField>
                        <asp:BoundField DataField="Banco_Id" HeaderText="Banco" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Agencia_Id" HeaderText="Ag&#234;ncia">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DigitoAgencia_Id" HeaderText="Dg" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Conta_id" HeaderText="Conta">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="DigitoConta_Id" HeaderText="Dg" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="TipoConta" HeaderText="Tipo">
                            <HeaderStyle HorizontalAlign="Center" />
                            <ItemStyle HorizontalAlign="Center" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Observacoes" HeaderText="Observa&#231;&#245;es" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="ContaContabil" HeaderText="Conta Contábil">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Empresa_id" HeaderText="Empresa">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="EndEmpresa_id" HeaderText="End">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Reduzido" HeaderText="Reduzido"></asp:BoundField>
                        <asp:BoundField DataField="Ativo" HeaderText="Ativo"></asp:BoundField>
                        <asp:BoundField DataField="FluxoDeCaixa" HeaderText="Fluxo Caixa">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="LimiteBancario" HeaderText="Limite Bancário">
                            <HeaderStyle HorizontalAlign="right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="right"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:CancelamentoDeCheque ID="ucCancelamentoDeCheque" runat="server" />
</asp:Content>
