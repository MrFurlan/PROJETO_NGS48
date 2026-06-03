<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="RateioDePesagem.aspx.vb" Inherits="NGS.Web.UI.RateioDePesagem" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }

        .txtNumerico {
            text-align: right;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngRateio" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlRateio" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Rateio De Pesagem
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" Style="margin-top: 4px;"
                ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="tabRateio" runat="server">
                    <HeaderTemplate>
                        Rateio
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconExcluir" runat="server">
                                        <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkRelatorio" Text="Reemprimir" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                    <li style="float: right;">
                                        <div class="row" style="margin-top: 0;">
                                            <div class="coltxt">
                                                Cópias:
                                            </div>
                                            <div class="coltxt">
                                                <asp:DropDownList ID="DdlCopias" runat="server" Enabled="False" Style="width: 80px;">
                                                    <asp:ListItem>1</asp:ListItem>
                                                    <asp:ListItem>2</asp:ListItem>
                                                </asp:DropDownList>
                                            </div>
                                        </div>
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 61%; margin-right: 2px;">
                            <div class="row">
                                <div class="collbl">
                                    Empresa:
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="txtCodigoEmpresa" runat="server" />
                                    <asp:TextBox ID="txtEmpresa" runat="server" Width="580px" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnEmpresa" runat="server" OnClick="btnEmpresa_Click" CssClass="btn"
                                        Text="&gt;" Enabled="False" UseSubmitBehavior="False" data-ToolTip="default"
                                        ToolTip="" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Laudo:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtLaudo" runat="server" CssClass="txtNumerico" Enabled="False" Width="105px" />
                                </div>
                                <div class="collbl">
                                    Romaneio:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtRomaneio" runat="server" CssClass="txtNumerico" Enabled="False" Width="105px" />
                                </div>
                                <div class="collbl">
                                    Movimento:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtMovimento" runat="server" Width="93px" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:CheckBox ID="chkFisico" runat="server" Enabled="False" Text="Físico" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Cliente:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCliente" runat="server" Width="581px" Enabled="False" />
                                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Depósito:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDeposito" runat="server" Width="581px" Enabled="False" />
                                    <asp:HiddenField ID="txtCodigoDeposito" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Pedido:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPedido" runat="server" CssClass="txtNumerico" Enabled="False" Width="100px" />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgExtratoPedido" CssClass="imgconsultar" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                        data-ToolTip="default" ToolTip="Informar o número do pedido." />
                                </div>
                                <asp:HiddenField ID="txtCodigoAutorizacao" runat="server" />
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Operação:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtOperacao" runat="server" Width="33px" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtSubOperacao" runat="server" Width="31px" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDescricaoOperacao" runat="server" Width="493px" Enabled="False" />
                                    <asp:HiddenField ID="txtCodigoOperacao" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Produto:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtProduto" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDescricaoProduto" runat="server" Width="472px" Enabled="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Classificação:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtClassificacao" runat="server" Width="32px" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDescricaoClassificacao" runat="server" Width="540px" Enabled="False" />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft" style="width: 38%; margin-right: 1px;">
                            <div id="divAnalises" runat="server" visible="False">
                                <div class="subtitulodiv">
                                    Análises
                                </div>
                                <div class="bordagrid" style="height: 185px;">
                                    <asp:GridView ID="GridDescontos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                        ForeColor="#333333" GridLines="None" Width="100%">
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <EditRowStyle BackColor="#999999" />
                                        <Columns>
                                            <asp:ButtonField Text=" &gt; " ButtonType="Button"></asp:ButtonField>
                                            <asp:BoundField DataField="CodigoAnalise" HeaderText="Codigo">
                                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Percentual" HeaderText="Percentual" DataFormatString="{0:N}">
                                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Indice" HeaderText="Indice" DataFormatString="{0:N}">
                                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Desconto" HeaderText="Desconto" DataFormatString="{0:N0}">
                                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                                <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                            </asp:BoundField>
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </div>
                        </div>
                        <div id="divRateio" runat="server" visible="False">
                            <div class="subtitulodiv">
                                Rateios
                            </div>
                            <div class="menu_acoes" id="divMenuRateio" runat="server" visible="False">
                                <div class="acoes">
                                    <ul>
                                        <li class="iconNovo" runat="server">
                                            <asp:LinkButton ID="lnkConfirmar" Text="Confirmar" runat="server" />
                                        </li>
                                        <li class="iconAtualizar" runat="server">
                                            <asp:LinkButton ID="lnkFinalizar" Text="Finalizar" runat="server" />
                                        </li>
                                        <li class="iconExcluir" runat="server">
                                            <asp:LinkButton ID="lnkExcluirRateio" Text="Excluir" runat="server" />
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Destinatário:
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="txtCodigoDestinatario" runat="server" />
                                    <asp:HiddenField ID="txtPedidoSaldo" runat="server" />
                                    <asp:HiddenField ID="txtPedidoEntregue" runat="server" />
                                    <asp:HiddenField ID="txtCodigoProdutoPedido" runat="server" />
                                    <asp:HiddenField ID="txtCodigoAutorizacaoDestinatario" runat="server" />
                                    <asp:TextBox ID="txtDestinatario" runat="server" Width="535px" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="cmdCliente" runat="server" OnClick="cmdCliente_Click" Text=" > "
                                        CssClass="btn" Enabled="False" UseSubmitBehavior="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Pedido:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtPedidoDestinatario" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgExtratoPedidoDestinatario" CssClass="imgconsultar" runat="server"
                                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" data-ToolTip="default"
                                        ToolTip="Visualizar Extrato" />
                                </div>
                                <div class="collbl">
                                    Saldo:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtSaldoPedido" runat="server" Enabled="False" Style="width: 100px;" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Depósito:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlDepositoDestinatario" runat="server" Width="577px" Enabled="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Operação:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlOperacao" runat="server" Width="577px" Enabled="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Bruto:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtBruto" runat="server" Enabled="False" Style="width: 100px;" />
                                </div>
                                <div class="collbl">
                                    Líquido:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtLiquido" runat="server" Enabled="False" Style="width: 100px;" />
                                </div>
                                <div class="collbl">
                                    Quantidade:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtQuantidade" CssClass="txtNumerico" runat="server" Enabled="False"
                                        Style="width: 90px;" />
                                </div>
                            </div>
                            <div class="bordagrid" style="height: 100px;">
                                <asp:GridView ID="GridRateio" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:CommandField>
                                        <asp:BoundField DataField="Codigo" HeaderText="Cliente">
                                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Pedido" HeaderText="Pedido">
                                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Deposito" HeaderText="Depósito">
                                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Produto" HeaderText="Produto">
                                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Op" HeaderText="OP">
                                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="SO" HeaderText="SO">
                                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Autorizacao" HeaderText="Autorização">
                                            <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Bruto" DataFormatString="{0:N0}" HeaderText="Bruto">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Desconto" DataFormatString="{0:N0}" HeaderText="Desconto">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Liquido" DataFormatString="{0:N0}" HeaderText="Líquido">
                                            <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                            <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        </asp:BoundField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:ImageButton ID="btnExcluirItemRateio" runat="server" ImageUrl="~/Images/deletar.gif"
                                                    Style="border: 0;" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Confirma a exclusão deste item?');"
                                                    OnClick="btnExcluirItemRateio_Click" />
                                            </ItemTemplate>
                                            <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                            <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="tabAlterar" runat="server">
                    <HeaderTemplate>
                        Alterar Rateio
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultarAR" Text="Consultar Laudo" runat="server" />
                                    </li>
                                    <li class="iconNovo" runat="server">
                                        <asp:LinkButton ID="lnkConsultarNF" Text="Consultar NF" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparAR" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjudaAR" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoEmpresaAlt" runat="server" />
                                <asp:TextBox ID="txtEmpresaAlt" runat="server" Width="580px" Enabled="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnEmpresaAlt" runat="server" Text=" > " CssClass="btn" UseSubmitBehavior="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Laudo:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtLaudoAlt" CssClass="txtNumerico" Width="130px" runat="server" />
                            </div>
                            <div class="collbl">
                                Nota Fiscal:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNotaAlt" CssClass="txtNumerico" Width="130px" runat="server" />
                            </div>
                            <div class="collbl">
                                Movimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMovimentoAlt" runat="server" Width="88px" Enabled="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtClienteAlt" runat="server" Width="580px" Enabled="False">
                                </asp:TextBox><asp:HiddenField ID="txtCodigoClienteAlt" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedidoAlt" Width="130px" runat="server" Enabled="False" />
                                <asp:HiddenField ID="txtCodigoAutorizacaoAlt" runat="server" />
                                <asp:HiddenField ID="txtCodigoOperacaoAlt" runat="server" />
                                <asp:HiddenField ID="txtProdutoAlt" runat="server" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 100px;">
                            <asp:GridView ID="gridRomaneios" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                        <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Romaneio" HeaderText="Romaneio">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Cliente">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Pedido" HeaderText="Pedido">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Deposito" HeaderText="Depósito">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Produto" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Op" HeaderText="OP">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="SO" HeaderText="SO">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Autorizacao" HeaderText="Autorização">
                                        <HeaderStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Bruto" DataFormatString="{0:N0}" HeaderText="Bruto">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Desconto" DataFormatString="{0:N0}" HeaderText="Desconto">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Liquido" DataFormatString="{0:N0}" HeaderText="Líquido">
                                        <HeaderStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                        <ItemStyle HorizontalAlign="Right" VerticalAlign="Middle" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgRemoverRomaneio" runat="server" Visible="false" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                OnClick="imgRemoverRomaneio_Click" data-ToolTip="default" ToolTip="Excluir Romaneio" OnClientClick="return confirm('Deseja realmente remover o Romaneio?');" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox Width="530px" ID="txtClienteSel" Enabled="False" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnSelCliente" Text=">" runat="server" Enabled="False" CssClass="btn" />
                                <asp:HiddenField ID="txtCodigoClienteSel" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnAltRomaneio" Enabled="False" runat="server" Text="Alterar" CssClass="botao" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedidoSel" Enabled="False" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Depósito:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlDepositoDestinatarioSel" runat="server" Width="572px" Enabled="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Operação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlOperacaoSel" runat="server" Width="572px" Enabled="False"
                                    AutoPostBack="True">
                                </asp:DropDownList>
                                <asp:HiddenField ID="txtCodigoAutorizacaoAltSel" runat="server" />
                                <asp:HiddenField ID="txtPedidoSaldoSel" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Nota Fiscal:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNotaFiscalAlt" Enabled="False" CssClass="txtNumerico" Width="130px" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnBuscarNF" Enabled="False" runat="server" Text="Buscar NF" CssClass="botao" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnGravarNF" Visible="false" runat="server" Text="Gravar NF" CssClass="botao" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaEmpresas ID="ucConsultaEmpresas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaAutorizacaoDeRetirada ID="ucConsultaAutorizacaoDeRetirada" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:ConsultaPedidosXNotas ID="ucConsultaPedidosXNotas" runat="server" />
</asp:Content>
