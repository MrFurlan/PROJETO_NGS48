<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="FrmTitulo.aspx.vb" Inherits="NGS.Web.UI.FrmTitulo" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1080px !important;
        }

        .none {
            display: none;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmTitulo" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updFrmTitulo" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="txtLiberarPedido" runat="server" />
            <div class="titulodiv">
                <asp:Label ID="LblRecPag" runat="server" Text="Contas A Receber" />
            </div>
            <ajaxToolkit:TabContainer ID="tbcFrmTitulo" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="tabTitulo" runat="server">
                    <HeaderTemplate>
                        Título
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovo" runat="server" Text="Gravar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconExcluir" ID="lnkExcluir" runat="server" OnClick="lnkExcluir_Click"
                                            Text="Excluir" OnClientClick="return confirm('Deseja excluir o registro selecionado');" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" runat="server" Text="Consultar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" Text="Limpar" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRecibo" runat="server" Text="Recibo" />
                                    </li>
                                    <li class="iconRelatorio" runat="server">
                                        <asp:LinkButton ID="lnkDuplicata" runat="server" Text="Duplicata" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Lançamento:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RdReceber" runat="server" AutoPostBack="True" GroupName="XRP"
                                    OnCheckedChanged="RdReceber_CheckedChanged" Text="Receber" />
                                <asp:RadioButton ID="RdPagar" runat="server" AutoPostBack="True" GroupName="XRP"
                                    OnCheckedChanged="RdPagar_CheckedChanged" Text="Pagar" />
                                <asp:RadioButton ID="RdContabil" runat="server" AutoPostBack="True" GroupName="XRP"
                                    OnCheckedChanged="RdContabil_CheckedChanged" Text="Contábil" />
                                &nbsp;
                                <asp:Label ID="lblAcaoTitulo" ForeColor="Red" runat="server" Text="  Título Bloqueado!"
                                    Visible="False" />
                                <asp:CheckBox runat="server" ID="chkManterLancamento" Text="Manter dados do lançamento" />
                                <asp:CheckBox runat="server" ID="chkEmitirRecibo" Text="Emitir Recibo" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Registro:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistro" CssClass="txtNumerico" runat="server" ClientIDMode="Static"
                                    Width="90px" Font-Bold="True" AutoPostBack="True" />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgBloqueio" OnClick="imgBloqueio_Click" runat="server" Width="25px"
                                    Height="25px" ImageUrl="~/Images/liberar.png" ImageAlign="AbsMiddle" data-ToolTip="default"
                                    ToolTip="Liberar Registro" />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="txtMestre" runat="server" CssClass="primario" />
                            </div>
                            <div class="collbl">
                                Movimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMovimento" TabIndex="1" runat="server" Width="90px" Font-Bold="False"
                                    CssClass="calendario" />
                            </div>
                            <div class="collbl">
                                Borderô:
                            </div>
                            <div class="coltxt">
                                <asp:Panel ID="pnlBordero" runat="server" CssClass="primario" Visible="False">
                                    <asp:TextBox ID="txtBordero" runat="server" Width="90px" />
                                    <asp:Label ID="lblBorderoMestre" runat="server" CssClass="primario" />
                                </asp:Panel>
                                <asp:Panel ID="pnlBorderoDuplicatasImpressao" runat="server" Visible="False">
                                    <asp:Button ID="btnImprimirBordero" runat="server" Text="Imprimir Bordero" CssClass="botao" />
                                    <asp:CheckBox ID="chkImprimirDuplicatas" runat="server" Text="Imprimir Duplicatas" />
                                    <asp:Button ID="btnRecomprar" runat="server" Text="Recomprar" CssClass="botao" />
                                </asp:Panel>
                            </div>
                        </div>
                        <div class="row" runat="server">
                            <div class="collbl">
                                Lote:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlLote" runat="server" Width="110px" AutoPostBack="True" OnSelectedIndexChanged="ddlLote_SelectedIndexChanged" />
                            </div>
                        </div>
                        <div class="row" runat="server">
                            <div class="collbl">
                                Sequência:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSequencia" runat="server" Width="110px" />
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            <asp:Label ID="LblCliForTit" runat="server" Text="Clientes" />
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlUnidadeDeNegocioEmpresaCliente" TabIndex="3" runat="server"
                                    Width="625px" OnSelectedIndexChanged="DdlUnidadeDeNegocioEmpresaCliente_SelectedIndexChanged"
                                    AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaCliente" TabIndex="4" runat="server" Width="625px"
                                    AutoPostBack="True" OnSelectedIndexChanged="DdlEmpresaCliente_SelectedIndexChanged" />
                            </div>
                        </div>
                        <div class="row" style="display: none;">
                            <div class="collbl">
                                Conta Contábil:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtEmpresa" runat="server" Visible="false" ClientIDMode="Static" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conta Contábil:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="hdnCodigoContaContabilCliFor" runat="server" />
                                <asp:TextBox ID="txtContaContabilCliFor" runat="server" Width="585px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnPlanoDeContas" runat="server" Text=" > " UseSubmitBehavior="False"
                                    CssClass="btn" />
                            </div>
                        </div>
                        <div class="row" runat="server">
                            <div class="collbl">
                                <asp:Label ID="lblCliFor" runat="server" Text="Cliente" />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCliente" runat="server" Width="585px" />
                                <asp:Button ID="btnClientesTitulo" runat="server" OnClick="btnClientesTitulo_Click"
                                    Text=" &gt; " UseSubmitBehavior="False" CssClass="btn" />
                            </div>
                        </div>
                        <div class="row" runat="server">
                            <div class="collbl">
                                Depósito:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDepositoCliente" runat="server" Width="600px" />
                                <asp:Button ID="btnDepositoCliente" runat="server" Text=" &gt; " UseSubmitBehavior="False"
                                    CssClass="btn" />
                            </div>
                        </div>
                        <div id="divAdiantamento" class="row" runat="server">
                            <div class="collbl">
                                Vencimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVencimentoAdiantamento" CssClass="calendario" runat="server"
                                    AutoPostBack="True" ClientIDMode="Static" OnTextChanged="txtVencimentoAdiantamento_TextChanged" />
                            </div>
                            <div class="collbl">
                                Taxa Juro A.M:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTaxaJuroAdiantamento" runat="server" AutoPostBack="True" CssClass="txtDecimal" />
                            </div>
                        </div>
                        <div id="divBanco" runat="server" class="row">
                            <div class="collbl">
                                Banco:
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnDadosBancarios" runat="server" OnClick="btnDadosBancarios_Click"
                                    Text=" &gt; " UseSubmitBehavior="False" CssClass="btn" />
                                &nbsp;
                                <asp:Label ID="lblBanco" runat="server" Text="Banco" />
                                &nbsp;
                                <asp:Label ID="lblAgencia" runat="server" Text="Agência" />
                                &nbsp;
                                <asp:Label ID="lblContaCorrente" runat="server" Text="Conta" />
                            </div>
                        </div>
                        <div id="divPedido" runat="server" class="row">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedido" runat="server" Enabled="False" Width="147px" />
                                <asp:Button ID="cmdPedido" OnClick="cmdPedido_Click" runat="server" UseSubmitBehavior="False"
                                    CssClass="btn" Text=" > " />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imbLimparPedido" runat="server" Height="22px" ImageAlign="AbsMiddle"
                                    ImageUrl="~/images/borracha.JPG" Width="22px" CssClass="btn" />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgExtratoPedido" OnClick="imgExtratoPedido_Click" runat="server"
                                    ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ImageAlign="AbsMiddle"
                                    data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido" />
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            <asp:Label ID="lblEmpresaRecPag" runat="server" Text="Empresa Recebedora" />
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUnidadeDeNegocioRecPag" TabIndex="3" runat="server" Width="625px"
                                    AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaRecPag" TabIndex="6" runat="server" Width="625px"
                                    OnSelectedIndexChanged="DdlEmpresaRecPag_SelectedIndexChanged" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row" runat="server">
                            <div class="collbl">
                                Depósito:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDepositoRecPag" runat="server" Width="600px" />
                                <asp:Button ID="btnDepositoRecPag" runat="server" Text=" &gt; " UseSubmitBehavior="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Conta Banco:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlContaContabilEmpresaRecPag" runat="server" Width="518px"
                                    AutoPostBack="True" OnSelectedIndexChanged="ddlContaContabilEmpresaRecPag_SelectedIndexChanged" />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtContaContabilRecPag" runat="server" Width="450px" Visible="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnPlanoDeContaRecPag" runat="server" Text=" > " OnClick="btnPlanoDeContas_Click"
                                    CssClass="btn" UseSubmitBehavior="False" Visible="False" />
                            </div>
                            <div runat="server">
                                <div class="collbl">
                                    <asp:Label ID="lblCotacao" runat="server" Text="Cotação" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCotacao" runat="server" AutoPostBack="True" CssClass="txtDecimal8"
                                        OnTextChanged="txtCotacao_TextChanged" />
                                </div>
                            </div>
                        </div>
                        <div class="row" runat="server">
                            <div class="collbl">
                                Cliente
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtClienteLancContabil" runat="server" Width="585px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnClienteLancamentoContabil" CssClass="btn" runat="server" OnClick="btnClienteLancamentoContabil_Click"
                                    Text=" > " UseSubmitBehavior="False" />
                            </div>
                        </div>
                        <div id="divTipoDeRecebimento" class="row" runat="server">
                            <div class="collbl">
                                <asp:Label ID="lblTipoRecPag" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlTiposDeRecebimentos" TabIndex="7" runat="server" Width="518px"
                                    AutoPostBack="True" OnSelectedIndexChanged="DdlTiposDeRecebimentos_SelectedIndexChanged" />
                            </div>
                            <div class="collbl">
                                Moeda:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlMoeda" runat="server" Width="160px" OnSelectedIndexChanged="ddlMoeda_SelectedIndexChanged"
                                    AutoPostBack="True" />
                            </div>
                        </div>
                        <div id="divCarteiraDoTitulo" runat="server" class="row">
                            <div class="collbl">
                                Carteira do Título
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCarteiraDoTitulo" runat="server" Width="370px" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlCarteiraDoTitulo_SelectedIndexChanged" />
                            </div>
                            <div class="collbl" style="margin-left: 148px;">
                                Indexador:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlIndexador" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlIndexador_SelectedIndexChanged"
                                    Width="160px" />
                            </div>
                        </div>
                        <div id="divPrevisaoBaixa" runat="server" class="row">
                            <div class="collbl">
                                Previsão Baixa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlProvisoes" TabIndex="9" runat="server" Width="370px" AutoPostBack="True"
                                    OnSelectedIndexChanged="DdlProvisoes_SelectedIndexChanged" />
                            </div>
                            <div class="collbl" style="margin-left: 148px;">
                                Venc. Atual:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtProrrogacao" ClientIDMode="Static" TabIndex="13" runat="server"
                                    data-ToolTip="default" ToolTip="Informe a Data de Vencimento" AutoPostBack="True"
                                    OnTextChanged="txtProrrogacao_TextChanged" CssClass="calendario" />
                            </div>
                        </div>
                        <div class="row">
                            <div id="divVencOriginal" runat="server">
                                <div class="collbl">
                                    Venc. Original:
                                </div>
                                <div class="coltxt" style="width: 518px;">
                                    <asp:Label ID="LblVencOriginal" runat="server" Text="Label" Width="125px" />&nbsp;
                                </div>
                            </div>
                            <div id="divDataBaixa" runat="server">
                                <div class="collbl">
                                    <asp:Label ID="lblDataBaixa" runat="server" Text="Data Baixa" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataBaixa" ClientIDMode="Static" runat="server" Width="104px"
                                        OnTextChanged="txtDataBaixa_TextChanged" AutoPostBack="True" CssClass="calendario" />
                                </div>
                            </div>
                        </div>
                        <div id="divPedidoRecPag" class="row" runat="server">
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPedidoRecPag" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnPedidoRecPag" runat="server" OnClick="btnPedidoRecPag_Click" Text=" > "
                                    CssClass="btn" UseSubmitBehavior="False" />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgLimparPedidoRecPag" runat="server" Height="22px" ImageAlign="AbsMiddle"
                                    ImageUrl="~/images/borracha.JPG" OnClick="imgLimparPedidoRecPag_Click" Width="22px"
                                    CssClass="btn" />
                            </div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgExtratoPedidoRecPag" runat="server" Height="23px" ImageAlign="AbsMiddle"
                                    ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imgExtratoPedidoRecPag_Click"
                                    data-ToolTip="default" ToolTip="Visualizar Extrato do Pedido" Width="23px" />
                            </div>
                        </div>
                        <div class="row" id="divCodigoDeBarras" runat="server">
                            <div class="collbl">
                                Código de Barras:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigoDeBarras" runat="server" Width="350px" />
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="CkbCodigoDeBarras" runat="server" Text="Digitado" />
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="ckPreImpresso" runat="server" Text="Pré Impresso" />
                            </div>
                        </div>
                        <div id="divBaixaAdiantamento" runat="server" class="row">
                            <div class="subtitulodiv">
                                <asp:Label ID="lblAdiantamento" runat="server" Text="Adiantamentos Abertos" />
                            </div>
                            <div class="bordagrid" style="height: 200px;">
                                <asp:GridView ID="gridAdiantamentosDisponiveis" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Visible="False">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:TemplateField HeaderText="Pedido">
                                            <ItemTemplate>
                                                <%# IIf(Eval("Titulo.CodigoPedido") = 0, "Avulso", Eval("Titulo.CodigoPedido")) %>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Conta Contábil">
                                            <ItemTemplate>
                                                <%# Eval("Titulo.ContaContabilCliFor.Titulo")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="CodigoTitulo" HeaderText="Título" />
                                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento" />
                                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento" />
                                        <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor Oficial" />
                                        <asp:BoundField DataField="TotalOficialBaixado" DataFormatString="{0:N2}" HeaderText="Valor Baixado" />
                                        <asp:BoundField DataField="SaldoValorOficial" DataFormatString="{0:N2}" HeaderText="Saldo" />
                                        <asp:BoundField DataField="Taxa" DataFormatString="{0:N2}" HeaderText="Taxa" />
                                        <asp:TemplateField HeaderText="Valor Baixa">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtVlrBaixa" runat="server" DataFormatString="{0:N2}" CssClass="txtDecimal"
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
                                    CellPadding="4" ForeColor="#333333" GridLines="None">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:BoundField DataField="CodigoTituloAdiantamento" DataFormatString="{0:N0}" HeaderText="Adiant." />
                                        <asp:BoundField DataField="Sequencia" DataFormatString="{0:N0}" HeaderText="Seq." />
                                        <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento" />
                                        <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor Oficial" />
                                        <asp:BoundField DataField="ValorMoeda" DataFormatString="{0:N2}" HeaderText="Valor Moeda" />
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
                        <div id="divLancamentoContabil" runat="server">
                            <div class="row">
                                <div class="subtitulodiv">
                                    <asp:Label ID="lblLancContabeis" runat="server" Text="Lançamentos Contábeis" />
                                </div>
                            </div>
                            <div class="row" id="divProduto" runat="server">
                                <div class="collbl">
                                    Produto:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtProduto" runat="server" Enabled="False" Width="337px" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnProduto" runat="server" OnClick="btnProduto_Click" Text=" > "
                                        CssClass="btn" UseSubmitBehavior="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnProdutoExcluir" runat="server" OnClick="btnProdutoExcluir_Click"
                                        CssClass="btn" Text=" X " UseSubmitBehavior="False" />
                                </div>
                            </div>
                            <div class="row" id="divQuantidade" runat="server">
                                <div class="collbl">
                                    Quantidade:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtQuantidade" runat="server" AutoPostBack="True" OnTextChanged="txtQuantidade_TextChanged" />
                                </div>
                            </div>
                        </div>
                        <div class="row" id="divBordero" runat="server" visible="False">
                            <div class="collbl">
                                Data de envio:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataEnvioBordero" runat="server" CssClass="calendario" />
                            </div>
                            <div class="collbl">
                                Juros(%):
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtJurosBordero" runat="server" CssClass="txtDecimal" />
                            </div>
                            <div class="collbl">
                                D+ (Float):
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblDiasBanco" runat="server" Text="0 Dia(s)" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Histórico:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtHistorico" TabIndex="10" runat="server" Width="351px" TextMode="MultiLine"
                                    MaxLength="200" />
                            </div>
                            <div class="collbl">
                                Observação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacoes" runat="server" Width="392px" TextMode="MultiLine" />
                            </div>
                        </div>
                        <div class="painelleft" style="width: 70%; margin-right: 10px;">
                            <div class="bordagrid" style="height: 150px;">
                                <asp:GridView ID="gridValores" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" OnRowDataBound="gridValores_RowDataBound">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:BoundField DataField="CodigoContaEncargo" HeaderText="Conta" />
                                        <asp:BoundField DataField="Descricao" HeaderText="Titulo" />
                                        <asp:TemplateField HeaderText="Valor Oficial">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtValorOficial" Style="text-align: right;" runat="server" AutoPostBack="True"
                                                    OnTextChanged="txtValorOficial_TextChanged" Text='<%# Eval("ValorOficial", "{0:N2}") %>'
                                                    Width="100px" CssClass="txtDecimal" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Valor Moeda">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtValorMoeda" Style="text-align: right;" runat="server" AutoPostBack="True"
                                                    OnTextChanged="txtValorMoeda_TextChanged" Text='<%# Eval("ValorMoeda", "{0:N2}") %>'
                                                    Width="100px" CssClass="txtDecimal" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Deb./Cred.">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="ddlDebitoCredito" runat="server" Enabled="False" SelectedValue='<%# Bind("DC") %>'
                                                    Width="126px">
                                                    <asp:ListItem Selected="True" Value="D">Debito</asp:ListItem>
                                                    <asp:ListItem Value="C">Credito</asp:ListItem>
                                                    <asp:ListItem Value="I">INFORME</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Centro De Custo">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="ddlCentroDeCusto" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlCentroDeCusto_SelectedIndexChanged"
                                                    Width="369px">
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
                                <asp:HiddenField ID="hdnValorTitulo" runat="server" />
                            </div>
                        </div>
                        <div class="painelleft" style="width: 28%;">
                            <div class="row">
                                <div class="coltxt">
                                    <asp:Button ID="BtnAdicionarConta" runat="server" Text=" + " UseSubmitBehavior="False"
                                        data-ToolTip="default" ToolTip="Adicionar Conta" CssClass="btn" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt">
                                    <asp:Label ID="lblTotalDeCreditos" runat="server" Font-Bold="True" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt">
                                    <asp:Label ID="lblTotalDeDebitos" runat="server" Font-Bold="True" />
                                </div>
                            </div>
                        </div>
                        <div class="row" id="divtblBordero" runat="server" visible="False">
                            <div class="subtitulodiv">
                                <asp:Label ID="lblContaBordero" runat="server" />
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Cálculo:
                                </div>
                                <div class="coltxt">
                                    <asp:CheckBox ID="chkJuros" runat="server" Text="Aplicar juros" AutoPostBack="True" />
                                    <asp:CheckBox ID="chkTaxa" runat="server" Text="Calcular Taxa" AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row" runat="server" visible="False">
                                <div class="collbl">
                                    Juros:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtJuros" runat="server" CssClass="txtDecimal" />
                                </div>
                            </div>
                            <div class="row" runat="server" visible="False">
                                <div class="collbl">
                                    Valor:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtTaxa" runat="server" CssClass="txtDecimal" />
                                    <asp:CheckBox ID="chkProporcional" runat="server" Text="Proporcional" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt">
                                    <asp:Button ID="btnCalcular" runat="server" Text="Calcular" CssClass="botao" />
                                </div>
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 150px; display: none;">
                            <asp:GridView ID="gdvDesmembramentoValores" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:BoundField DataField="Codigo" HeaderText="Registro" />
                                    <asp:BoundField DataField="CliFor.Nome" HeaderText="Nome" />
                                    <asp:BoundField DataField="Valores(0).ValorOficial" HeaderText="Valor" DataFormatString="{0:N2}" />
                                    <asp:BoundField DataField="Vencimento" HeaderText="Vencimento" DataFormatString="{0:dd/MM/yyyy}" />
                                    <asp:TemplateField HeaderText="Valor">
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtValor" runat="server" Width="100px" CssClass="txtDecimal" />
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
                        <asp:HiddenField ID="txtUsuarioLiberarPedido" runat="server"></asp:HiddenField>
                        <asp:HiddenField ID="txtUsuarioLiberarPedidoData" runat="server"></asp:HiddenField>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="tabParcelas" runat="server">
                    <HeaderTemplate>
                        Parcelas / Contabilização
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="subtitulodiv">
                            Todas as Parcelas
                        </div>
                        <div class="bordagrid" style="height: 200px;">
                            <asp:GridView ID="GridParcelas" runat="server" Width="100%" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="GridParcelas_SelectedIndexChanged">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:BoundField DataField="Registro" HeaderText="Registro">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Situacao" HeaderText="Situacao" />
                                    <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Baixa" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data da Baixa"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Historico" HeaderText="Hist&#243;rico">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Moeda" HeaderText="Moeda" />
                                    <asp:BoundField DataField="ValorDoDocumento" DataFormatString="{0:N}" HeaderText="Valor da Parcela"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Deducoes" DataFormatString="{0:N}" HeaderText="Dedu&#231;&#245;es"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Acrescimos" DataFormatString="{0:N}" HeaderText="Acr&#233;scimos"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorLiquido" DataFormatString="{0:N}" HeaderText="Valor da Baixa"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Saldo" DataFormatString="{0:N}" HeaderText="Saldo" HtmlEncode="False">
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
                        <div class="subtitulodiv">
                            <asp:Label ID="lblContabilizacao" runat="server" Font-Bold="False" Font-Italic="False" />
                        </div>
                        <div class="bordagrid" style="height: 200px;">
                            <asp:GridView ID="gridRazao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <Columns>
                                    <asp:BoundField HeaderText="Empresa" DataField="Reduzido">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Conta" DataField="Conta">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Cliente" DataField="Cliente">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Titulo" HeaderText="Descri&#231;&#227;o">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento" DataField="Movimento">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Lote" DataField="Lote">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Produto" DataField="Produto">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Custo" DataField="Custo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Hist&#243;rico" DataField="Historico">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="D&#233;bito" DataField="Debito" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Cr&#233;dito" DataField="Credito" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField HeaderText="Saldo" DataField="Saldo" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
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
                        <div class="subtitulodiv">
                            <asp:Label ID="lblNota" runat="server" Text=" Nota Fiscal Relacionada ao Título" />
                        </div>
                        <div class="bordagrid" style="height: 200px;">
                            <asp:GridView ID="gridNota" runat="server" CellPadding="4" ForeColor="#333333" GridLines="None">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" Width="100%" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="tabConsulta" runat="server">
                    <HeaderTemplate>
                        Consulta Títulos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div id="divParametros" runat="server">
                            <div class="menu_acoes">
                                <div class="acoes">
                                    <ul>
                                        <li runat="server">
                                            <asp:LinkButton class="iconConsultar" ID="lnkConsultarAbaConsulta" Text="Consultar"
                                                runat="server" />
                                        </li>
                                        <li runat="server">
                                            <asp:LinkButton class="iconLimpar" ID="lnkLimparAbaConsulta" runat="server" Text="Limpar" />
                                        </li>
                                        <li runat="server" visible="false">
                                            <asp:LinkButton class="iconNovo" ID="lnkAcao" runat="server" Text="Agrupar" />
                                        </li>
                                        <li runat="server">
                                            <asp:LinkButton class="iconNovo" ID="lnkSlip" runat="server" Text="Slip" />
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div class="painelright" style="width: 28%;">
                                <div class="subtitulodiv">
                                    Ação:
                                </div>
                                <div class="row">
                                    <div class="coltxt" style="line-height: 20px;">
                                        <asp:RadioButton ID="rdConsultar" runat="server" AutoPostBack="True" Checked="True"
                                            GroupName="rdAcao" Text="Consultar" />
                                        <br />
                                        <asp:RadioButton ID="rdConsultarBordero" runat="server" Text="Consultar Borderos"
                                            GroupName="rdAcao" />
                                        <br />
                                        <asp:RadioButton ID="rdAgrupar" runat="server" AutoPostBack="True" GroupName="rdAcao"
                                            Text="Agrupar" />
                                        <br />
                                        <asp:RadioButton ID="rdBaixar" runat="server" AutoPostBack="True" GroupName="rdAcao"
                                            Text="Baixar" />
                                        <br />
                                        <asp:RadioButton ID="rdReprogramar" runat="server" AutoPostBack="True" GroupName="rdAcao"
                                            Text="Reprogramar" />
                                        <br />
                                        <asp:RadioButton ID="rdAdTituloAgrupamento" runat="server" AutoPostBack="True" GroupName="rdAcao"
                                            Text="Adicionar/Remover Titulo do Agrupamento" />
                                        <br />
                                        <asp:RadioButton ID="rdDuplicatasDescontadas" runat="server" AutoPostBack="True"
                                            GroupName="rdAcao" Text="Borderô (Gera Titulo de Recebimento)" />
                                        <br />
                                        <asp:RadioButton ID="rdGeraDuplicatasCobrancaSimples" runat="server" AutoPostBack="True"
                                            GroupName="rdAcao" Text="Gerar Duplicatas (Cobrança Simples)" />
                                        <br />
                                        <asp:RadioButton ID="rbCheque" runat="server" AutoPostBack="True" GroupName="rdAcao"
                                            Text="Emissão de Cheque" Visible="false" />
                                    </div>
                                </div>
                            </div>
                            <div class="painelright" style="width: 72%">
                                <div class="row" runat="server" visible="False" style="line-height: 16px;">
                                    <div class="collbl">
                                        Ação:
                                    </div>
                                    <div class="coltxt">
                                        <asp:CheckBoxList ID="chkAcao" runat="server" RepeatDirection="Horizontal" RepeatColumns="4">
                                            <asp:ListItem>Consultar</asp:ListItem>
                                            <asp:ListItem>Consultar Borderô</asp:ListItem>
                                            <asp:ListItem>Agrupar</asp:ListItem>
                                            <asp:ListItem>Reprogramar</asp:ListItem>
                                            <asp:ListItem>Adicionar/Remover Titulo do Agrupamento</asp:ListItem>
                                            <asp:ListItem>Gerar Duplicatas Descontadas (Gera Titulo de Recebimento)</asp:ListItem>
                                            <asp:ListItem>Gerar Duplicatas (Cobrança Simples)</asp:ListItem>
                                            <asp:ListItem>Emissão de Cheque</asp:ListItem>
                                        </asp:CheckBoxList>
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Unidade:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlUnidadeConsultaTitulos" runat="server" AutoPostBack="True"
                                            OnSelectedIndexChanged="DdlUnidadeConsultaTitulos_SelectedIndexChanged" Width="596px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Empresa:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlEmpresaConsultaTitulos" runat="server" Width="594px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Borderô:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtBorderoConsulta" runat="server" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        <asp:CheckBox ID="chkConsolidarClienteConsulta" runat="server" Text="Cliente:" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:HiddenField ID="txtCodigoClienteConsulta" runat="server" />
                                        <asp:TextBox ID="txtClienteConsulta" runat="server" Width="552px" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:Button ID="cmdConsultaClientes" runat="server" OnClick="cmdConsultaClientes_Click"
                                            CssClass="btn" Text=" &gt; " UseSubmitBehavior="False" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Período de:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" runat="server" CausesValidation="True"
                                            CssClass="calendario" data-ToolTip="default" ToolTip="Periodo Inicial" Width="100px"
                                            ClientIDMode="Static" />
                                        &nbsp;à&nbsp;<asp:TextBox ID="txtPeriodoFinalConsultaTitulos" runat="server" CausesValidation="True"
                                            CssClass="calendario" data-ToolTip="default" ToolTip="Periodo Final" Width="100px"
                                            ClientIDMode="Static" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Situação:
                                    </div>
                                    <div class="coltxt">
                                        <asp:CheckBox ID="chkBaixa" runat="server" Text="Baixados" />
                                        <asp:CheckBox ID="chkProvisao" runat="server" Text="Provisão" />
                                        <asp:CheckBox ID="chkPrevisao" runat="server" Checked="True" Text="Previsão" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Tipo Título:
                                    </div>
                                    <div class="coltxt">
                                        <asp:CheckBox ID="chkPagar" runat="server" Checked="True" Text="Pagar" />
                                        <asp:CheckBox ID="chkReceber" runat="server" Checked="True" Text="Receber" />
                                        <asp:CheckBox ID="chkContabil" runat="server" Checked="True" Text="Contabil" />
                                    </div>
                                </div>
                                <div class="row" id="divParametrosAdicionarTituloAgrupado" runat="server">
                                    <div class="collbl">
                                        Título Mestre:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtTituloConsulta" runat="server" />
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div id="divAcoes" runat="server">
                            <div class="row">
                                <div class="subtitulodiv">
                                    <asp:Label ID="lblAcao" runat="server" Text="Consulta de Título(s)" />
                                </div>
                            </div>
                            <div class="row" id="divAcaoAgrupar" runat="server" />
                            <div id="divAcaoBaixar" runat="server">
                                <div class="row">
                                    <div class="collbl">
                                        Unidade:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlUnidadeDeNegocioBaixa" runat="server" AutoPostBack="True"
                                            Width="594px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Empresa:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlEmpresaBaixaTodos" runat="server" AutoPostBack="True" Width="594px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Tipo:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="DdlTiposDeRecebimentosBaixarTodos" runat="server" TabIndex="7"
                                            Width="370px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Banco:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlContaContabilEmpresaRecPagBaixarTodos" runat="server" Width="594px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Data:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtDataBaixarTodos" runat="server" data-ToolTip="default" ToolTip="Periodo Inicial"
                                            Width="100px" CssClass="calendario" />
                                    </div>
                                </div>
                            </div>
                            <div id="divAcaoReprogramacao" runat="server">
                                <div class="row">
                                    <div class="collbl">
                                        Reprogramação:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtDataReprogramacao" runat="server" data-ToolTip="default" ToolTip="Data de Reprogramação"
                                            Width="100px" CssClass="calendario" />
                                    </div>
                                </div>
                            </div>
                            <div id="divAcaoAdicionar" runat="server" />
                            <div id="divAcaoDuplicatas" runat="server">
                                <div class="row">
                                    <div class="coltxt">
                                        <asp:CheckBox ID="chkRecompraPorBordero" runat="server" AutoPostBack="True" Text="Havera Recompra de Titulos" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:ImageButton ID="imgRecompraPorBordero" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" />
                                    </div>
                                </div>
                            </div>
                            <div id="divRecompraPorBordero" runat="server">
                                <div class="subtitulodiv">
                                    Titulos Recompra
                                </div>
                                <div class="bordagrid">
                                    <asp:GridView ID="gridRecompra" runat="server" AllowSorting="True" AutoGenerateColumns="False"
                                        CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                        <Columns>
                                            <asp:TemplateField HeaderText="CK">
                                                <HeaderTemplate>
                                                    <asp:CheckBox ID="chkTodosRecompra" runat="server" AutoPostBack="True" OnCheckedChanged="chkTodosRecompra_CheckedChanged" />
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:UpdatePanel ID="updGridConsultaTitulos0" runat="server" ChildrenAsTriggers="false"
                                                        UpdateMode="Conditional">
                                                        <ContentTemplate>
                                                            <asp:CheckBox ID="ChkGridRecompra" runat="server" AutoPostBack="True" OnCheckedChanged="ChkGridRecompra_CheckedChanged" />
                                                        </ContentTemplate>
                                                        <Triggers>
                                                            <asp:AsyncPostBackTrigger ControlID="ChkGridRecompra" EventName="CheckedChanged" />
                                                        </Triggers>
                                                    </asp:UpdatePanel>
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle Width="30px" />
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="Doc" HeaderText="Doc." />
                                            <asp:BoundField DataField="Titulo_id" HeaderText="Titulo">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" Width="60px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Nome" HeaderText="Cliente" />
                                            <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencto"
                                                HtmlEncode="False">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" Width="50px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Baixa" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Baixa" />
                                            <asp:BoundField DataField="Historico" HeaderText="Histórico" HtmlEncode="False">
                                                <HeaderStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Moeda" HeaderText="Moeda" HtmlEncode="False" />
                                            <asp:BoundField DataField="ValorDoDocumento" DataFormatString="{0:N2}" HeaderText="Vlr Documento" />
                                            <asp:BoundField DataField="Deducoes" DataFormatString="{0:N2}" HeaderText="Deduções"
                                                HtmlEncode="False" />
                                            <asp:BoundField DataField="Acrescimos" DataFormatString="{0:N2}" HeaderText="Acrescimos">
                                                <HeaderStyle HorizontalAlign="Center" />
                                                <ItemStyle HorizontalAlign="Center" Width="30px" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ValorLiquido" DataFormatString="{0:N2}" HeaderText="Vlr Liquido" />
                                            <asp:BoundField DataField="Situacao" HeaderText="Situacao" />
                                            <asp:TemplateField HeaderText="Cheque">
                                                <ItemTemplate>
                                                    <%# IIf((Boolean.Parse(Eval("Cheque").ToString())),"Sim","Não") %>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="Bordero" HeaderText="Bordero" />
                                        </Columns>
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
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
                            <div id="divAcaoDuplicatasCobrancaSimples" runat="server">
                                <div class="row">
                                    <div class="collbl">
                                        Banco:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlBancoDuplicataCobrancaSimples" runat="server" Width="594px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Carteira:
                                    </div>
                                    <div class="coltxt">
                                        <asp:DropDownList ID="ddlCarteiraDuplicataCobrancaSimples" runat="server" Width="594px" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="collbl">
                                        Data Envio:
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtDataEnvioBorderoCobrancaDuplicataSimples" runat="server" CssClass="calendario" />
                                    </div>
                                </div>
                                <div class="row">
                                    <div class="coltxt">
                                        <asp:CheckBox ID="chkImpBorderoCobrancaSimples" runat="server" Text="Imprimir Bordero" />
                                    </div>
                                    <div class="coltxt">
                                        <asp:CheckBox ID="chkImprimirDuplicatasCobrancaSimples" runat="server" Text="Imprimir Duplicatas" />
                                    </div>
                                </div>
                            </div>
                            <div id="divAcaoConsultaBordero" runat="server">
                                <div class="bordagrid" style="height: 200px;">
                                    <asp:GridView ID="gridBordero" Width="100%" runat="server" CellPadding="4" ForeColor="#333333"
                                        GridLines="None" AutoGenerateColumns="False">
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:TemplateField>
                                                <ItemTemplate>                                                    
                                                    <asp:LinkButton ID="lnkSelecionar" CssClass="lnk" Visible="True"
                                                        data-tooltip="default" ToolTip="Selecionar registro." runat="server"
                                                        CommandName="Select">
                                                            <i class="fa fa-arrow-right seta"></i>
                                                    </asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="TituloBordero" HeaderText="Título" />
                                            <asp:BoundField DataField="DataEnvio" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data" />
                                            <asp:TemplateField HeaderText="Borderô">
                                                <ItemTemplate>
                                                    <%# Eval("Contrato") & "-" & Eval("ContaBordero")%>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="Nome" HeaderText="Empresa" />
                                            <asp:BoundField DataField="Empresa" HeaderText="CNPJ" />
                                            <asp:TemplateField HeaderText="Cidade">
                                                <ItemTemplate>
                                                    <%# Eval("Cidade") & "/" & Eval("Estado") %>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:BoundField DataField="Historico" HeaderText="Histórico" />
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
                            <div class="row" runat="server" visible="False">
                                <div class="coltxt">
                                    <asp:Label ID="lblTotalRegistroAgrupado" runat="server" Font-Bold="True" Text="Título agrupados no valor total:" />
                                </div>
                            </div>
                            <div class="bordagrid">
                                <asp:GridView ID="gridTitulosAcao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None">
                                    <Columns>
                                        <asp:BoundField DataField="Codigo" HeaderText="Titulo" />
                                    </Columns>
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
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
                                <asp:GridView ID="GridConsultaTitulos" runat="server" AllowSorting="True" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnRowDataBound="GridConsultaTitulos_RowDataBound"
                                    OnSelectedIndexChanged="GridConsultaTitulos_SelectedIndexChanged">
                                    <Columns>
                                        <asp:BoundField DataField="Titulo_id" HeaderText="Titulo">
                                            <ControlStyle CssClass="none" />
                                            <HeaderStyle CssClass="none" />
                                            <ItemStyle CssClass="none" />
                                        </asp:BoundField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="" CssClass="lnk" Visible="True"
                                                    data-tooltip="default" TlnkSelecionaroolTip="Selecionar registro." runat="server"
                                                    CommandName="Select">
                                                        <i class="fa fa-arrow-right seta"></i>
                                                </asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lnkSelecionarAgrupado" CssClass="lnk" Visible="True"
                                                    data-tooltip="default" ToolTip="Selecionar registro." runat="server"
                                                    CommandName="Select">
                                                        <i class="fa fa-arrow-right seta"></i>
                                                </asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <HeaderTemplate>
                                                <asp:CheckBox ID="chkSelectAll" runat="server" data-ToolTip="default" ToolTip="Selecionar Todos"
                                                    OnCheckedChanged="chkSelectAll_CheckedChanged" AutoPostBack="true" />
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:UpdatePanel ID="updGridConsultaTitulos" runat="server" ChildrenAsTriggers="false"
                                                    UpdateMode="Conditional">
                                                    <ContentTemplate>
                                                        <asp:CheckBox ID="ChkGridTitulos" runat="server" AutoPostBack="True" OnCheckedChanged="ChkGridTitulos_CheckedChanged" />
                                                    </ContentTemplate>
                                                    <Triggers>
                                                        <asp:AsyncPostBackTrigger ControlID="ChkGridTitulos" EventName="CheckedChanged" />
                                                    </Triggers>
                                                </asp:UpdatePanel>
                                            </ItemTemplate>
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle Width="30px" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Documento">
                                            <ItemTemplate>
                                                Título&nbsp;<%# Eval("Titulo_id")%><br />
                                                Documento:&nbsp;<%# Eval("Doc")%><br />
                                                Situação&nbsp;<%# Eval("Situacao")%><br />
                                                <b>
                                                    <%# Eval("Agrupamento")%></b>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Doc" HeaderText="Doc.">
                                            <ControlStyle CssClass="none" />
                                            <HeaderStyle CssClass="none" />
                                            <ItemStyle CssClass="none" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="TituloRecompra" HeaderText="Recompra">
                                            <ControlStyle CssClass="none" />
                                            <HeaderStyle CssClass="none" />
                                            <ItemStyle CssClass="none" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Agrupamento" HeaderText="Agrup.">
                                            <ControlStyle CssClass="none" />
                                            <HeaderStyle CssClass="none" />
                                            <ItemStyle CssClass="none" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Cliente">
                                            <ItemTemplate>
                                                Cliente:&nbsp;<%# Eval("Nome")%><br />
                                                Histórico:&nbsp;<%# Eval("Historico")%>
                                            </ItemTemplate>
                                            <ItemStyle Width="300px" />
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Nome" HeaderText="Cliente" Visible="False" />
                                        <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" Width="50px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Baixa" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Baixa" />
                                        <asp:BoundField DataField="Historico" HeaderText="Histórico" HtmlEncode="False" Visible="False">
                                            <HeaderStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Valor">
                                            <ItemTemplate>
                                                Moeda:&nbsp;<%# Eval("Moeda")%><br />
                                                Valor...........:&nbsp;<%# Eval("ValorDoDocumento", "{0:N2}")%><br />
                                                Deduções.....:&nbsp;<%# Eval("Deducoes", "{0:N2}")%><br />
                                                Acréscimos....:&nbsp;<%# Eval("Acrescimos", "{0:N2}")%><br />
                                                Liquido.........:&nbsp;<%# Eval("ValorLiquido", "{0:N2}")%>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Moeda" HeaderText="Moeda" HtmlEncode="False" Visible="False" />
                                        <asp:BoundField DataField="ValorDoDocumento" DataFormatString="{0:N2}" HeaderText="Vlr Documento"
                                            Visible="False" />
                                        <asp:BoundField DataField="Deducoes" DataFormatString="{0:N2}" HeaderText="Deduções"
                                            HtmlEncode="False" Visible="False" />
                                        <asp:BoundField DataField="Acrescimos" DataFormatString="{0:N2}" HeaderText="Acrescimos"
                                            Visible="False">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" Width="30px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ValorLiquido" DataFormatString="{0:N2}" HeaderText="Vlr Liquido"
                                            Visible="False" />
                                        <asp:BoundField DataField="Situacao" HeaderText="Situacao" Visible="False" />
                                        <asp:TemplateField HeaderText="Cheque">
                                            <ItemTemplate>
                                                <%# IIf((Boolean.Parse(Eval("Cheque").ToString())),"Sim","Não") %>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Acoes">
                                            <ItemTemplate>
                                                <asp:Label ID="lblEmiteDuplicatasGrid" runat="server" Text='<%# Bind("EmiteDuplicata") %>' />
                                                <asp:Button ID="btnDuplicataGrid" runat="server" Text="Duplicata" OnClick="btnDuplicataGrid_Click" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                    </Columns>
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
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
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
    <uc:ConsultaDadosBancarios ID="ucConsultaDadosBancarios" runat="server" />
    <uc:ConsultaEncargosPlanoDeContas ID="ucConsultaEncargosPlanoDeContas" runat="server" />
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
</asp:Content>
