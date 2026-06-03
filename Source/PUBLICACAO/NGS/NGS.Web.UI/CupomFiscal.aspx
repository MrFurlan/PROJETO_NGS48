<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="CupomFiscal.aspx.vb" Inherits="NGS.Web.UI.CupomFiscal" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1360px !important;
        }
    </style>
    <script type="text/javascript">
        function pageLoadCupomFiscal() {
            $('#MainContent_txtCPF').attr('placeholder', 'CPF');
            $('#MainContent_txtNomeCliente').attr('placeholder', 'Nome do Cliente');
            $('#MainContent_txtCelular').attr('placeholder', 'Telefone');
            $('#MainContent_txtNovoCliente').attr('placeholder', 'Nome do Cliente');

        }

        $(document).ready(function () {
            pageLoadCupomFiscal();
            var prmCupomFiscal = Sys.WebForms.PageRequestManager.getInstance();
            prmCupomFiscal.add_endRequest(pageLoadCupomFiscal);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngNotasFiscaisGerais" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlNotaFiscalXItens" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Cupom Fiscal Eletrônico
                <div class="row" style="float: right; margin-top: 0; padding-top: 2px;">
                    <div class="coltxt">
                        <asp:Image ID="imgUsuario" runat="server" Height="20px" ImageAlign="AbsMiddle" ImageUrl="~/images/man2.png"
                            Width="18px" />
                    </div>
                    <div class="coltxt">
                        <asp:Label ID="lblUsuario" Visible="false" runat="server" Font-Bold="True" Font-Size="11px" />
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlUsuarios" runat="server" Width="175px" />
                    </div>
                </div>
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconAtualizar" runat="server">
                            <asp:LinkButton ID="lnkReenviar" Text="Reenviar" runat="server" />
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkCancelar" Text="Cancelar" runat="server" OnClientClick="if(!confirm('Deseja realmente cancelar este cupom fiscal?')) return false;" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" AccessKey="C" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:UpdatePanel ID="updpnlVisualizar" runat="server">
                                <ContentTemplate>
                                    <asp:LinkButton ID="lnkVisualizar" Text="Pré-Visualizar" runat="server" />
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="lnkVisualizar" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                        <li runat="server" id="btnContabil" visible="false">
                            <asp:Button ID="BtnRecontabilizar" runat="server" CssClass="botao" OnClick="BtnRecontabilizar_Click"
                                Style="margin: 0px;" Text="Recontabilizar" UseSubmitBehavior="False" />
                        </li>
                        <li runat="server">
                            <div class="row" style="margin-top: 0;">
                                <div class="coltxt">
                                    <asp:Button ID="btnModo" CssClass="btn" runat="server" BackColor="Green" BorderStyle="None"
                                        Enabled="False" Font-Bold="True" ForeColor="White" Style="cursor: pointer; font-family: Tahoma,Arial,Helvetica,sans-serif; font-size: 11px; height: 24px; text-align: center; width: 150px;"
                                        Text="MODO NORMAL" UseSubmitBehavior="False" />
                                </div>
                            </div>
                        </li>
                        <li runat="server" style="float: right;">
                            <div class="row" style="margin-top: 0;">
                                <div class="coltxt">
                                    <asp:TextBox ID="txtAvisoCupom" runat="server" Font-Size="17px" BorderStyle="None" Font-Bold="true" ForeColor="Red" data-ToolTip="default" />
                                </div>
                            </div>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="painelleft">
                <div class="row">
                    <div class="collbl">
                        Unidade:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" AutoPostBack="True" TabIndex="1"
                            Width="600px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Empresa:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlEmpresa" runat="server" AutoPostBack="True" Width="600px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Cliente:
                    </div>
                    <div class="coltxt">
                        <asp:HiddenField ID="hdfCodigoCliente" runat="server" />
                        <asp:TextBox ID="txtCPF" runat="server" MaxLength="11" Width="90px"
                            data-ToolTip="default" ToolTip="CPF" />
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtNovoCliente" runat="server" MaxLength="80" Width="342px"
                            data-ToolTip="default" ToolTip="Nome do Cliente" />
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtCelular" CssClass="txtFone" runat="server" MaxLength="75" Width="100px"
                            data-ToolTip="default" ToolTip="Num Telefone" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnConsultaClientes" runat="server" TabIndex="4" Text="&gt;" CssClass="btn"
                            UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                    </div>
                </div>
            </div>
            <div class="painelleft">
                <div class="row">
                    <div class="collbl">
                        Situação:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlSituacao" runat="server" Width="470px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Data Nota:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataNota" runat="server" CssClass="calendario" Width="88px" data-ToolTip="default"
                            ToolTip="Data de Criação da Nota." />
                    </div>
                </div>
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%" Style="clear: both;">
                <ajaxToolkit:TabPanel runat="server" ID="TabProdutos">
                    <HeaderTemplate>
                        Produtos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="painelleft" style="width: 77%;">
                            <div class="row" style="margin-top: 0;">
                                <div class="menu_acoes" runat="server" style="margin-top: 0;">
                                    <div class="acoes">
                                        <ul>
                                            <li class="iconMais" runat="server">
                                                <asp:LinkButton ID="lnkAdicionarItem" Text="Adicionar" runat="server" />
                                            </li>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                            <div class="bordagrid" style="height: 300px;">
                                <asp:GridView ID="grdProdutos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <Columns>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgExcluir" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                    OnClick="imgExcluir_Click" data-ToolTip="default" Visible="true" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente excluir o produto?');" />
                                            </ItemTemplate>
                                            <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                            <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" HtmlEncode="False"
                                            ReadOnly="True">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" Width="60px" />
                                        </asp:BoundField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:Label ID="lblCodigoProdutoGrid" runat="server" Text='<%# Bind("Produto") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Quantidade">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                            <ItemTemplate>
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtQuantidadeFiscalGrid" CssClass="txtNumerico9" DataFormatString="{0:N0}" AutoPostBack="true" OnTextChanged="txt_TextChanged" runat="server" Text='<%# Bind("QuantidadeFiscal") %>' />
                                                </div>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Unitário">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                            <ItemTemplate>
                                                <div class="coltxt">
                                                    <asp:TextBox ID="txtUnitarioGrid" CssClass="txtDecimal2" DataFormatString="{0:N2}" AutoPostBack="true" OnTextChanged="txt_TextChanged" runat="server" Text='<%# Bind("Unitario") %>' />
                                                </div>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="ValorTotal" DataFormatString="{0:N2}" HeaderText="Total"
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Lote(s)">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgSelecionar" OnClick="imgSelecionar_Click" runat="server"
                                                    Width="18px" Height="18px" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" data-ToolTip="default"
                                                    ToolTip="Visualizar Lote(s)" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Limpar">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" Width="50px" />
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgLimpar" OnClick="imgLimpar_Click" runat="server"
                                                    Width="20px" Height="20px" ImageUrl="~/images/Borracha.jpg" data-ToolTip="default"
                                                    ToolTip="Limpar unitário e valor" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                    </Columns>
                                    <EditRowStyle BackColor="#999999" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                </asp:GridView>
                            </div>
                        </div>
                        <div class="painelright" style="width: 22%;">
                            <div class="titulodiv" style="font-size: 12px;">
                                Encargos Cupom Fiscal
                            </div>
                            <div class="bordagrid" style="height: 300px;">
                                <asp:GridView ID="gridEncargosGerais" runat="server" AutoGenerateColumns="False"
                                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:BoundField DataField="Codigo" HeaderText="Encargo">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor">
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
                                    <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                    <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                    <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                    <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                </asp:GridView>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabVencimentosold">
                    <HeaderTemplate>
                        Vencimentos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconMais" runat="server">
                                        <asp:LinkButton ID="LnkParcelamento" Text="Parcelar" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparParcelamento" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconNovo" runat="server" visible="False">
                                        <asp:CheckBox ID="chkProvisao" runat="server" Text="Gerar Titulos em Provisão" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Carteira:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbCarteira" runat="server" Font-Names="monospace" Width="395px"
                                    AutoPostBack="True" />
                            </div>
                            <div class="collbl">
                                Total Nota:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTotalNota" runat="server" CssClass="txtDecimal" BackColor="#FFFFC0" BorderColor="White"
                                    Font-Bold="True" ForeColor="Red">0</asp:TextBox>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Condicões:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="hdCondicaoDePagamento" runat="server" />
                                <asp:DropDownList ID="ddlCondicaoDePagamento" runat="server" AutoPostBack="True"
                                    Font-Names="monospace" Width="395px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Forma:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbFormas" runat="server" Font-Names="monospace" Width="395px" />
                            </div>
                        </div>
                        <div class="row" id="divBanco" runat="server" visible="False">
                            <div class="collbl">
                                Banco
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnDadosBancarios" runat="server" OnClick="btnDadosBancarios_Click"
                                    Text=" &gt; " UseSubmitBehavior="False" />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblBanco" runat="server" Font-Bold="True" Text="Banco" />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblAgencia" runat="server" Font-Bold="True" Text="Agência" />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblContaCorrente" runat="server" Font-Bold="True" Text="Conta" />
                            </div>
                        </div>
                        <div class="painelleft" style="width: 57%;">
                            <div class="bordagrid" style="height: 145px;">
                                <asp:GridView ID="grdCondicoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:BoundField DataField="codigo" HeaderText="Título" />
                                        <asp:TemplateField HeaderText="Situação">
                                            <ItemTemplate>
                                                <asp:Label ID="lblDescricaoProvisao" runat="server" Text='<%# Bind("DescricaoProvisao") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="Prorrogacao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ValorDoDocumento" DataFormatString="{0:N2}" HeaderText="Valor Doc."
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Descontos" HeaderText="Descontos" DataFormatString="{0:N2}"
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Deducoes" HeaderText="Deduções" DataFormatString="{0:N2}"
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Juros" HeaderText="Juros" DataFormatString="{0:N2}" HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Acrescimos" HeaderText="Acréscimos" DataFormatString="{0:N2}"
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ValorLiquido" DataFormatString="{0:N2}" HeaderText="Valor Líq."
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                        <div class="painelright" id="divTitulo" runat="server" visible="False" style="width: 42.6%;">
                            <div class="row">
                                <div class="collbl">
                                    Título:
                                </div>
                                <div class="coltxt">
                                    <asp:Label ID="lblTitulo" runat="server" Font-Bold="True" ForeColor="Red" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Forma:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlFormasDePagamento" runat="server" AutoPostBack="True" Font-Names="monospace"
                                        Width="395px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Tipo Cod. Barras:
                                </div>
                                <div class="coltxt">
                                    <asp:CheckBox ID="CkbCodigoDeBarras" runat="server" Enabled="False" Text="Digitado" />
                                </div>
                                <div class="coltxt">
                                    <asp:CheckBox ID="ckPreImpresso" runat="server" Enabled="False" Text="Pré Impresso" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Código de Barras:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCodigoDeBarras" ClientIDMode="Static" runat="server" Enabled="False"
                                        MaxLength="49" Width="350px" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="BtValidarCodBarras" runat="server" Enabled="False" Text="Validar"
                                        CssClass="botao" UseSubmitBehavior="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Data:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataVencimento" runat="server" CssClass="calendario" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Valor:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtValorVencimento" runat="server" CssClass="txtDecimal" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="cmdOkVencimento" CssClass="botao" runat="server" Text="OK" UseSubmitBehavior="False" />
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabContabil">
                    <HeaderTemplate>
                        Contabilização
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProdutoContabilizacao" runat="server" AutoPostBack="True"
                                    Width="600px" />
                            </div>
                            <div class="collbl">
                                Débitos:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblDebito" runat="server" Font-Bold="True" ForeColor="Red" Text="0,00" />
                                &nbsp;
                            </div>
                            <div class="collbl">
                                Créditos:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblCredito" runat="server" Font-Bold="True" ForeColor="Red" Text="0,00" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 300px;">
                            <asp:GridView ID="gridRazao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:BoundField DataField="CodigoConta" HeaderText="Conta">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Descrição">
                                        <ItemTemplate>
                                            <asp:Label ID="Label1" runat="server" Text='<%# Bind("Conta.Titulo") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoCliente" HeaderText="Cliente">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Movimento" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Lote" HeaderText="Lote">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoCusto" HeaderText="Custo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Historico" HeaderText="Hist&#243;rico" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" Width="300px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DebitoOficial" DataFormatString="{0:N2}" HeaderText="D&#233;bito">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CreditoOficial" DataFormatString="{0:N2}" HeaderText="Cr&#233;dito">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Saldo" DataFormatString="{0:N2}" HeaderText="Saldo">
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
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaProdutoCupomFiscal ID="ucConsultaProdutoCupomFiscal" runat="server" />
    <uc:ConsultaOperacoes ID="ucConsultaOperacoes" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaDadosBancarios ID="ucConsultaDadosBancarios" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
    <uc:ComissoesXBaixas ID="ucComissoesXBaixas" runat="server" />
    <uc:ProdutoCupomFiscal ID="ucProdutoCupomFiscal" runat="server" />
    <uc:NFObsProduto ID="ucNFObsProduto" runat="server" />
    <uc:NFOrigem ID="ucNFOrigem" runat="server" />
    <uc:Rateio ID="ucRateio" runat="server" />
    <uc:ConsultaEncargosPlanoDeContas ID="ucConsultaEncargosPlanoDeContas" runat="server" />
    <uc:NotaFiscalReferencial ID="ucNotaFiscalReferencial" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:NFEncargo ID="ucNFEncargo" runat="server" />
    <uc:MonitorCupomFiscal ID="ucMonitorCupomFiscal" runat="server" />
    <uc:ConsultaDeLote ID="ucConsultaLote" runat="server" />
    <uc:EmailNFe ID="ucEmailNFe" runat="server" />
</asp:Content>

