<%@ Page Title="" Language="vb" AutoEventWireup="True" MasterPageFile="~/Principal.Master"
    CodeBehind="NotasFiscaisGerais.aspx.vb" Inherits="NGS.Web.UI.NotasFiscaisGerais" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1360px !important;
        }

        .colw {
            width: 140px;
        }
    </style>
    <script type="text/javascript">
        function pageLoadNotasFiscaisGeraisNova() {

            $("#txtCodigoDeBarras").bind('keypress', function (e) {
                var code = (e.keyCode ? e.keyCode : e.which);
                if (code == 13)
                    return false;
            });

            $("#txtChaveNFe").focusout(function () {
                if ($("#txtChaveNFe").val().replace(/\./g, '').length == 44) {
                    $("#MainContent_lnkVerificarChaveNFE").click();
                    return false;
                }
            });

            $("#txtChaveNFe").live('keypress', function (e) {
                var code = (e.keyCode ? e.keyCode : e.which);
                if (code == 13) {
                    $("#MainContent_lnkVerificarChaveNFE").click();
                    return false;
                }
            });
        }

        $(document).ready(function () {
            pageLoadNotasFiscaisGeraisNova();
            var prmNotasFiscaisGerais = Sys.WebForms.PageRequestManager.getInstance();
            prmNotasFiscaisGerais.add_endRequest(pageLoadNotasFiscaisGeraisNova);
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
                Notas Fiscais Gerais
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
                            <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" AccessKey="C" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkEspelho" Text="Espelho" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRecontabilizar" Text="Recontabilizar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
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
                            Width="590px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Empresa:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlEmpresa" runat="server" AutoPostBack="True" Width="590px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Cliente:
                    </div>
                    <div class="coltxt">
                        <asp:HiddenField ID="hdfCodigoCliente" runat="server" />
                        <asp:TextBox ID="txtNomeCliente" runat="server" ReadOnly="True" TabIndex="3" Width="550px"
                            Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnConsultaClientes" runat="server" TabIndex="4" Text="&gt;" CssClass="btn"
                            UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Chave NF-e:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtChaveNFe" ClientIDMode="Static" runat="server" CssClass="chaveNFe"
                            Width="400px" data-ToolTip="default" ToolTip="Código de identificação da NFe ou CTe perante a Receita Federal." />
                    </div>
                    <div class="coltxt" style="padding-top: 5px">
                        <asp:LinkButton ID="lnkVerificarChaveNFE" CssClass="lnk"
                            data-tooltip="default" ToolTip="Consultar/Validar Chave Eletrônica" runat="server" Text=" &gt; "
                            OnClick="lnkVerificarChaveNFE_Click">
                                    <i class="fa fa-arrow-right seta"></i>
                        </asp:LinkButton>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Safra:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlSafra" runat="server" Width="300px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Observação:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtObservacao" runat="server" MaxLength="200" Width="580px" data-ToolTip="default"
                            ToolTip="Informar caso hajam informações relevantes." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Ações:
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkReaproveitar" ClientIDMode="Static" runat="server" Font-Bold="True"
                            Text="Reaproveitar dados" data-ToolTip="default" ToolTip="Selecionar para aproveitar dados para próximo preenchimento." />
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkEspelho" ClientIDMode="Static" runat="server" Text="Imprimir Espelho"
                            Font-Bold="True" data-ToolTip="default" ToolTip="Selecionar para aproveitar dados para próximo preenchimento." />
                    </div>

                    <div class="coltxt">
                        <asp:CheckBox ID="chkNaturezaDeRendimento" ClientIDMode="Static" runat="server" AutoPostBack="True" OnCheckedChanged="chkNaturezaDeRendimento_CheckedChanged" Text="Informar Natureza De Rendimento"
                            Font-Bold="True" data-ToolTip="default" ToolTip="Selecionar Natureza De Rendimento para informação no SPED REINF." />
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkAGruparNCM" ClientIDMode="Static" runat="server" Text="Agrupar NCM" AutoPostBack="True" OnCheckedChanged="chkImportarXML_CheckedChanged"
                            Font-Bold="True" data-ToolTip="default" ToolTip="Agrupar NCM" GroupName="IMPORTARXML" />
                    </div>
                </div>
                <div class="row" runat="server" visible="true">
                    <div class="collbl">
                        Arquivo:
                    </div>
                    <div class="coltxt">
                        <uc:File ID="ucFile" runat="server" />
                    </div>
                </div>
            </div>
            <div class="painelleft">
                <div class="row">
                    <div class="collbl colw">
                        Situação:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlSituacao" runat="server" Width="472px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl colw">
                        Tipo Documento:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlTipoDeDocumento" runat="server" Width="472px" AutoPostBack="True" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl colw">
                        Pedido:
                    </div>
                    <div class="coltxt">
                        <asp:Label ID="lblPedido" runat="server" Font-Bold="True" Text="0" data-ToolTip="default"
                            ToolTip="Número de cadastro do pedido." />
                    </div>
                    <div class="coltxt">
                        <asp:ImageButton ID="imgExtratoPedido" runat="server" ImageAlign="AbsMiddle" CssClass="imgconsultar"
                            ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" data-ToolTip="default"
                            ToolTip="Visualizar Extrato do Pedido" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl colw">
                        Número/Série:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtNumeroNota" runat="server" TabIndex="7" CssClass="txtNumerico9"
                            Width="88px" data-ToolTip="default" ToolTip="Informar o número da nota." />
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtSerie" runat="server" CssClass="texto" MaxLength="3" Style="text-transform: uppercase"
                            TabIndex="8" Width="36px" data-ToolTip="default" ToolTip="Informar a série da nota." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl colw">
                        Movimento:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMovimento" runat="server" CssClass="calendario" Width="88px"
                            data-ToolTip="default" ToolTip="Data da criação do pedido." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl colw">
                        Data Nota:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDataNota" runat="server" CssClass="calendario" Width="88px" data-ToolTip="default"
                            ToolTip="Data de Criação da Nota." />
                    </div>
                </div>
                <div class="row">
                    <div class="coltxt">
                        <asp:CheckBox ID="chkImportarProdutoUnico" ClientIDMode="Static" runat="server" Text="Importar produto único" AutoPostBack="True" OnCheckedChanged="chkImportarXML_CheckedChanged"
                            Font-Bold="True" data-ToolTip="default" ToolTip="Importar produto único" GroupName="IMPORTARXML" />
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkInformarDados" ClientIDMode="Static" runat="server" Text="Informar dados para importar XML" AutoPostBack="True" OnCheckedChanged="chkImportarXML_CheckedChanged"
                            Font-Bold="True" data-ToolTip="default" ToolTip="Informar dados para importar XML" GroupName="IMPORTARXML" />
                    </div>
                </div>
                <div class="row" id="divNaturezaDeRendimento" runat="server" visible="false">
                    <div class="collbl colw">
                        Natureza do Rendimento:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlNaturezaDeRendimento" AutoPostBack="True" runat="server" Width="472px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl" id="divConferencia" runat="server" visible="false">
                        Conferência
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkConferencia" ClientIDMode="Static" Visible="false" runat="server"
                            Font-Bold="True" Text="Nota Fiscal Conferida" />
                    </div>
                </div>
                <div class="row" id="divNaviosXInvoice" runat="server" visible="False">
                    <div class="collbl">
                        Invoice:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtNaviosXInvoice" runat="server" Width="209px" Enabled="False" ToolTip="Código Invoice e descrição do navio." />
                        <asp:LinkButton class="iconConsultar" ID="lnkConsultarNaviosXInvoice" runat="server" />
                    </div>
                </div>
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="2" Width="100%" Style="clear: both;">
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
                                            <li class="iconNovo" runat="server">
                                                <asp:LinkButton ID="lnkComissoes" Text="Comissões" runat="server" />
                                            </li>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                            <div class="bordagrid" style="height: 300px;">
                                <asp:GridView ID="grdProdutos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <Columns>
                                        <asp:TemplateField ShowHeader="False">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lnkSelecionarProduto" CssClass="lnk"
                                                    data-tooltip="default" ToolTip="Selecionar registro." runat="server" Text=" &gt; "
                                                    OnClick="lnkSelecionarProduto_Click">
                                                        <i class="fa fa-arrow-right seta"></i>
                                                </asp:LinkButton>
                                            </ItemTemplate>
                                            <ItemStyle Width="30px" />
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="CodigoProduto" HeaderText="Produto" HtmlEncode="False"
                                            ReadOnly="True">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" Width="60px" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Sequencia" HeaderText="Sequência" HtmlEncode="False"
                                            ReadOnly="True"></asp:BoundField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:Label ID="Label1" runat="server" Text='<%# Bind("Produto.Nome") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgProdutoDeTerceiro" runat="server" Text='<%# Eval("ProdutoXMLDeTerceiro")%>' Visible="false" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ToolTip='<%# Eval("ProdutoXMLDeTerceiro")%>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="QuantidadeFiscal" DataFormatString="{0:N4}" HeaderText="Quantidade"
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Unitario" DataFormatString="{0:N10}" HeaderText="Unitário"
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ValorTotal" DataFormatString="{0:N2}" HeaderText="Valor"
                                            HtmlEncode="False">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:TemplateField HeaderText="Custo" ShowHeader="False">
                                            <ItemTemplate>
                                                <asp:Label ID="lblCentroDeCusto" runat="server" Text='<%# Eval("Encargos.EncProduto.CentroDeCusto")%>' Visible='<%# Eval("Encargos.EncProduto.CentroDeCusto") > 0%>' ReadOnly="true" />
                                                <asp:Button ID="btnRateio" runat="server" Text=" + " UseSubmitBehavior="False" OnClick="btnRateio_Click"
                                                    Visible='<%# Bind("TemCentroDeCusto") %>' ToolTip="Rateio de Centro de Custo" />
                                            </ItemTemplate>
                                            <ItemStyle Width="30px" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Encargos" ShowHeader="False">
                                            <ItemTemplate>
                                                <asp:Button ID="btnEncargos" runat="server" Text=" + " UseSubmitBehavior="False" ToolTip="Encargos" OnClick="btnEncargos_Click" />
                                            </ItemTemplate>
                                            <ItemStyle Width="30px" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Obs.">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgObsProduto" runat="server"
                                                    Width="18px" Height="18px" ImageUrl="~/images/ico_OBS_ativo.gif" data-ToolTip="default"
                                                    ToolTip="Observação do produto" OnClick="imgObsProduto_Click" />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="CodigoOperacao" HeaderText="Op." />
                                        <asp:BoundField DataField="CodigoSuboperacao" HeaderText="SOp." />
                                        <asp:TemplateField HeaderText="Desc. Operação">
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TextBox1" runat="server" />
                                            </EditItemTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="Label2" runat="server" Text='<%# Bind("Suboperacao.Descricao") %>' />
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:BoundField DataField="CFOP" HeaderText="CFOP">
                                            <HeaderStyle HorizontalAlign="Right" Width="30px" />
                                            <ItemStyle HorizontalAlign="Right" Width="30px" />
                                        </asp:BoundField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgExcluir" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                    OnClick="imgExcluir_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente excluir o produto?');" />
                                            </ItemTemplate>
                                            <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                            <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
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
                                ENCARGOS GERAIS DA NOTA
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
                                Encargo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTributos" runat="server" Font-Names="monospace" Width="300px"
                                    AutoPostBack="True" />
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
                            <div class="collbl">
                                Total Nota:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTotalNota" runat="server" CssClass="txtDecimal" BackColor="#FFFFC0" BorderColor="White"
                                    Font-Bold="True" ForeColor="Red">0</asp:TextBox>
                            </div>
                            <div class="collbl">
                                Total Pago:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTotalPago" runat="server" CssClass="txtDecimal" BackColor="#FFFFC0" Font-Bold="True"
                                    ForeColor="Red">0</asp:TextBox>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Forma:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbFormas" runat="server" Font-Names="monospace" Width="395px" />
                            </div>
                            <div class="collbl">
                                Total Parcelado:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTotalParcelado" runat="server" CssClass="txtDecimal" BackColor="#FFFFC0" BorderColor="White"
                                    Font-Bold="True" ForeColor="Red">0</asp:TextBox>
                            </div>
                            <div class="collbl">
                                Saldo:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSaldoVencimentos" runat="server" CssClass="txtDecimal" BackColor="#FFFFC0" Font-Bold="True"
                                    ForeColor="Red">0</asp:TextBox>
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
                                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grdCondicoes_SelectedIndexChanged">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:CommandField ButtonType="Button" SelectText="&gt;&gt;" ShowSelectButton="True">
                                            <HeaderStyle HorizontalAlign="Right" Width="30px" />
                                            <ItemStyle HorizontalAlign="Center" Width="30px" />
                                        </asp:CommandField>
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
                <ajaxToolkit:TabPanel runat="server" ID="TabNFOrigem">
                    <HeaderTemplate>
                        NF de Origem
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconMais" runat="server">
                                        <asp:LinkButton ID="lnkAddNota" Text="Adicionar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row" id="divPedidoReferenciaRPA" runat="server">
                            <div class="collbl">
                                Período Inicial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataInicio" runat="server" CssClass="calendario" Width="88px"
                                    data-ToolTip="default" ToolTip="Data inicial da busca." />
                            </div>
                            <div class="collbl">
                                Período Final:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataFim" runat="server" CssClass="calendario" Width="88px"
                                    data-ToolTip="default" ToolTip="Data final da busca." />
                            </div>
                            <div class="collbl">
                                Pedido:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoClienteNFReferencial" runat="server" />
                                <asp:TextBox ID="txtPedidoNFReferencial" runat="server" Width="90px" CssClass="txtNumerico"
                                    Style="text-align: right" AutoPostBack="True" ToolTip="Caso o campo pedido esteja bloqueado é porque já existem notas referenciadas. Para liberá-lo será necessário excluí-las." />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnPedidoNFReferencial" runat="server" Text=">" UseSubmitBehavior="False"
                                    CssClass="btn" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 200px;">
                            <asp:GridView ID="grdNotasFretes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField HeaderText="Cliente">
                                        <ItemTemplate>
                                            <asp:Label ID="Label1" runat="server" Text='<%# Bind("DescClienteCompleto") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Depósito">
                                        <ItemTemplate>
                                            <asp:Label ID="Label2" runat="server" Text='<%# Bind("DescDepositoCompleto") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Destino">
                                        <ItemTemplate>
                                            <asp:Label ID="Label3" runat="server" Text='<%# Bind("DescDestinoCompleto") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Operação">
                                        <ItemTemplate>
                                            <asp:Label ID="Label4" runat="server" Text='<%# Bind("DescSubOperacaoCompleto") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Nota" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Tipo">
                                        <ItemTemplate>
                                            <asp:Label ID="Label5" runat="server" Text='<%# Bind("TipoDeDocumento.Descricao") %>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirNF" runat="server" ImageUrl="~/Images/deletar.gif"
                                                Style="border: 0;" OnClick="imgExcluirNF_Click" data-ToolTip="default" ToolTip="Excluir"
                                                OnClientClick="return confirm('Deseja realmente excluir o registro selecionado?');" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                            <asp:GridView runat="server" AutoGenerateColumns="False" CellPadding="4" GridLines="None"
                                ForeColor="#333333" Width="100%" ID="grdNotasReferenciais" ShowFooter="True">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775"></AlternatingRowStyle>
                                <Columns>
                                    <asp:TemplateField HeaderText="Nota">
                                        <ItemTemplate>
                                            <asp:Label ID="lblNota" runat="server" Text='<%# String.Format("{0:N0}", Eval("Nota_Id"))%>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Serie_id" HeaderText="Serie"></asp:BoundField>
                                    <asp:BoundField DataField="Cliente_Id" HeaderText="Cliente"></asp:BoundField>
                                    <asp:BoundField DataField="EndCliente_id" HeaderText="End.">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Nome">
                                        <ItemTemplate>
                                            <asp:Label ID="lblClienteNFReferencial" runat="server" Text='<%# Bind("ParentOrigem.NotaFiscal.cliente.nome")%>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Faz">
                                        <ItemTemplate>
                                            <asp:Label ID="lblFazendaNFReferencial" runat="server" Text='<%# Bind("ParentOrigem.NotaFiscal.cliente.complemento")%>' />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Quantidade" DataFormatString="{0:n2}" HeaderText="Qtd.">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" DataFormatString="{0:n2}" HeaderText="Valor do Frete">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgDelNFReferencial" runat="server" ImageUrl="~/Images/deletar.gif"
                                                OnClick="imgDelNFReferencial_Click" OnClientClick="if(!(confirm('Confirma Exclusão da Nota Referencial?'))) return false;" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                                <EditRowStyle BackColor="#999999"></EditRowStyle>
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></FooterStyle>
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></HeaderStyle>
                                <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White"></PagerStyle>
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333"></RowStyle>
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>
                            </asp:GridView>
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
                <ajaxToolkit:TabPanel runat="server" ID="TabRepresentante">
                    <HeaderTemplate>
                        Dados Adicionais
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="painelleft" style="width: 59%;">
                            <div class="menu_acoes">
                                <div class="acoes">
                                    <ul>
                                        <li class="iconLimpar" runat="server">
                                            <asp:LinkButton ID="lnkLimparRepresentante" Text="Limpar Representante" runat="server" />
                                        </li>
                                    </ul>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Representante:
                                </div>

                                <div class="coltxt">
                                    <asp:TextBox ID="txtRepresentante" runat="server" Enabled="False" Width="600px" />
                                    <asp:HiddenField ID="hdfRepresentante" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnRepresentante" runat="server" OnClick="btnRepresentante_Click"
                                        Text="&gt;" CssClass="btn" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Nome do Representante." />
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
    <uc:ConsultaOperacoes ID="ucConsultaOperacoes" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaDadosBancarios ID="ucConsultaDadosBancarios" runat="server" />
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
    <uc:ComissoesXBaixas ID="ucComissoesXBaixas" runat="server" />
    <uc:ProdutoNFG ID="ucProdutoNFG" runat="server" />
    <uc:InformarDadosProdutoNFG ID="ucInformarDadosProdutoNFG" runat="server" />
    <uc:NFObsProduto ID="ucNFObsProduto" runat="server" />
    <uc:NFOrigem ID="ucNFOrigem" runat="server" />
    <uc:Rateio ID="ucRateio" runat="server" />
    <uc:ConsultaEncargosPlanoDeContas ID="ucConsultaEncargosPlanoDeContas" runat="server" />
    <uc:NotaFiscalReferencial ID="ucNotaFiscalReferencial" runat="server" />
    <uc:ConsultaPedidos ID="ucConsultaPedidos" runat="server" />
    <uc:NFEncargo ID="ucNFEncargo" runat="server" />
    <uc:ConsultarNaviosXInvoice ID="ucConsultarNaviosXInvoice" runat="server" />
</asp:Content>
