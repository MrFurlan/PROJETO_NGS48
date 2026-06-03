<%@ Page Title="" Language="vb" AutoEventWireup="True" MasterPageFile="~/Principal.Master"
    CodeBehind="NotasDeTerceiro.aspx.vb" Inherits="NGS.Web.UI.NotasDeTerceiro" %>

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
        }

        function TotalizarFrete() {
            var txtTarifaFrete = $("#<%=txtTarifaFrete.ClientID%>", "#divPesoDeChegada");
                    var txtBrutoDeChegada = $("#<%=txtBrutoDeChegada.ClientID%>", "#divPesoDeChegada");
            var txtValorDoFrete = $("#<%=txtValorDoFrete.ClientID%>", "#divPesoDeChegada");

            if (txtTarifaFrete.val() != undefined && txtTarifaFrete.val() != '' && txtBrutoDeChegada.val() != undefined && txtBrutoDeChegada.val() != '') {
                var tarifaFrete = parseFloat(txtTarifaFrete.val().replace('.', '').replace(',', '.'));
                var brutoDeChegada = parseFloat(txtBrutoDeChegada.val().replace('.', '').replace(',', '.'));
                var valorDoFrete = brutoDeChegada * tarifaFrete / 1000;
            }
            txtValorDoFrete.val(valorDoFrete.toFixed(2).replace(".", ","));
        }

        $(document).ready(function () {
            pageLoadNotasFiscaisGeraisNova();
            var prmNotasFiscaisGerais = Sys.WebForms.PageRequestManager.getInstance();
            prmNotasFiscaisGerais.add_endRequest(pageLoadNotasFiscaisGeraisNova);
        });

    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngNotasDeTerceiro" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlNotasDeTerceiro" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Notas Fiscais de Terceiro
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
                            <asp:LinkButton ID="lnkImportar" Text="Importar" AccessKey="I" runat="server" />
                        </li>
                        <li class="iconConsultar" runat="server">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" AccessKey="C" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
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
                        Empresa:
                    </div>
                    <div class="coltxt">
                        <asp:HiddenField ID="hdfCodigoEmpresa" runat="server" />
                        <asp:TextBox ID="txtNomeEmpresa" runat="server" ReadOnly="True" TabIndex="3" Width="550px"
                            Enabled="false" />
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
                </div>
                <div class="row">
                    <div class="collbl">
                        Chave NF-e:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtChaveNFe" ClientIDMode="Static" runat="server" CssClass="chaveNFe"
                            Width="400px" data-ToolTip="default" ToolTip="Código de identificação da NFe ou CTe perante a Receita Federal." />
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
                                                    data-tooltip="default" ToolTip="Selecionar registro." runat="server" Text=" &gt; ">
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
                                                <asp:Button ID="btnRateio" runat="server" Text=" + " UseSubmitBehavior="False"
                                                    Visible='<%# Bind("TemCentroDeCusto") %>' ToolTip="Rateio de Centro de Custo" />
                                            </ItemTemplate>
                                            <ItemStyle Width="30px" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Encargos" ShowHeader="False">
                                            <ItemTemplate>
                                                <asp:Button ID="btnEncargos" runat="server" Text=" + " UseSubmitBehavior="False" ToolTip="Encargos" />
                                            </ItemTemplate>
                                            <ItemStyle Width="30px" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Obs.">
                                            <HeaderStyle HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:ImageButton ID="imgObsProduto" runat="server"
                                                    Width="18px" Height="18px" ImageUrl="~/images/ico_OBS_ativo.gif" data-ToolTip="default"
                                                    ToolTip="Observação do produto" />
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
                                                    data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente excluir o produto?');" />
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
                <ajaxToolkit:TabPanel runat="server" Visible="false" ID="TabPesoDeChegada">
                    <HeaderTemplate>
                        Peso de Chegada
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="collbluc">
                                Bruto de Saída:
                            </div>
                            <div class="coltxt" style="width: 156px;">
                                <asp:TextBox ID="txtBrutoDeSaida" runat="server" CssClass="txtInteiro" Width="125px"
                                    Enabled="false" />
                            </div>
                            <div class="collbluc">
                                Bruto de Chegada:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtBrutoDeChegada" runat="server" CssClass="txtInteiro" AutoPostBack="True"
                                    MaxLength="6" OnTextChanged="txtBrutoDeChegada_TextChanged" Style="color: red; text-align: right"
                                    Width="125px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Data de Chegada:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataDeChegada" runat="server" CssClass="calendario" Width="125px" />
                            </div>
                            <div class="collbluc">
                                Quebra / Sobra:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSaldo" runat="server" CssClass="txtInteiro" Width="125px" Enabled="false" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Descontos:
                            </div>
                            <div class="coltxt" style="width: 156px;">
                                <asp:TextBox ID="txtDesconto" runat="server" CssClass="txtInteiro" AutoPostBack="True"
                                    Width="125px" MaxLength="9" OnTextChanged="txtDesconto_TextChanged" />
                            </div>
                            <div class="collbluc">
                                Líquido:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtLiquido" runat="server" CssClass="txtInteiro" Width="125px" MaxLength="9"
                                    Enabled="false" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Tarifa do Frete:
                            </div>
                            <div class="coltxt" style="width: 156px;">
                                <asp:TextBox ID="txtTarifaFrete" runat="server" ToolTip="Tarifa do Frete por Tonelada."
                                    onblur="TotalizarFrete();" CssClass="txtDecimal" Width="125px" />
                            </div>
                            <div class="collbluc">
                                Valor Total do Frete:
                            </div>
                            <div class="coltxt" style="width: 156px;">
                                <asp:TextBox ID="txtValorDoFrete" runat="server" ToolTip="Valor Total do Frete."
                                    Enabled="false" CssClass="txtDecimal" Width="125px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                <asp:CheckBox ID="chkSinistro" runat="server" Text="Possui sinistro?" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="BtnGravar" runat="server" CssClass="botao" Text="Gravar"
                                    UseSubmitBehavior="False" OnClick="BtnGravar_Click" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPedidosxNotas ID="ucConsultaPedidosXNotas" runat="server" />
</asp:Content>
