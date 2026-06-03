<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master" CodeBehind="OrdemDeProducao.aspx.vb" Inherits="NGS.Web.UI.OrdemDeProducao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }
    </style>
    <script type="text/javascript">
        function pageLoadOrdemDeProducao() {
            $("#txtLote").setMask("lote-producao");
        }

        $(document).ready(function () {
            pageLoadOrdemDeProducao();
            var prmItemProducao = Sys.WebForms.PageRequestManager.getInstance();
            prmItemProducao.add_endRequest(pageLoadOrdemDeProducao);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngOrdemDeProducao" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlOrdemDeProducao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Ordem de Produção
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%" Style="margin-top: 4px;">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Dados para Produção
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                            OnClientClick="if(!confirm('Deseja realmente cancelar a Ordem de Produção?')) return false;" />
                                    </li>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                                        <ul>
                                            <li class="iconPdf">
                                                <asp:LinkButton ID="lnkImprimirOrdem" runat="server" Text="Ordem" />
                                            </li>
                                            <li class="iconPdf">
                                                <asp:LinkButton ID="lnkImprimirLaudo" runat="server" Text="Laudo" />
                                            </li>
                                        </ul>
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconPdf" ID="lnkImprimirLaudoManual" runat="server" Text="Laudo Manual" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                    <li runat="server" style="float: right;">
                                        <div class="row" style="margin-top: 0;">
                                            <div class="coltxt">
                                                <asp:Image ID="Image2" runat="server" Height="20px" ImageAlign="AbsMiddle" ImageUrl="~/Images/man2.png"
                                                    data-ToolTip="default" ToolTip="Usuário Lançamento" Width="18px" />
                                            </div>
                                            <div class="coltxt">
                                                <asp:DropDownList ID="ddlUsuarios" runat="server" Width="200px">
                                                </asp:DropDownList>
                                            </div>
                                        </div>
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Ordem:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSequencia" CssClass="txtNumerico" runat="server" Enabled="False" data-ToolTip="default"
                                    ToolTip="Sequência de Produção." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUnidadeNegocio" runat="server" Enabled="False" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidadeNegocio_SelectedIndexChanged" Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmpresa" runat="server" Enabled="False" Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data de Produção:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataProducao" CssClass="calendario" runat="server" Enabled="False" data-ToolTip="default"
                                    ToolTip="Data inicial do lançamento." />
                            </div>
                            <div class="collbl">
                                Data de Validade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataValidade" CssClass="calendario" runat="server" Enabled="False" data-ToolTip="default"
                                    ToolTip="Data de Validade." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Gerar Estoque:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbNao" runat="server" Checked="True" Enabled="False" GroupName="gEstoque"
                                    Text="NÃO" data-ToolTip="default" ToolTip="Não gerar Estoque." AutoPostBack="True" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbSim" runat="server" Enabled="False" GroupName="gEstoque"
                                    Text="SIM" data-ToolTip="default" ToolTip="Gerar Estoque." AutoPostBack="True" />
                            </div>
                            <div class="coltxt">&nbsp;&nbsp;</div>
                            <div id="dataEstoque" runat="server" visible="False">
                                <div class="collbl">
                                    Data para Estoque:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataEstoque" CssClass="calendario" runat="server" data-ToolTip="default"
                                        ToolTip="Data para Estoque." />
                                </div>
                            </div>
                        </div>
                        <ajaxToolkit:TabContainer ID="TabInfConsumos" runat="server" ActiveTabIndex="0" Width="100%" Style="margin-top: 2px;">
                            <ajaxToolkit:TabPanel ID="TabProducao" runat="server">
                                <HeaderTemplate>
                                    Produção
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="subtitulodiv">
                                        Dados para Produção
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Quantidade:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtQuantidadeProducao" runat="server" Enabled="False" CssClass="txtDecimal4" data-ToolTip="default" ToolTip="Quantidade de Produção." />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbluc">
                                            Grupo:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlGrupoProdutoProducao" runat="server" Enabled="False" AutoPostBack="True" Width="576px" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:LinkButton ID="lnkBuscaProdutoProducao" runat="server" Enabled="False" Height="20px" Width="20px">
                                                <asp:Image ID="imgSearch" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                                                    ToolTip="Consulta Produto" />
                                            </asp:LinkButton>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbluc">
                                            Produto:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlProdutosProducao" runat="server" Enabled="False" AutoPostBack="True" Width="596px" />
                                        </div>
                                        <div class="collbluc">
                                            Peso Produto:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlUnidadeComercializacao" runat="server" Enabled="False" Width="150px" AutoPostBack="True" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:ImageButton ID="imgAdicionarDadosParaProducao" runat="server" ImageUrl="~/Images/detalhes.png" Style="border: 0;"
                                                OnClick="imgAdicionarDadosParaProducao_Click" data-ToolTip="default" ToolTip="Adicionar dados para a produção" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="bordagrid" style="width: 100%; height: 170px;">
                                            <asp:GridView ID="gridProducao" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridProducao_SelectedIndexChanged">
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <EditRowStyle BackColor="#999999" />
                                                <Columns>
                                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                                    <asp:BoundField DataField="Produto" HeaderText="Produto" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                                        <ItemStyle HorizontalAlign="Left" Width="100px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="NomeProduto" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="CodigoUnidadeComercializacao" HeaderText="Unidade" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="FatorConversao" HeaderText="Peso" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:TemplateField>
                                                        <ItemTemplate>
                                                            <asp:ImageButton ID="imgRemoverProducao" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                                OnClick="imgRemoverProducao_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente remover o Produto? Todos os outros serão removidos!');" />
                                                        </ItemTemplate>
                                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    </asp:TemplateField>
                                                    <asp:TemplateField HeaderText="Ajuste">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtAjusteProducao" CssClass="txtDecimal4" Width="100px" Enabled="false" runat="server" Text='<%# Eval("AjusteProducao", "{0:N4}")%>' />
                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="Lote" HeaderText="Lote" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                    </div>
                                    <div class="row" id="divAjusteProducao" runat="server">
                                        <div class="menu_acoes" runat="server">
                                            <div class="acoes">
                                                <ul>
                                                    <li runat="server">
                                                        <asp:LinkButton class="iconAtualizar" ID="lnkAjustaProducao" Text="Ajustar" runat="server" />
                                                    </li>
                                                    <li runat="server">
                                                        <asp:LinkButton class="iconNovo" ID="lnkConfirmaProducao" Text="Confirmar" runat="server" />
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabConsumos" runat="server">
                                <HeaderTemplate>
                                    Consumos
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="subtitulodiv">
                                        Dados para Consumo(s)
                                    </div>
                                    <div id="divConsumo" class="painelleft" style="width: 100%;" runat="server">
                                        <div class="bordagrid" style="height: 260px;">
                                            <asp:GridView ID="gridConsumo" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <EditRowStyle BackColor="#999999" />
                                                <Columns>
                                                    <asp:TemplateField ShowHeader="False">
                                                        <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                                        <ItemStyle HorizontalAlign="Center" Width="30px" />
                                                        <ItemTemplate>
                                                            <asp:ImageButton ID="imgSelecionaLote" runat="server" ImageUrl="~/images/select.jpg"
                                                                OnClick="imgSelecionaLote_Click" />
                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="Produto" HeaderText="Produto" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                                        <ItemStyle HorizontalAlign="Left" Width="100px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="NomeProduto" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="NumeroDoLote" HeaderText="Lote" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Percentual" HeaderText="%" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                    <asp:TemplateField HeaderText="Sinal">
                                                        <HeaderStyle HorizontalAlign="Center" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:DropDownList ID="DdlSinal" runat="server" Enabled="false" Width="70px">
                                                                <asp:ListItem Value="+">+ Mais</asp:ListItem>
                                                                <asp:ListItem Value="-">- Menos</asp:ListItem>
                                                            </asp:DropDownList>
                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                    <asp:TemplateField HeaderText="Ajuste">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtAjusteConsumo" CssClass="txtDecimal4" Width="90px" Enabled="false" runat="server" Text='<%# Eval("AjusteConsumo", "{0:N4}")%>' />
                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                    <asp:TemplateField>
                                                        <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                                        <ItemStyle HorizontalAlign="Center" Width="30px" />
                                                        <ItemTemplate>
                                                            <asp:ImageButton ID="imgExcluir" runat="server" Height="15px" ImageUrl="~/images/erro.jpg"
                                                                ToolTip="Excluir Consumo" Width="15px" Style="cursor: pointer" OnClick="imgExcluir_Click" />
                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                    </div>
                                    <div class="painelright" id="divConsumoLote" runat="server" visible="False">
                                        <div class="bordagrid" style="height: 260px;">
                                            <asp:GridView ID="gridLoteDeFornecedor" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <EditRowStyle BackColor="#999999" />
                                                <Columns>
                                                    <asp:BoundField DataField="Lote" HeaderText="Lote" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Validade" HeaderText="Validade" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Quantidade" HeaderText="Saldo" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:TemplateField HeaderText="Consumo">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtConsumoLote" CssClass="txtDecimal4" Width="80px" Enabled="false" runat="server" Text='<%# Eval("Consumo", "{0:N4}")%>' />
                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                        <div class="painelright">
                                            <asp:ImageButton ID="imgConfirmar" runat="server" ImageUrl="~/images/confirmar.gif" OnClick="imgConfirmar_Click" />
                                        </div>
                                    </div>
                                    <div class="row" id="divAjusteConsumo" runat="server">
                                        <div class="menu_acoes" runat="server">
                                            <div class="acoes">
                                                <ul>
                                                    <li runat="server">
                                                        <asp:LinkButton class="iconAtualizar" ID="lnkAjustaConsumo" Text="Ajustar" runat="server" />
                                                    </li>
                                                    <li runat="server">
                                                        <asp:LinkButton class="iconNovo" ID="lnkConfirmaConsumo" Text="Confirmar" runat="server" />
                                                    </li>
                                                </ul>
                                            </div>
                                        </div>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabInsumos" runat="server">
                                <HeaderTemplate>
                                    Insumos
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="subtitulodiv">
                                        Dados para Insumo(s)
                                    </div>
                                    <div id="divInsumo" class="painelleft" style="width: 100%;" runat="server">
                                        <div class="bordagrid" style="height: 290px;">
                                            <asp:GridView ID="gridInsumo" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                <EditRowStyle BackColor="#999999" />
                                                <Columns>
                                                    <asp:BoundField DataField="Produto" HeaderText="Produto" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                                        <ItemStyle HorizontalAlign="Left" Width="100px" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="NomeProduto" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:BoundField DataField="Base" HeaderText="Base" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Left" />
                                                        <ItemStyle HorizontalAlign="Left" />
                                                    </asp:BoundField>
                                                    <asp:TemplateField HeaderText="Quantidade">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtQuantidadeInsumo" CssClass="txtDecimal4" Enabled="false" Width="50px" runat="server" Text='<%# Eval("Quantidade", "{0:N4}")%>' />
                                                        </ItemTemplate>
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="Estoque" HeaderText="Estoque" HtmlEncode="False">
                                                        <HeaderStyle HorizontalAlign="Right" />
                                                        <ItemStyle HorizontalAlign="Right" />
                                                    </asp:BoundField>
                                                </Columns>
                                            </asp:GridView>
                                        </div>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabIncluirParaConsumo" runat="server">
                                <HeaderTemplate>
                                    Dados Para Consumo(s) | Insumo(s)
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="menu_acoes">
                                        <div class="acoes">
                                            <ul>
                                                <li runat="server">
                                                    <asp:LinkButton class="iconNovo" ID="lnkGravarConsumo" Text="Gravar" runat="server" />
                                                </li>
                                                <li runat="server">
                                                    <asp:LinkButton class="iconLimpar" ID="lnkLimparConsumo" Text="Limpar" runat="server" />
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                    <div class="subtitulodiv">
                                        Dados Para Consumo(s)
                                    </div>
                                    <div class="row">
                                        <div class="painelleft">
                                            <div class="collbl">
                                                Quantidade:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtQuantidadeAlt" runat="server" Enabled="false" CssClass="txtDecimal4" data-ToolTip="default" ToolTip="Quantiddae de Consumo." />
                                            </div>
                                            <div class="coltxt">
                                                <asp:Button ID="btnQuantidade" OnClick="btnQuantidade_Click" runat="server" Visible="false" UseSubmitBehavior="False"
                                                    Text="&gt;" CssClass="btn" />
                                            </div>
                                        </div>
                                        <div class="painelright">
                                            <asp:Label ID="lblQuantidadeAlt" runat="server" Enabled="false" Text="" Width="250px" BackColor="#FFFFC0" Font-Bold="True" ForeColor="Red" />
                                        </div>
                                        <div class="collbl">
                                            Sinal:
                                        </div>
                                        <div class="coltxt">
                                            <asp:RadioButton ID="rbMais" runat="server" Checked="True" GroupName="gSinal"
                                                Text="+" data-ToolTip="default" ToolTip="Sinal de mais" AutoPostBack="True" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:RadioButton ID="rbMenos" runat="server" GroupName="gSinal"
                                                Text="-" data-ToolTip="default" ToolTip="Sinal de menos" AutoPostBack="True" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbluc">
                                            Grupo:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlGrupoProdutoConsumo" runat="server" AutoPostBack="True" Width="576px" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:LinkButton ID="lnkBuscaProdutoConsumo" runat="server" Height="20px" Width="20px">
                                                <asp:Image ID="Image1" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                                                    ToolTip="Consultar Produto para Consumo" />
                                            </asp:LinkButton>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbluc">
                                            Produto:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlProdutosConsumo" Enabled="false" runat="server" AutoPostBack="True" Width="596px" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Número do Lote:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtLoteAlt" runat="server" Enabled="false" ClientIDMode="Static" MaxLength="20" data-ToolTip="default"
                                                ToolTip="Número do Lote." />
                                            <asp:ImageButton ID="imgSelecionaLoteAlt" runat="server" ImageUrl="~/images/search.png" ToolTip="Buscar Número(s) do(s) Lote(s)." OnClick="imgSelecionaLoteAlt_Click" />
                                        </div>
                                        <div class="collbl">
                                            Data de Validade:
                                        </div>
                                        <div class="coltxt">
                                            <asp:TextBox ID="txtValidadeAlt" CssClass="calendario" runat="server" Enabled="false" data-ToolTip="default"
                                                ToolTip="Data de Validade." />
                                        </div>
                                    </div>
                                    <div class="bordagrid" style="height: 100px;">
                                        <asp:GridView ID="gridConsumoAlt" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EditRowStyle BackColor="#999999" />
                                            <Columns>
                                                <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                                    <HeaderStyle Width="50px" />
                                                    <ItemStyle Width="50px" />
                                                </asp:CommandField>
                                                <asp:BoundField DataField="Produto" HeaderText="Produto" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                                    <ItemStyle HorizontalAlign="Left" Width="100px" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="NomeProduto" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Lote" HeaderText="Lote" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Validade" HeaderText="Validade" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Sinal" HeaderText="Sinal" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:ImageButton ID="imgRemoverConsumo" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                            OnClick="imgRemoverConsumo_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente remover o Produto?');" />
                                                    </ItemTemplate>
                                                    <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:TemplateField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                    <div class="subtitulodiv">
                                        Dados Para Insumo(s)
                                    </div>
                                    <div class="row">
                                        <div class="painelleft" style="width: 85%;">
                                            <div class="collbl">
                                                Quantidade:
                                            </div>
                                            <div class="coltxt">
                                                <asp:TextBox ID="txtQuantidadeIns" runat="server" CssClass="txtDecimal4" data-ToolTip="default" ToolTip="Quantidade do Insumo." />
                                            </div>
                                            <div>
                                                <asp:Button ID="btnQuantidadeIns" OnClick="btnQuantidadeIns_Click" runat="server" UseSubmitBehavior="False"
                                                    Text="&gt;" CssClass="btn" />
                                            </div>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbluc">
                                            Grupo:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlGrupoProdutoInsumo" runat="server" AutoPostBack="True" Width="576px" />
                                        </div>
                                        <div class="coltxt">
                                            <asp:LinkButton ID="lnkBuscaProdutoInsumo" runat="server" Height="20px" Width="20px">
                                                <asp:Image ID="Image3" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                                                    ToolTip="Consultar Insumo" />
                                            </asp:LinkButton>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbluc">
                                            Produto:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlProdutosInsumo" Enabled="false" runat="server" AutoPostBack="True" Width="596px" />
                                        </div>
                                    </div>
                                    <div class="bordagrid" style="height: 100px;">
                                        <asp:GridView ID="gridInsumoAlt" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                            <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                            <EditRowStyle BackColor="#999999" />
                                            <Columns>
                                                <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                                    <HeaderStyle Width="50px" />
                                                    <ItemStyle Width="50px" />
                                                </asp:CommandField>
                                                <asp:BoundField DataField="Produto" HeaderText="Produto" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Left" Width="100px" />
                                                    <ItemStyle HorizontalAlign="Left" Width="100px" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="NomeProduto" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" HtmlEncode="False">
                                                    <HeaderStyle HorizontalAlign="Left" />
                                                    <ItemStyle HorizontalAlign="Left" />
                                                </asp:BoundField>
                                                <asp:TemplateField>
                                                    <ItemTemplate>
                                                        <asp:ImageButton ID="imgRemoverInsumo" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                            OnClick="imgRemoverInsumo_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente remover o Produto?');" />
                                                    </ItemTemplate>
                                                    <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                    <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                </asp:TemplateField>
                                            </Columns>
                                        </asp:GridView>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                        </ajaxToolkit:TabContainer>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Especificações do Produto
                    </HeaderTemplate>
                    <ContentTemplate>
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
                                    <asp:BoundField DataField="CodigoProdutoProducao" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Código">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Especificação">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="FaixaInicial" HeaderText="Faixa Inicial" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="FaixaFinal" HeaderText="Faixa Final" DataFormatString="{0:N2}">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Resultado" Visible="true">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtResultado" CssClass="txtDecimal2" Width="60px" Enabled="false" runat="server" Text='<%# Eval("Resultado", "{0:N2}")%>' />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <div class="menu_acoes" id="divEspecificacao" runat="server">
                            <div class="acoes">
                                <ul style="float: right;">
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkAjustaEspecificacao" Text="Ajustar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkConfirmaEspecificacao" Text="Confirmar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel3" runat="server">
                    <HeaderTemplate>
                        EPI
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="subtitulodiv">
                            Equipamento de Proteção Individual
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                EPI:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEPI" runat="server" Enabled="false" AutoPostBack="True" Width="596px" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 100px;">
                            <asp:GridView ID="gridEPI" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="descricao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgRemoverEPI" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                OnClick="imgRemoverEPI_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente remover esse EPI?');" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel4" runat="server">
                    <HeaderTemplate>
                        Embalagens
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="subtitulodiv">
                            Embalagens
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Embalagem:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmbalagem" runat="server" Enabled="false" Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Quantidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQuantidadeEmbalagem" runat="server" Enabled="false" CssClass="txtNumerico" data-ToolTip="default" ToolTip="Quantidade de Embalagem." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUnidadeDeComercializacao" runat="server" Enabled="false" Width="596px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Capacidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCapacidade" runat="server" Enabled="false" CssClass="txtDecimal2" data-ToolTip="default" ToolTip="Capacidade de Embalagem." />
                            </div>
                            <div>
                                <asp:ImageButton ID="imgAddEmbalagem" runat="server" CssClass="imgconsultar" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_mais.png"
                                    OnClick="imgAddEmbalagem_Click" data-ToolTip="default" ToolTip="Adicionar Embalagem" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 100px;">
                            <asp:GridView ID="gridEmbalagem" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Quantidade" HeaderText="Quantidade" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoUnidade" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Capacidade" HeaderText="Capacidade" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgRemoverEmbalagem" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                OnClick="imgRemoverEmbalagem_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente remover essa Embalagem?');" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel5" runat="server">
                    <HeaderTemplate>
                        Procedimentos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="subtitulodiv">
                            Procedimentos
                        </div>
                        <div class="row">
                            <div class="collbluc">
                                Procedimento:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProcedimento" runat="server" Enabled="false" AutoPostBack="True" Width="596px" />
                            </div>
                        </div>
                        <div class="bordagrid" style="height: 100px;">
                            <asp:GridView ID="gridProcedimento" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="descricao" HeaderText="Descri&#231;&#227;o" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgRemoverProcedimento" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                OnClick="imgRemoverProcedimento_Click" data-ToolTip="default" ToolTip="Excluir" OnClientClick="return confirm('Deseja realmente remover esse Procedimento?');" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel6" runat="server">
                    <HeaderTemplate>
                        Observações
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="collbl">
                                Observações:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacoes" runat="server" Width="750px" Height="540px" TextMode="MultiLine"
                                    data-ToolTip="default" ToolTip="Preencher quando houver alguma observação relevante." />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel7" runat="server">
                    <HeaderTemplate>
                        Estoque
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid" style="height: 200px;">
                            <asp:GridView ID="gridEstoque" runat="server" AutoGenerateColumns="False" CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Produto" HeaderText="Produto" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                                <Columns>
                                    <asp:BoundField DataField="Operacao" HeaderText="OP" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                                <Columns>
                                    <asp:BoundField DataField="SubOperacao" HeaderText="SO" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                                <Columns>
                                    <asp:BoundField DataField="Tipo" HeaderText="Tipo" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                                <Columns>
                                    <asp:BoundField DataField="Movimento" HeaderText="Movimento" DataFormatString="{0:dd/MM/yyyy}" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                                <Columns>
                                    <asp:BoundField DataField="Estoque" HeaderText="Estoque" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                                <Columns>
                                    <asp:BoundField DataField="ProdutoDerivado" HeaderText="Produto Derivado" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                                <Columns>
                                    <asp:BoundField DataField="Entradas" HeaderText="Entradas" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                                <Columns>
                                    <asp:BoundField DataField="Saidas" HeaderText="Saídas" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel8" runat="server">
                    <HeaderTemplate>
                        Relatório
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="collbl">
                                Data inicial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data inicial para o Relatório." />
                            </div>
                            <div class="collbl">
                                Data final:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" data-ToolTip="default"
                                    ToolTip="Data final para o Relatório." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tipo:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbDtMov" runat="server" Checked="True" GroupName="gData"
                                    Text="Movimento" data-ToolTip="default" ToolTip="Por data de Movimento." />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbDtVen" runat="server" GroupName="gData"
                                    Text="Validade" data-ToolTip="default" ToolTip="Por data de Validade." />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbDtEstoque" runat="server" GroupName="gData"
                                    Text="Estoque" data-ToolTip="default" ToolTip="Por data de Estoque." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Ordem:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbOrdem" runat="server" Checked="True" GroupName="gOrdem"
                                    Text="Ordem" data-ToolTip="default" ToolTip="Ordena por Ordem." />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbOPrd" runat="server" GroupName="gOrdem"
                                    Text="Produto" data-ToolTip="default" ToolTip="Ordena por Produto." />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbOPrdNome" runat="server" GroupName="gOrdem"
                                    Text="Nome do Produto" data-ToolTip="default" ToolTip="Ordena por Nome do Produto." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tipo:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbAberta" runat="server" Checked="True" GroupName="gTipo"
                                    Text="Aberta(s)" data-ToolTip="default" ToolTip="Ordens Aberta(s)." />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbEncerrada" runat="server" GroupName="gTipo"
                                    Text="Encerrada(s)" data-ToolTip="default" ToolTip="Ordens Encerrada(s)." />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rbCancelada" runat="server" GroupName="gTipo"
                                    Text="Cancelada(s)" data-ToolTip="default" ToolTip="Ordens Cancelada(s)." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt" style="width: 120px;">&nbsp;</div>
                            <div class="coltxt">
                                <asp:ImageButton ID="imgCRelatorio" runat="server" ImageUrl="~/images/confirmar.gif" OnClick="imgCRelatorio_Click" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
    <uc:ConsultaOrdemDeProducao ID="ucConsultaOrdemDeProducao" runat="server" />
    <uc:LaudoManual ID="ucLaudoManual" runat="server" />
    <uc:ConsultaDeLote ID="ucConsultaLote" runat="server" />
</asp:Content>
