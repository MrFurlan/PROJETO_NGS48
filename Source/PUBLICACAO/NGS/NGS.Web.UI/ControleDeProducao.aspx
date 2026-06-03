<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ControleDeProducao.aspx.vb" Inherits="NGS.Web.UI.ControleDeProducao" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngProducao" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlProducao" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Produção e Estoques
            </div>
            <ajaxToolkit:TabContainer ID="tabcProducao" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel1" HeaderText="TabPanel1">
                    <HeaderTemplate>
                        Manutenção
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
                                            OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:CheckBox ID="chkManterDados" Text="Manter dados do lançamento"
                                            runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbUnidade" runat="server" OnSelectedIndexChanged="cmbUnidade_SelectedIndexChanged"
                                    Font-Names="monospace" AutoPostBack="True" Width="702px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbEmpresa" runat="server" Font-Names="monospace" Width="702px"
                                    AutoPostBack="True" OnSelectedIndexChanged="cmbEmpresa_SelectedIndexChanged" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                <asp:CheckBox ID="CkCDeposito" runat="server" Text="Depósito:" data-ToolTip="default" ToolTip="" />
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoDeposito" runat="server" />
                                <asp:TextBox ID="txtDeposito" runat="server" Font-Names="monospace" Width="662px"
                                    Enabled="False" data-ToolTip="default" ToolTip="Local de armazenamento da mercadoria." />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdConsultaDeposito" OnClick="cmdConsultaDeposito_Click" runat="server"
                                    Text=">" Width="20px" CssClass="btn" data-ToolTip="default" ToolTip="Local de armazenamento da mercadoria." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Operação:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="txtCodigoOperacao" runat="server" />
                                <asp:TextBox ID="txtOperacao" runat="server" Font-Names="monospace" Width="662px"
                                    Enabled="False" data-ToolTip="default" ToolTip="Selecionar o tipo de operação do pedido." />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdOperacao" OnClick="cmdOperacao_Click" runat="server" Text=">"
                                    Width="20px" CssClass="btn" data-ToolTip="default" ToolTip="Selecionar o tipo de operação do pedido." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Movimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMovimento" CssClass="calendario" runat="server" Width="64px" data-ToolTip="default" ToolTip="Data do lançamento em estoque." />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMovimentoHora" runat="server" Width="34px" ReadOnly="True" data-ToolTip="default" ToolTip="Data/Hora do lançamento em estoque.">00:00</asp:TextBox>
                            </div>
                            <div class="collbl">
                                Tipo de Estoque:
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="chkFisico" runat="server" Text="Físico" ValidationGroup="TipoEstoque" data-ToolTip="default" ToolTip="Informar se o estoque é físico ou fiscal." />
                            </div>
                            <div class="coltxt">
                                <asp:CheckBox ID="chkFiscal" runat="server" Text="Fiscal" ValidationGroup="TipoEstoque" data-ToolTip="default" ToolTip="Informar se o estoque é físico ou fiscal." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Etapa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbEtapas" runat="server" Font-Names="monospace" Width="702px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbGruposProdutos" runat="server" OnSelectedIndexChanged="cmbGruposProdutos_SelectedIndexChanged"
                                    Font-Names="monospace" AutoPostBack="True" Width="348px" />
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbProdutos" runat="server" OnSelectedIndexChanged="cmbProdutos_SelectedIndexChanged"
                                    Font-Names="monospace" AutoPostBack="True" Width="350px" />
                            </div>
                            <div class="coltxt">
                                <asp:LinkButton ID="lnkBuscaProduto" runat="server" Height="20px" Width="20px">
                                    <asp:Image ID="imgSearch" ImageUrl="~/Images/search.png" runat="server" data-ToolTip="default"
                                        ToolTip="Consulta Produto" />
                                </asp:LinkButton>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto Derivado:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbGruposProdutosDerivados" runat="server" OnSelectedIndexChanged="cmbGruposProdutosDerivados_SelectedIndexChanged"
                                    Font-Names="monospace" AutoPostBack="True" Width="348px" />
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="cmbProdutosDerivados" runat="server" Font-Names="monospace"
                                    Width="350px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Lote / Classificação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlLote" runat="server" OnSelectedIndexChanged="ddlLote_SelectedIndexChanged"
                                    Width="348px" AutoPostBack="True" Font-Names="monospace" />
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlLoteClassificacao" runat="server" Width="350px" Font-Names="monospace" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Embalagem:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEmbalagem" runat="server" Width="702px" Font-Names="monospace" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Quantidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQuantidade" CssClass="txtDecimal4" runat="server" data-ToolTip="default" ToolTip="Informar a quantidade em Kg." />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="LblUn" runat="server" data-ToolTip="default" ToolTip="" />
                            </div>
                            <div id="divCentroDeCusto" runat="server" visible="false">
                                <div class="collbl">
                                    Centro de Custo:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlCentroDeCusto" runat="server" Width="463px" data-ToolTip="default"
                                        ToolTip="Selecionar o Centro de Custo." />
                                </div>
                            </div>
                        </div>
                        <div id="idControlarNumeroDoLote" runat="server" visible="false">
                            <div class="collbl">
                                Nº do Lote:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNumeroDoLote" MaxLength="30" runat="server" Enabled="false" data-ToolTip="default" />
                                <asp:ImageButton ID="imgSelecionaLote" runat="server" Visible="false" ImageUrl="~/images/search.png" ToolTip="Buscar Número(s) do(s) Lote(s)." OnClick="imgSelecionaLote_Click" />
                            </div>
                            <div class="collbl">
                                Fabricado:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataFabricacao" CssClass="calendario" runat="server" Enabled="false" Width="64px" data-ToolTip="default" ToolTip="Data de Fabricação do Lote." />
                            </div>

                            <div class="collbl">
                                Validade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataValidade" CssClass="calendario" runat="server" Enabled="false" Width="64px" data-ToolTip="default" ToolTip="Data de Validade do Lote." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacao" runat="server" TextMode="MultiLine" Width="702px" data-ToolTip="default" ToolTip="Preencher quando houver alguma observação relevante." />
                            </div>
                        </div>
                        <div class="row" id="idAnalises" runat="server" visible="false">
                            <div class="subtitulodiv">
                                Análises:
                            </div>
                            <div class="bordagrid">
                                <asp:GridView ID="gridAnalises" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:TemplateField HeaderText="Código">
                                            <ItemTemplate>
                                                <asp:Label ID="lblCodigo" runat="server" Text='<%#Eval("Analise_Id")%>' />
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Análise">
                                            <ItemTemplate>
                                                <asp:Label ID="lblDescricao" runat="server" Text='<%#Eval("Descricao")%>' />
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Índice Mínimo">
                                            <ItemTemplate>
                                                <asp:Label ID="lblMinimo" runat="server" Text='<%#Eval("IndiceMinimo", "{0:N2}")%>' />
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Índice Máximo">
                                            <ItemTemplate>
                                                <asp:Label ID="lblMaximo" runat="server" Text='<%#Eval("IndiceMaximo", "{0:N2}")%>' />
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Índice">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtAnalise" CssClass="txtDecimal" runat="server" Columns="6" Text='<%#Eval("Indice", "{0:N2}") %>' />
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel2" HeaderText="TabPanel2">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="grdConsulta" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" PageSize="15" OnSelectedIndexChanged="grdConsulta_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Link" SelectText="<i class='fa fa-arrow-right seta'></i>" ShowSelectButton="True" />
                                    <asp:BoundField DataField="ReduzidoEmpresa" HeaderText="Empresa" />
                                    <asp:BoundField DataField="ReduzidoDeposito" HeaderText="Dep&#243;sito" />
                                    <asp:BoundField DataField="Movimento" DataFormatString="{0:dd/MM/yyyy HH:mm:ss}"
                                        HeaderText="Movimento" HtmlEncode="False" />
                                    <asp:BoundField DataField="Tipo" HeaderText="Tipo" />
                                    <asp:BoundField DataField="ValorProduto" HeaderText="Produto" />
                                    <asp:BoundField DataField="ValorProdutoDerivado" HeaderText="Produto Derivado" />
                                    <asp:BoundField DataField="Lote_id" HeaderText="Lote" />
                                    <asp:BoundField DataField="Classificacao_id" HeaderText="Class." />
                                    <asp:BoundField DataField="CentroDeCusto" HeaderText="Centro Custo" />
                                    <asp:BoundField DataField="Operacao" HeaderText="Oper.">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Embalagem" HeaderText="Embalagem" />
                                    <asp:BoundField DataField="Quantidade" DataFormatString="{0:N4}" HeaderText="Quantidade"
                                        HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EntradaSaida" HeaderText="E/S" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaOperacoes ID="ucConsultaOperacoes" runat="server" />
    <uc:ConsultaProduto ID="ucConsultaProduto" runat="server" />
    <uc:ConsultaDeLote ID="ucConsultaLote" runat="server" />
</asp:Content>
