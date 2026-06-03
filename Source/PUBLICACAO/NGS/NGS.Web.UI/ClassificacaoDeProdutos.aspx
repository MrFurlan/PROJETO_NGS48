<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ClassificacaoDeProdutos.aspx.vb" Inherits="NGS.Web.UI.ClassificacaoDeProdutos" %>

<asp:Content ID="scriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="bodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngClassificacaoDeProdutos" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="upnClassificacaoDeProdutos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Classificações de Produtos
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel1">
                    <HeaderTemplate>
                        Cadastro
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
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tabela:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTabela" runat="server" Width="555px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupo" runat="server" Width="555px" AutoPostBack="True"
                                    OnSelectedIndexChanged="DdlGrupo_SelectedIndexChanged" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProduto" runat="server" Width="555px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Análise:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlAnalise" runat="server" Width="555px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Seqüência:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblSequencia" runat="server" Font-Bold="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Faixa:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFaixaInicial" runat="server" MaxLength="6" CssClass="txtDecimal"
                                    Width="88px" />
                            </div>
                            <div class="collbl">
                                á:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFaixaFinal" runat="server" MaxLength="6" CssClass="txtDecimal"
                                    Width="88px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Índice:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtIndice" runat="server" MaxLength="6" CssClass="txtDecimal4" Width="88px" />
                            </div>
                            <div class="collbl">
                                Tolerância:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTolerancia" runat="server" MaxLength="6" CssClass="txtDecimal"
                                    Width="88px" />
                            </div>
                            <div class="collbl">
                                Faixa Válida:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="radSim" runat="server" Text="Sim" GroupName="grnFaixaValida" />
                                <asp:RadioButton ID="radNao" runat="server" Text="Não" GroupName="grnFaixaValida" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtObservacao" runat="server" Width="551px" MaxLength="50" TextMode="MultiLine" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="grdAbaCadastro" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="grdAbaCadastro_SelectedIndexChanged">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                        <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Tabela_Id">
                                        <HeaderStyle HorizontalAlign="Left" CssClass="none" />
                                        <ItemStyle CssClass="none" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Grupo">
                                        <HeaderStyle HorizontalAlign="Left" CssClass="none" />
                                        <ItemStyle CssClass="none" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Produto_Id">
                                        <HeaderStyle HorizontalAlign="Left" CssClass="none" />
                                        <ItemStyle CssClass="none" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Analise_Id">
                                        <HeaderStyle HorizontalAlign="Left" CssClass="none" />
                                        <ItemStyle CssClass="none" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoClassificacao" HeaderText="Tabela">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoProduto" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoAnalise" HeaderText="Análise">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Sequencia_Id" HeaderText="Seq.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="FaixaInicial" HeaderText="Faixa">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Indice" HeaderText="Índice">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Tolerancia" HeaderText="Tolerância">
                                        <HeaderStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="FaixaValida">
                                        <HeaderStyle HorizontalAlign="Left" CssClass="none" />
                                        <ItemStyle CssClass="none" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Observacoes">
                                        <HeaderStyle HorizontalAlign="Left" CssClass="none" />
                                        <ItemStyle CssClass="none" />
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
                <ajaxToolkit:TabPanel runat="server" ID="TabPanel2">
                    <HeaderTemplate>
                        Replicação
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="subtitulodiv">
                            Tabela Origem
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tabela:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTabelaOrigem" runat="server" Width="555px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnVoltarReplicar" runat="server" Text="Voltar" CssClass="botao"
                                    Width="80px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupoOrigem" runat="server" Width="555px" AutoPostBack="True" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnReplicar" runat="server" Text="Replicar" CssClass="botao" Width="80px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProdutoOrigem" runat="server" Width="555px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="subtitulodiv">
                            Tabela Destino
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tabela:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTabelaDestino" runat="server" Width="555px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupoDestino" runat="server" Width="555px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProdutoDestino" runat="server" Width="555px" AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt">
                                <asp:CheckBox ID="chkTabelaProdutoExiste" runat="server" Visible="False" Text="A Tabela/Produto de Destino já existe!, marque está opção se deseja regravá-la?" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
